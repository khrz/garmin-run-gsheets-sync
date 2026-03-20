import os
import json
import base64
import zipfile
import io
import garth
import shutil
from datetime import datetime, timedelta
from garminconnect import Garmin
from google.oauth2.service_account import Credentials
import gspread

# --- HILFSFUNKTION: Rekursive Suche nach Keys ---
def find_value_by_key(data, target_key):
    if not data: return None
    if isinstance(data, dict):
        if target_key in data and data[target_key] is not None:
            return data[target_key]
        for k, v in data.items():
            if isinstance(v, (dict, list)):
                found = find_value_by_key(v, target_key)
                if found: return found
    elif isinstance(data, list):
        for item in data:
            found = find_value_by_key(item, target_key)
            if found: return found
    return None

# --- HILFSFUNKTION: Spezifische Sleep-Score Suche ---
def get_sleep_score_explicit(data):
    if not data: return None
    # 1. Struktur B (Fenix 7 / Epix / Forerunner neu)
    try:
        val = data.get('dailySleepDTO', {}).get('sleepScores', {}).get('overall', {}).get('value')
        if val: return val
    except: pass
    
    # 2. Struktur A (Ältere API)
    try:
        val = data.get('dailySleepDTO', {}).get('sleepScore')
        if val: return val
    except: pass
    
    # 3. Fallback: Suche nach Key "sleepScore" überall
    return find_value_by_key(data, 'sleepScore')

def main():
    print("🚀 Skript gestartet...")
    
    garmin_email = os.environ.get('GARMIN_EMAIL')
    garmin_password = os.environ.get('GARMIN_PASSWORD')
    google_creds_json = os.environ.get('GOOGLE_CREDENTIALS')
    sheet_id = os.environ.get('SHEET_ID')
    session_base64 = os.environ.get('GARMIN_SESSION_BASE64')

  # --- GARMIN LOGIN ---
    garmin_tokens_hex = os.environ.get('GARMIN_TOKENS')
    garmin = None
    
    if garmin_tokens_hex:
        try:
            import binascii, base64, re
            # 1. Nur echte Hex-Zeichen behalten
            h_clean = re.sub(r'[^0-9a-fA-F]', '', garmin_tokens_hex)
            
            # 2. In Base64-Text umwandeln
            b64_str = binascii.unhexlify(h_clean).decode('utf-8', errors='ignore')
            
            # 3. RADIKAL-REINIGUNG: Nur A-Z, 0-9, +, / behalten (kein =, kein \n)
            b64_pure = re.sub(r'[^A-Za-z0-9+/]', '', b64_str)
            
            # 4. Mathematische Korrektur auf Vielfaches von 4
            # Das tötet den "2785" Fehler endgültig
            valid_len = (len(b64_pure) // 4) * 4
            b64_final = b64_pure[:valid_len]
            
            # 5. Padding manuell hinzufügen (sauberer Weg)
            b64_final += "=" * ((4 - len(b64_final) % 4) % 4)
            
            # 6. Dekodieren und Laden
            json_text = base64.b64decode(b64_final).decode('utf-8')
            garmin = Garmin(garmin_email, garmin_password)
            garmin.garth.loads(json_text)
            
            if not garmin.garth.username:
                garmin.login()
            print(f"✅ Login erfolgreich: {garmin.get_display_name()}")
        except Exception as e:
            print(f"⚠️ Token-Login fehlgeschlagen: {e}")
            garmin = None

    if not garmin:
        print("Fallback auf Passwort-Login...")
        garmin = Garmin(garmin_email, garmin_password)
        garmin.login()

    # --- GOOGLE SHEETS SETUP ---
    creds_dict = json.loads(google_creds_json)
    creds = Credentials.from_service_account_info(creds_dict, scopes=['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive'])
    client = gspread.authorize(creds)
    spreadsheet = client.open_by_key(sheet_id)

    # --- TEIL 1: WORKOUTS & VO2MAX ACTIVITY SCAN ---
    print("🏃 Synchronisiere Workouts...")
    activity_vo2_map = {}
    
    try:
        workout_sheet = spreadsheet.worksheet("workout_database")
        all_rows = workout_sheet.get_all_values()
        existing_workouts = {f"{row[0]} {row[1]}" for row in all_rows if len(row) > 1}
        
        activities = garmin.get_activities(0, 20)
        
        for act in reversed(activities):
            start = act.get('startTimeLocal', '')
            date_part = start.split(" ")[0] if " " in start else start
            
            # VO2max aus Aktivität retten
            if 'vo2MaxValue' in act and act['vo2MaxValue']:
                activity_vo2_map[date_part] = round(act['vo2MaxValue'])
            
            if start not in existing_workouts:
                d, t = start.split(" ") if " " in start else (start, "")
                gct_raw = act.get('avgGroundContactBalance', 0)
                gct = f"{round(gct_raw, 1)}% L / {round(100 - gct_raw, 1)}% R" if 0 < gct_raw < 100 else "-"
                
                row = [
                    d, t, act.get('activityType', {}).get('typeKey', ''), act.get('activityName', ''),
                    round(act.get('distance', 0)/1000, 2), act.get('calories', 0), round(act.get('duration', 0)/60, 2),
                    act.get('averageHR', 0), act.get('maxHR', 0), act.get('aerobicTrainingEffect', 0),
                    act.get('averageRunningCadenceInStepsPerMinute', 0), act.get('maxRunningCadenceInStepsPerMinute', 0),
                    round(act.get('averageSpeed', 0)*3.6, 2), round(act.get('maxSpeed', 0)*3.6, 2),
                    act.get('elevationGain', 0), act.get('elevationLoss', 0), round(act.get('avgStrideLength', 0)/100, 2),
                    gct, round(act.get('avgGroundContactTime', 0), 0), round(act.get('avgVerticalOscillation', 0), 1),
                    act.get('avgGradeAdjustedSpeed', 0), act.get('avgPower', 0), act.get('maxPower', 0),
                    act.get('trainingStressScore', 0), act.get('steps', 0), act.get('totalReps', 0), act.get('totalPoses', 0),
                    act.get('bodyBatteryDrainValue', 0), act.get('minTemperature', 0), act.get('maxTemperature', 0),
                    act.get('averageRespirationRate', 0), round(act.get('movingDuration', 0)/60, 2),
                    round(act.get('elapsedDuration', 0)/60, 2), act.get('minElevation', 0), act.get('maxElevation', 0)
                ]
                workout_sheet.append_row(row)
                print(f"✅ Workout: {d}")
    except Exception as e: print(f"❌ Workout-Fehler: {e}")

    # --- TEIL 2: HEALTH DATABASE ---
    print("🩺 Synchronisiere Health-Datenbank...")
    try:
        health_sheet = spreadsheet.worksheet("health_data")
        health_values = health_sheet.get_all_values()
        date_map = {row[0]: i + 1 for i, row in enumerate(health_values) if row}
        
        for i in range(7):
            date_obj = datetime.now() - timedelta(days=i)
            date_str = date_obj.strftime("%Y-%m-%d")
            
            try:
                stats = garmin.get_user_summary(date_str)
                sleep_data = garmin.get_sleep_data(date_str)
                
                # Training Status holen
                train_status = None
                try: train_status = garmin.get_training_status(date_str)
                except: pass

                # 1. ACUTE LOAD
                acute_load = "-"
                if train_status:
                    val = find_value_by_key(train_status, 'dailyTrainingLoadAcute')
                    if val: acute_load = round(val)
                    if acute_load == "-":
                        val = find_value_by_key(train_status, 'acuteLoad')
                        if val: acute_load = round(val)

                # 2. VO2MAX
                vo2_max = "-"
                if train_status:
                    v = train_status.get('mostRecentVO2Max', {}).get('generic', {}).get('vo2MaxValue')
                    if v: vo2_max = round(v)
                if vo2_max == "-":
                    v = stats.get('vo2MaxRunning')
                    if v: vo2_max = round(v)
                if vo2_max == "-" and date_str in activity_vo2_map:
                    vo2_max = activity_vo2_map[date_str]

                # 3. SLEEP SCORE
                sleep_score = get_sleep_score_explicit(sleep_data)
                if not sleep_score: 
                    sleep_score = get_sleep_score_explicit(stats)
                if not sleep_score: sleep_score = "-"
                
                dto = sleep_data.get('dailySleepDTO', {})
                seconds = dto.get('sleepTimeSeconds', 0)
                sleep_duration = round(seconds / 3600, 2) if seconds > 0 else "-"

                # 4. HRV & Rest
                hrv_avg = "-"
                try:
                    hrv_info = garmin.get_hrv_data(date_str)
                    hrv_avg = hrv_info.get('hrvSummary', {}).get('lastNightAvg', "-")
                except: pass
                
                rhr = stats.get('restingHeartRate', "-")
                bb_max = stats.get('bodyBatteryHighestValue', stats.get('bodyBatteryMostRecentValue', "-"))
                stress = stats.get('averageStressLevel', "-")
                steps = stats.get('totalSteps', stats.get('steps', 0))

                health_row = [date_str, sleep_score, sleep_duration, hrv_avg, rhr, bb_max, stress, steps, vo2_max, acute_load]
                
                if date_str in date_map:
                    row_idx = date_map[date_str]
                    health_sheet.update(f"A{row_idx}:J{row_idx}", [health_row])
                    print(f"📊 {date_str}: Sleep {sleep_score}, Load {acute_load}, VO2 {vo2_max}")
                else:
                    health_sheet.append_row(health_row)
                    print(f"📊 {date_str}: Neu (Sleep {sleep_score})")
                    
            except Exception as e:
                print(f"⚠️ Fehler am {date_str}: {e}")
    except Exception as e:
        print(f"❌ Health-Fehler: {e}")

    print("🏁 Fertig")

if __name__ == "__main__":
    main()
