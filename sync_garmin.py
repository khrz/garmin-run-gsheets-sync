import os
import json
import base64
import zipfile
import io
import garth
from datetime import datetime, timedelta
from garminconnect import Garmin
from google.oauth2.service_account import Credentials
import gspread

# --- HILFSFUNKTION: Rekursive Suche (Deep Search) ---
def find_value_by_key(data, target_key):
    """Sucht rekursiv nach einem Key in verschachtelten Strukturen"""
    if not data: return None
    
    # Direkte Suche
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

def main():
    print("🚀 Skript gestartet...")
    
    garmin_email = os.environ.get('GARMIN_EMAIL')
    garmin_password = os.environ.get('GARMIN_PASSWORD')
    google_creds_json = os.environ.get('GOOGLE_CREDENTIALS')
    sheet_id = os.environ.get('SHEET_ID')
    session_base64 = os.environ.get('GARMIN_SESSION_BASE64')
    
    # --- GARMIN LOGIN ---
    garmin = None
    if session_base64:
        try:
            decoded_data = base64.b64decode(session_base64)
            with zipfile.ZipFile(io.BytesIO(decoded_data)) as z:
                z.extractall('session_data')
            garth.resume('session_data')
            garmin = Garmin()
            garmin.garth = garth.client
            garmin.display_name = garmin.garth.profile.get("displayName")
            print(f"✅ Login via Session erfolgreich: {garmin.display_name}")
        except Exception:
            print("⚠️ Session fehlgeschlagen, nutze Passwort...")

    if not garmin or not garmin.display_name:
        garmin = Garmin(garmin_email, garmin_password)
        garmin.login()
        garmin.display_name = garmin.get_display_name()
        print(f"✅ Login via Passwort: {garmin.display_name}")

    # --- GOOGLE SHEETS SETUP ---
    creds_dict = json.loads(google_creds_json)
    creds = Credentials.from_service_account_info(
        creds_dict, 
        scopes=['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
    )
    client = gspread.authorize(creds)
    spreadsheet = client.open_by_key(sheet_id)

    # --- TEIL 1: WORKOUTS ---
    print("🏃 Synchronisiere Workouts...")
    try:
        workout_sheet = spreadsheet.worksheet("workout_database")
        all_rows = workout_sheet.get_all_values()
        existing_workouts = {f"{row[0]} {row[1]}" for row in all_rows if len(row) > 1}
        
        activities = garmin.get_activities(0, 10)
        
        for act in reversed(activities):
            start = act.get('startTimeLocal', '')
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

    # --- TEIL 2: HEALTH DATABASE (DEEP SEARCH VO2 & LOAD) ---
    print("🩺 Synchronisiere Health-Datenbank...")
    try:
        health_sheet = spreadsheet.worksheet("health_data")
        health_values = health_sheet.get_all_values()
        date_map = {row[0]: i + 1 for i, row in enumerate(health_values) if row}
        
        for i in range(7):
            date_obj = datetime.now() - timedelta(days=i)
            date_str = date_obj.strftime("%Y-%m-%d")
            
            try:
                # 1. Datenquellen laden
                stats = garmin.get_user_summary(date_str)
                sleep_data = garmin.get_sleep_data(date_str)
                train_status = None
                try:
                    train_status = garmin.get_training_status(date_str)
                except: pass

                # 2. VO2max Suche (Priorität: Running > Cycling > Generic)
                vo2_max = stats.get('vo2MaxRunning')
                if not vo2_max: vo2_max = stats.get('vo2MaxCycling')
                if not vo2_max: vo2_max = stats.get('vo2Max') # Generisch
                if not vo2_max and train_status:
                    vo2_max = find_value_by_key(train_status, 'vo2Max') # Deep Search in TrainingStatus
                
                if not vo2_max: vo2_max = "-"

                # 3. ACUTE LOAD Suche (Priorität: acuteLoad > loadCombined > 7DayLoad)
                acute_load = "-"
                if train_status:
                    acute_load = train_status.get('acuteLoad')
                    if not acute_load: acute_load = train_status.get('loadCombined')
                    if not acute_load: acute_load = find_value_by_key(train_status, 'acuteLoad')
                    if not acute_load: acute_load = find_value_by_key(train_status, 'chronicLoad') # Fallback

                # 4. Sleep & Score
                sleep_score = find_value_by_key(sleep_data, 'sleepScore')
                if not sleep_score: sleep_score = find_value_by_key(stats, 'sleepScore')
                if not sleep_score: sleep_score = "-"
                
                dto = sleep_data.get('dailySleepDTO', {})
                seconds = dto.get('sleepTimeSeconds', 0)
                sleep_duration = round(seconds / 3600, 2) if seconds > 0 else "-"

                # 5. HRV & Rest
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
                
                # Update oder Neu
                if date_str in date_map:
                    row_idx = date_map[date_str]
                    health_sheet.update(f"A{row_idx}:J{row_idx}", [health_row])
                    print(f"📊 {date_str}: Sleep {sleep_score}, VO2 {vo2_max}, Load {acute_load}")
                else:
                    health_sheet.append_row(health_row)
                    print(f"📊 {date_str}: Neu angelegt (VO2: {vo2_max}, Load: {acute_load})")
                    
            except Exception as e:
                print(f"⚠️ Fehler am {date_str}: {e}")
    except Exception as e:
        print(f"❌ Health-Fehler: {e}")

    print("🏁 Fertig")

if __name__ == "__main__":
    main()
