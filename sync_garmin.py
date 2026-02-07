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

def main():
    print("🚀 Skript gestartet: Verbinde mit Garmin...")
    
    # Credentials laden
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
            print("✅ Login via Session erfolgreich")
        except Exception as e:
            print(f"⚠️ Session failed: {e}")

    if not garmin or not garmin.garth.profile:
        print("🔑 Nutze Standard-Login...")
        garmin = Garmin(garmin_email, garmin_password)
        garmin.login()

    # --- FIX: USER-ID INITIALISIEREN ---
    print("👤 Validiere Benutzerprofil...")
    try:
        # get_full_name() triggert das Laden der Profil-Daten inklusive User-ID
        full_name = garmin.get_full_name()
        print(f"✅ Angemeldet als: {full_name}")
    except Exception as e:
        print(f"⚠️ Profil-Check fehlgeschlagen: {e}")

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
        
        activities = garmin.get_activities(0, 20)
        new_workouts = 0

        for act in reversed(activities):
            full_start_time = act.get('startTimeLocal', '')
            if full_start_time in existing_workouts:
                continue

            if " " in full_start_time:
                date_part, time_part = full_start_time.split(" ")
            else:
                date_part, time_part = full_start_time, ""

            gct_raw = act.get('avgGroundContactBalance', 0)
            gct_display = f"{round(gct_raw, 1)}% L / {round(100 - gct_raw, 1)}% R" if gct_raw and 0 < gct_raw < 100 else "-"

            workout_row = [
                date_part, time_part,
                act.get('activityType', {}).get('typeKey', ''),
                act.get('activityName', ''),
                round(act.get('distance', 0) / 1000, 2),
                act.get('calories', 0),
                round(act.get('duration', 0) / 60, 2),
                act.get('averageHR', 0), act.get('maxHR', 0),
                act.get('aerobicTrainingEffect', 0),
                act.get('averageRunningCadenceInStepsPerMinute', act.get('averageBikeCadence', 0)) or 0,
                act.get('maxRunningCadenceInStepsPerMinute', act.get('maxBikeCadence', 0)) or 0,
                round(act.get('averageSpeed', 0) * 3.6, 2),
                round(act.get('maxSpeed', 0) * 3.6, 2),
                round(act.get('elevationGain', 0), 1), round(act.get('elevationLoss', 0), 1),
                round(act.get('avgStrideLength', 0) / 100, 2) if act.get('avgStrideLength') else 0,
                gct_display,
                round(act.get('avgGroundContactTime', 0), 0) if act.get('avgGroundContactTime') else 0,
                round(act.get('avgVerticalOscillation', 0), 1) if act.get('avgVerticalOscillation') else 0,
                act.get('avgGradeAdjustedSpeed', 0), act.get('avgPower', 0), act.get('maxPower', 0),
                act.get('trainingStressScore', 0), act.get('steps', 0),
                act.get('totalReps', 0), act.get('totalPoses', 0),
                act.get('bodyBatteryDrainValue', 0), act.get('minTemperature', 0), act.get('maxTemperature', 0),
                act.get('averageRespirationRate', 0),
                round(act.get('movingDuration', 0) / 60, 2), round(act.get('elapsedDuration', 0) / 60, 2),
                round(act.get('minElevation', 0), 1), round(act.get('maxElevation', 0), 1)
            ]
            
            workout_sheet.append_row(workout_row)
            print(f"✅ Workout hinzugefügt: {date_part} - {act.get('activityName')}")
            new_workouts += 1
    except Exception as e:
        print(f"❌ Fehler bei Workouts: {e}")

    # --- TEIL 2: HEALTH DATA ---
    print("🩺 Synchronisiere Health-Daten (letzte 3 Tage)...")
    try:
        health_sheet = spreadsheet.worksheet("health_data")
        health_values = health_sheet.get_all_values()
        date_map = {row[0]: i + 1 for i, row in enumerate(health_values) if row}
        
        for i in range(3):
            date_obj = datetime.now() - timedelta(days=i)
            date_str = date_obj.strftime("%Y-%m-%d")
            print(f"🔍 Verarbeite Health für: {date_str}")
            
            try:
                # get_user_summary ist die stabilste Methode für Daily Stats
                stats = garmin.get_user_summary(date_str)
                
                try:
                    sleep = garmin.get_sleep_data(date_str)
                    sleep_score = sleep.get('dailySleepDTO', {}).get('sleepScore', '-')
                except:
                    sleep_score = "-"
                
                try:
                    rhr_data = garmin.get_rhr_and_hrv_data(date_str)
                    hrv_avg = rhr_data.get('hrvSummary', {}).get('lastNightAvg', '-')
                    rhr = rhr_data.get('restingHeartRate', '-')
                except:
                    hrv_avg = "-"
                    rhr = "-"

                health_row = [
                    date_str,
                    sleep_score,
                    hrv_avg,
                    rhr,
                    stats.get('bodyBatteryMostRecentValue', '-'),
                    stats.get('averageStressLevel', '-'),
                    stats.get('steps', 0)
                ]
                
                if date_str in date_map:
                    row_idx = date_map[date_str]
                    health_sheet.update(f"A{row_idx}:G{row_idx}", [health_row])
                    print(f"🔄 Health Update: {date_str}")
                else:
                    health_sheet.append_row(health_row)
                    print(f"✅ Health Neu: {date_str}")
                    
            except Exception as e:
                print(f"⚠️ Keine Health-Daten für {date_str} verfügbar: {e}")
    except Exception as e:
        print(f"❌ Fehler bei Health-Tab Zugriff: {e}")

    print("🏁
