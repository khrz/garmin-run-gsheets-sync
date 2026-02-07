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
    print("Skript gestartet...")
    
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
            print(f"Login via Session erfolgreich: {garmin.display_name}")
        except Exception:
            print("Session Login fehlgeschlagen, nutze Passwort...")

    if not garmin or not garmin.display_name:
        garmin = Garmin(garmin_email, garmin_password)
        garmin.login()
        garmin.display_name = garmin.get_display_name()
        print(f"Login via Passwort erfolgreich: {garmin.display_name}")

    # --- GOOGLE SHEETS SETUP ---
    creds_dict = json.loads(google_creds_json)
    creds = Credentials.from_service_account_info(
        creds_dict, 
        scopes=['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
    )
    client = gspread.authorize(creds)
    spreadsheet = client.open_by_key(sheet_id)

    # --- TEIL 1: WORKOUTS ---
    print("Synchronisiere Workouts...")
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
            print(f"✅ Workout hinzugefügt: {date_part}")
            new_workouts += 1
    except Exception as e:
        print(f"Fehler bei Workouts: {e}")

    # --- TEIL 2: HEALTH DATABASE ---
    print("Synchronisiere Health-Datenbank...")
    try:
        health_sheet = spreadsheet.worksheet("health_data")
        health_values = health_sheet.get_all_values()
        date_map = {row[0]: i + 1 for i, row in enumerate(health_values) if row}
        
        for i in range(7):
            date_obj = datetime.now() - timedelta(days=i)
            date_str = date_obj.strftime("%Y-%m-%d")
            
            try:
                stats = garmin.get_user_summary(date_str)
                
                # Schlaf-Daten & Dauer & Score
                sleep_score = "-"
                sleep_duration = "-"
                try:
                    sleep_data = garmin.get_sleep_data(date_str)
                    dto = sleep_data.get('dailySleepDTO', {})
                    
                    # Schlafdauer
                    seconds = dto.get('sleepTimeSeconds', 0)
                    if seconds > 0:
                        sleep_duration = round(seconds / 3600, 2)
                    
                    # Sleep Score (Versuch A und B)
                    sleep_score = dto.get('sleepScore', "-")
                    if sleep_score == "-":
                        sleep_score = sleep_data.get('sleepScore', "-")
                except: pass

                # HRV
                hrv_avg = "-"
                try:
                    hrv_data = garmin.get_hrv_data(date_str)
                    hrv_avg = hrv_data.get('hrvSummary', {}).get('lastNightAvg', "-")
                except:
                    try:
                        rhr_data = garmin.get_rhr_and_hrv_data(date_str)
                        hrv_avg = rhr_data.get('hrvSummary', {}).get('lastNightAvg', "-")
                    except: pass

                # Basiswerte
                rhr = stats.get('restingHeartRate', "-")
                bb_max = stats.get('bodyBatteryHighestValue', stats.get('bodyBatteryMostRecentValue', "-"))
                stress = stats.get('averageStressLevel', "-")
                steps = stats.get('totalSteps', stats.get('steps', 0))

                health_row = [date_str, sleep_score, sleep_duration, hrv_avg, rhr, bb_max, stress, steps]
                
                if date_str in date_map:
                    row_idx = date_map[date_str]
                    health_sheet.update(f"A{row_idx}:H{row_idx}", [health_row])
                    print(f"📊 {date_str}: Update (Score: {sleep_score}, Dur: {sleep_duration}h)")
                else:
                    health_sheet.append_row(health_row)
                    print(f"📊 {date_str}: Neu angelegt")
                    
            except Exception as e:
                print(f"Keine Daten für {date_str}: {e}")
    except Exception as e:
        print(f"Fehler bei Health: {e}")

    print("Skript beendet")

if __name__ == "__main__":
    main()
