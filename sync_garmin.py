import os
import json
import sys
import requests
from datetime import datetime, timedelta
from google.oauth2.service_account import Credentials
import gspread

def main():
    print("🚀 Skript gestartet (Intervals.icu Edition)...")
    
    intervals_id = os.environ.get('INTERVALS_ID')
    intervals_api_key = os.environ.get('INTERVALS_API_KEY')
    google_creds_json = os.environ.get('GOOGLE_CREDENTIALS')
    sheet_id = os.environ.get('SHEET_ID')

    if not all([intervals_id, intervals_api_key, google_creds_json, sheet_id]):
        print("❌ Fehler: Umgebungsvariablen fehlen (INTERVALS_ID, INTERVALS_API_KEY, GOOGLE_CREDENTIALS, SHEET_ID).")
        sys.exit(1)

    # --- GOOGLE SHEETS SETUP ---
    creds_dict = json.loads(google_creds_json)
    creds = Credentials.from_service_account_info(creds_dict, scopes=['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive'])
    client = gspread.authorize(creds)
    spreadsheet = client.open_by_key(sheet_id)

    # --- INTERVALS.ICU API SETUP ---
    # Login über Basic Auth: Nutzername ist immer 'API_KEY', das Passwort ist dein generierter Key.
    auth = requests.auth.HTTPBasicAuth('API_KEY', intervals_api_key)
    base_url = f"https://intervals.icu/api/v1/athlete/{intervals_id}"
    
    now = datetime.now()
    oldest_workout = (now - timedelta(days=21)).strftime("%Y-%m-%dT00:00:00")
    newest = now.strftime("%Y-%m-%dT23:59:59")
    oldest_health = (now - timedelta(days=7)).strftime("%Y-%m-%d")

    # --- TEIL 1: WORKOUTS ---
    print("🏃 Synchronisiere Workouts...")
    try:
        workout_sheet = spreadsheet.worksheet("workout_database")
        all_rows = workout_sheet.get_all_values()
        existing_workouts = {f"{row[0]} {row[1]}" for row in all_rows if len(row) > 1}
        
        # Aktivitäten von Intervals abrufen
        act_url = f"{base_url}/activities?oldest={oldest_workout}&newest={newest}"
        response = requests.get(act_url, auth=auth)
        response.raise_for_status()
        activities = response.json()
        
        # Chronologisch sortieren (älteste zuerst)
        activities = sorted(activities, key=lambda x: x.get('start_date_local', ''))
        
        for act in activities:
            start_local = act.get('start_date_local', '')
            if not start_local: continue
            
            date_part, time_part = start_local.split("T")
            
            if f"{date_part} {time_part}" not in existing_workouts:
                # Fallback für tiefe Garmin-Laufmetriken, die Intervals evtl. nicht mitspeichert
                row = [
                    date_part, 
                    time_part, 
                    act.get('type', ''), 
                    act.get('name', ''),
                    round(act.get('distance', 0) / 1000, 2),
                    act.get('calories', 0), 
                    round(act.get('moving_time', 0) / 60, 2),
                    act.get('average_heartrate', 0), 
                    act.get('max_heartrate', 0), 
                    0, # aerobicTE
                    act.get('average_cadence', 0), 
                    act.get('max_cadence', 0),
                    round(act.get('average_speed', 0) * 3.6, 2),
                    round(act.get('max_speed', 0) * 3.6, 2),
                    act.get('total_elevation_gain', 0), 
                    0, # elevationLoss
                    act.get('average_stride_length', 0),
                    "-", # gct balance
                    0, # gct
                    0, # vertOsc
                    0, # gradeAdjustedSpeed
                    act.get('average_watts', 0), 
                    act.get('max_watts', 0),
                    act.get('icu_tss', 0), # TSS wird direkt von Intervals berechnet
                    0, # steps in activity
                    0, # totalReps
                    0, # poses
                    0, # bodyBatteryDrain
                    0, # minTemp
                    0, # maxTemp
                    0, # avgResp
                    round(act.get('moving_time', 0) / 60, 2),
                    round(act.get('elapsed_time', 0) / 60, 2), 
                    0, # minEle
                    0  # maxEle
                ]
                workout_sheet.append_row(row)
                print(f"✅ Workout: {date_part} {time_part} ({act.get('name', '')})")
    except Exception as e:
        print(f"❌ Workout-Fehler: {e}")

    # --- TEIL 2: HEALTH DATABASE ---
    print("🩺 Synchronisiere Health-Datenbank...")
    try:
        health_sheet = spreadsheet.worksheet("health_data")
        health_values = health_sheet.get_all_values()
        date_map = {row[0]: i + 1 for i, row in enumerate(health_values) if row}
        
        # Wellness-Daten holen
        well_url = f"{base_url}/wellness?oldest={oldest_health}&newest={newest[:10]}"
        response = requests.get(well_url, auth=auth)
        response.raise_for_status()
        wellness_data = response.json()
        
        for day in wellness_data:
            date_str = day.get('id') 
            if not date_str: continue
            
            sleep_score = day.get('sleepScore', "-")
            sleep_secs = day.get('sleepSecs', 0)
            sleep_duration = round(sleep_secs / 3600, 2) if sleep_secs > 0 else "-"
            hrv_avg = day.get('hrv', "-")
            rhr = day.get('restingHR', "-")
            
            # Garmin-Metriken aus Intervals übernehmen
            bb_max = day.get('bodyBatteryHighest', "-")
            stress = day.get('stress', "-")
            steps = day.get('steps', 0)
            vo2_max = day.get('vo2max', "-")
            acute_load = round(day.get('atl', 0)) if day.get('atl') else "-"

            health_row = [date_str, sleep_score, sleep_duration, hrv_avg, rhr, bb_max, stress, steps, vo2_max, acute_load]
            
            if date_str in date_map:
                row_idx = date_map[date_str]
                health_sheet.update(f"A{row_idx}:J{row_idx}", [health_row])
                print(f"📊 {date_str}: Update (Sleep {sleep_score}, Load {acute_load})")
            else:
                health_sheet.append_row(health_row)
                print(f"📊 {date_str}: Neu (Sleep {sleep_score})")

    except Exception as e:
        print(f"❌ Health-Fehler: {e}")

    print("🏁 Fertig")

if __name__ == "__main__":
    main()
