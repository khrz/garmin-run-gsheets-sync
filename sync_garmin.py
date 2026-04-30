import os
import json
import sys
import requests
from datetime import datetime, timedelta
from google.oauth2.service_account import Credentials
import gspread

# --- HILFSFUNKTIONEN ---
def safe_num(val, default=0):
    return float(val) if val is not None else default

def main():
    print("🚀 Skript gestartet (Intervals.icu Edition)...")
    
    intervals_id = os.environ.get('INTERVALS_ID')
    intervals_api_key = os.environ.get('INTERVALS_API_KEY')
    google_creds_json = os.environ.get('GOOGLE_CREDENTIALS')
    sheet_id = os.environ.get('SHEET_ID')

    if not all([intervals_id, intervals_api_key, google_creds_json, sheet_id]):
        print("❌ Fehler: Umgebungsvariablen fehlen.")
        sys.exit(1)

    # --- GOOGLE SHEETS SETUP ---
    creds_dict = json.loads(google_creds_json)
    creds = Credentials.from_service_account_info(creds_dict, scopes=['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive'])
    client = gspread.authorize(creds)
    spreadsheet = client.open_by_key(sheet_id)

    # --- INTERVALS.ICU API SETUP ---
    auth = requests.auth.HTTPBasicAuth('API_KEY', intervals_api_key)
    base_url = f"https://intervals.icu/api/v1/athlete/{intervals_id}"
    
    now = datetime.now()
    oldest_workout = (now - timedelta(days=45)).strftime("%Y-%m-%dT00:00:00")
    newest = now.strftime("%Y-%m-%dT23:59:59")
    oldest_health = (now - timedelta(days=7)).strftime("%Y-%m-%d")

    # --- TEIL 1: WORKOUTS ---
    print("🏃 Synchronisiere Workouts...")
    try:
        workout_sheet = spreadsheet.worksheet("workout_database")
        all_rows = workout_sheet.get_all_values()
        existing_workouts = {f"{row[0]} {row[1]}" for row in all_rows if len(row) > 1}
        
        act_url = f"{base_url}/activities?oldest={oldest_workout}&newest={newest}"
        response = requests.get(act_url, auth=auth)
        response.raise_for_status()
        activities = response.json()
        
        activities = sorted(activities, key=lambda x: x.get('start_date_local', ''))
        
        for act in activities:
            start_local = act.get('start_date_local', '')
            if not start_local: continue
            
            date_part, time_part = start_local.split("T")
            
            if f"{date_part} {time_part}" not in existing_workouts:
                
                # --- SONDERFELDER AUSLESEN ---
                # 1. GCT Balance (Links/Rechts)
                lr_bal = act.get('avg_lr_balance')
                gct_str = f"{round(lr_bal, 1)}% L / {round(100 - lr_bal, 1)}% R" if lr_bal else "-"
                
                # 2. Power-Werte (Intervals hat mehrere Felder dafür)
                avg_pwr = safe_num(act.get('icu_average_watts') or act.get('device_watts') or act.get('average_watts'))
                max_pwr = safe_num(act.get('icu_pm_p_max') or act.get('p_max') or act.get('max_watts'))
                
                row = [
                    date_part, 
                    time_part, 
                    act.get('type') or '', 
                    act.get('name') or '',
                    round(safe_num(act.get('distance')) / 1000, 2),
                    safe_num(act.get('calories')), 
                    round(safe_num(act.get('moving_time')) / 60, 2),
                    safe_num(act.get('average_heartrate')), 
                    safe_num(act.get('max_heartrate')), 
                    0, # aerobicTE
                    safe_num(act.get('average_cadence')), 
                    safe_num(act.get('max_cadence')),
                    round(safe_num(act.get('average_speed')) * 3.6, 2),
                    round(safe_num(act.get('max_speed')) * 3.6, 2),
                    safe_num(act.get('total_elevation_gain')), 
                    safe_num(act.get('total_elevation_loss')),
                    safe_num(act.get('average_stride_length')),
                    gct_str, # ✅ GCT Balance eingefügt
                    0, # gct
                    0, # vertOsc
                    0, # gradeAdjustedSpeed
                    avg_pwr, # ✅ Avg Power eingefügt
                    max_pwr, # ✅ Max Power eingefügt
                    safe_num(act.get('icu_tss')), 
                    0, # steps in activity
                    0, # totalReps
                    0, # poses
                    0, # bodyBatteryDrain
                    safe_num(act.get('min_temp')), 
                    safe_num(act.get('max_temp')),
                    0, # avgResp
                    round(safe_num(act.get('moving_time')) / 60, 2),
                    round(safe_num(act.get('elapsed_time')) / 60, 2), 
                    safe_num(act.get('min_altitude')), 
                    safe_num(act.get('max_altitude'))
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
            
            bb_max = day.get('bodyBatteryHighest', "-")
            stress = day.get('stress', "-")
            steps = day.get('steps', 0)
            vo2_max = day.get('vo2max', "-")
            acute_load = round(day.get('atl', 0)) if day.get('atl') else "-"

            health_row = [date_str, sleep_score, sleep_duration, hrv_avg, rhr, bb_max, stress, steps, vo2_max, acute_load]
            
            if date_str in date_map:
                row_idx = date_map[date_str]
                health_sheet.update(f"A{row_idx}:J{row_idx}", [health_row])
                print(f"📊 {date_str}: Update (Sleep {sleep_score})")
            else:
                health_sheet.append_row(health_row)
                print(f"📊 {date_str}: Neu (Sleep {sleep_score})")

    except Exception as e:
        print(f"❌ Health-Fehler: {e}")

    print("🏁 Fertig")

if __name__ == "__main__":
    main()
