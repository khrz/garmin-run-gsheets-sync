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
    print("🚀 Skript gestartet...")
    
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
            garmin.display_name = garmin.garth.profile.get("displayName")
            print(f"✅ Login via Session: {garmin.display_name}")
        except Exception as e:
            print(f"⚠️ Session failed: {e}")

    if not garmin or not garmin.display_name:
        garmin = Garmin(garmin_email, garmin_password)
        garmin.login()
        garmin.display_name = garmin.get_display_name()
        print("✅ Login via Password")

    # --- GOOGLE SHEETS SETUP ---
    creds_dict = json.loads(google_creds_json)
    creds = Credentials.from_service_account_info(creds_dict, scopes=['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive'])
    client = gspread.authorize(creds)
    spreadsheet = client.open_by_key(sheet_id)

    # --- TEIL 1: WORKOUTS (Bleibt wie gehabt) ---
    # ... (Dein Workout-Code)

    # --- TEIL 2: HEALTH DATABASE ---
    print("🩺 Synchronisiere Health-Datenbank...")
    try:
        health_sheet = spreadsheet.worksheet("health_data")
        health_values = health_sheet.get_all_values()
        date_map = {row[0]: i + 1 for i, row in enumerate(health_values) if row}
        
        # ÄNDERE DIESE ZAHL FÜR DEN BACKFILL (z.B. 30 für den ersten Run, danach wieder 7)
        DAYS_TO_FETCH = 30 
        
        for i in range(DAYS_TO_FETCH):
            date_obj = datetime.now() - timedelta(days=i)
            date_str = date_obj.strftime("%Y-%m-%d")
            
            try:
                # 1. User Summary
                stats = garmin.get_user_summary(date_str)
                
                # 2. Schlafdaten mit tieferer Suche
                sleep_score = "-"
                try:
                    sleep_data = garmin.get_sleep_data(date_str)
                    sleep_score = sleep_data.get('dailySleepDTO', {}).get('sleepScore', "-")
                except: pass

                # 3. HRV & RHR
                hrv_avg = "-"
                rhr = stats.get('restingHeartRate', "-")
                try:
                    hrv_data = garmin.get_rhr_and_hrv_data(date_str)
                    hrv_avg = hrv_data.get('hrvSummary', {}).get('lastNightAvg', "-")
                    if rhr == "-": rhr = hrv_data.get('restingHeartRate', "-")
                except: pass

                health_row = [
                    date_str,
                    sleep_score,
                    hrv_avg,
                    rhr,
                    stats.get('bodyBatteryHighestValue', stats.get('bodyBatteryMostRecentValue', "-")),
                    stats.get('averageStressLevel', "-"),
                    stats.get('totalSteps', stats.get('steps', 0))
                ]
                
                if date_str in date_map:
                    row_idx = date_map[date_str]
                    health_sheet.update(f"A{row_idx}:G{row_idx}", [health_row])
                else:
                    health_sheet.append_row(health_row)
                
                print(f"📊 {date_str}: Sleep {sleep_score}, HRV {hrv_avg}")
                    
            except Exception as e:
                print(f"⚠️ {date_str} übersprungen: {e}")
    except Exception as e:
        print(f"❌ Fehler: {e}")

    print("🏁 Skript beendet")

if __name__ == "__main__":
    main()
