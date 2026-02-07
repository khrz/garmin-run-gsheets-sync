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
            # Wir nehmen die ID aus dem Profil
            garmin.display_name = garmin.garth.profile.get("displayName")
            print(f"Login via Session erfolgreich: {garmin.display_name}")
        except Exception as e:
            print(f"Session failed: {e}")

    if not garmin or not garmin.display_name:
        print("Nutze Standard-Login...")
        garmin = Garmin(garmin_email, garmin_password)
        garmin.login()
        garmin.display_name = garmin.get_display_name()

    # --- GOOGLE SHEETS SETUP ---
    creds_dict = json.loads(google_creds_json)
    creds = Credentials.from_service_account_info(creds_dict, scopes=['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive'])
    client = gspread.authorize(creds)
    spreadsheet = client.open_by_key(sheet_id)

    # --- TEIL 1: WORKOUTS (Bleibt unverändert) ---
    # ... (Dein Workout-Code hier einfügen oder einfach so lassen, falls du ihn schon hast)

    # --- TEIL 2: HEALTH DATA (Optimiert) ---
    print("Synchronisiere Health-Daten...")
    try:
        health_sheet = spreadsheet.worksheet("health_data")
        health_values = health_sheet.get_all_values()
        date_map = {row[0]: i + 1 for i, row in enumerate(health_values) if row}
        
        for i in range(3):
            date_obj = datetime.now() - timedelta(days=i)
            date_str = date_obj.strftime("%Y-%m-%d")
            print(f"Verarbeite Health für: {date_str}")
            
            try:
                # 1. Basis-Stats (Schritte, RHR, Body Battery, Stress)
                stats = garmin.get_user_summary(date_str)
                
                # 2. Schlaf-Daten
                sleep_score = "-"
                try:
                    sleep_data = garmin.get_sleep_data(date_str)
                    sleep_score = sleep_data.get('dailySleepDTO', {}).get('sleepScore', "-")
                except: pass

                # 3. HRV-Daten
                hrv_avg = "-"
                try:
                    hrv_data = garmin.get_rhr_and_hrv_data(date_str)
                    hrv_avg = hrv_data.get('hrvSummary', {}).get('lastNightAvg', "-")
                except: pass

                # Extraktion der Basis-Werte aus stats
                # RHR steht oft direkt in den stats oder in rhr_data
                rhr = stats.get('restingHeartRate', "-")
                if rhr == "-" and 'rhr_data' in locals():
                    rhr = rhr_data.get('restingHeartRate', "-")

                bb_max = stats.get('bodyBatteryMostRecentValue', "-") # Aktueller Wert
                # Falls wir den echten Max-Wert wollen:
                if 'bodyBatteryHighestValue' in stats:
                    bb_max = stats['bodyBatteryHighestValue']

                stress_avg = stats.get('averageStressLevel', "-")
                steps = stats.get('totalSteps', stats.get('steps', 0))

                health_row = [
                    date_str,
                    sleep_score,
                    hrv_avg,
                    rhr,
                    bb_max,
                    stress_avg,
                    steps
                ]
                
                if date_str in date_map:
                    row_idx = date_map[date_str]
                    health_sheet.update(f"A{row_idx}:G{row_idx}", [health_row])
                    print(f"Update Health: {date_str} (Score: {sleep_score}, HRV: {hrv_avg})")
                else:
                    health_sheet.append_row(health_row)
                    print(f"Neu Health: {date_str}")
                    
            except Exception as e:
                print(f"Fehler am {date_str}: {e}")
    except Exception as e:
        print(f"Fehler bei Health-Tab: {e}")

    print("Skript beendet")

if __name__ == "__main__":
    main()
