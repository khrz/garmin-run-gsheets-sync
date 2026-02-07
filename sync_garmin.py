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
        except Exception:
            print("⚠️ Session failed, falling back to password")

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

    # --- TEIL 2: HEALTH DATABASE ---
    print("🩺 Synchronisiere Health-Datenbank...")
    try:
        health_sheet = spreadsheet.worksheet("health_data")
        health_values = health_sheet.get_all_values()
        date_map = {row[0]: i + 1 for i, row in enumerate(health_values) if row}
        
        for i in range(7): # Letzte 7 Tage für volle Historie
            date_obj = datetime.now() - timedelta(days=i)
            date_str = date_obj.strftime("%Y-%m-%d")
            
            try:
                stats = garmin.get_user_summary(date_str)
                
                # 1. Schlaf & Dauer
                sleep_score = "-"
                sleep_duration = "-"
                try:
                    sleep_data = garmin.get_sleep_data(date_str)
                    dto = sleep_data.get('dailySleepDTO', {})
                    sleep_score = dto.get('sleepScore', "-")
                    seconds = dto.get('sleepTimeSeconds', 0)
                    if seconds > 0:
                        sleep_duration = round(seconds / 3600, 2)
                except: pass

                # 2. HRV (Versuche zwei verschiedene Wege)
                hrv_avg = "-"
                try:
                    hrv_data = garmin.get_hrv_data(date_str)
                    hrv_avg = hrv_data.get('hrvSummary', {}).get('lastNightAvg', "-")
                except:
                    # Backup-Versuch über RHR-Endpunkt
                    try:
                        rhr_data = garmin.get_rhr_and_hrv_data(date_str)
                        hrv_avg = rhr_data.get('hrvSummary', {}).get('lastNightAvg', "-")
                    except: pass

                # 3. Restliche Metriken
                rhr = stats.get('restingHeartRate', "-")
                bb_max = stats.get('bodyBatteryHighestValue', stats.get('bodyBatteryMostRecentValue', "-"))
                stress = stats.get('averageStressLevel', "-")
                steps = stats.get('totalSteps', stats.get('steps', 0))

                health_row = [date_str, sleep_score, sleep_duration, hrv_avg, rhr, bb_max, stress, steps]
                
                if date_str in date_map:
                    row_idx = date_map[date_str]
                    health_sheet.update(f"A{row_idx}:H{row_idx}", [health_row])
                    print(f"📊 {date_str}: Sleep {sleep_score}, Dur {sleep_duration}h, HRV {hrv_avg}")
                else:
                    health_sheet.append_row(health_row)
                    print(f"📊 {date_str}: Neu angelegt")
                    
            except Exception as e:
                print(f"⚠️ {date_str} Fehler: {e}")
    except Exception as e:
        print(f"❌ Fehler: {e}")

    print("🏁 Fertig")

if __name__ == "__main__":
    main()
