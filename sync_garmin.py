import os
import json
import base64
import zipfile
import io
import garth
from garminconnect import Garmin
from google.oauth2.service_account import Credentials
import gspread

def main():
    print("Starting AI-Coach optimized Garmin sync (Split Date/Time)...")
    
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
            print("✅ Login via Session successful")
        except Exception as e:
            print(f"⚠️ Session failed: {e}")

    if not garmin or not garmin.garth.profile:
        garmin = Garmin(garmin_email, garmin_password)
        garmin.login()

    activities = garmin.get_activities(0, 30)
    
    # --- GOOGLE SHEETS ---
    creds_dict = json.loads(google_creds_json)
    creds = Credentials.from_service_account_info(creds_dict, scopes=['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive'])
    client = gspread.authorize(creds)
    sheet = client.open_by_key(sheet_id).sheet1
    all_rows = sheet.get_all_values()
    
    # Header mit getrenntem Date und Time
    headers = [
        "Date", "Time", "Activity Type", "Title", "Distance (km)", "Calories", "Duration (min)", 
        "Avg HR", "Max HR", "Aerobic TE", "Avg Cadence", "Max Cadence", 
        "Avg Speed (km/h)", "Max Speed (km/h)", "Total Ascent", "Total Descent", 
        "Avg Stride Length (m)", "Avg GCT Balance", "Avg GCT (ms)", "Avg Vert. Osc. (cm)", 
        "Avg GAP", "Avg Power", "Max Power", "Training Stress Score", "Steps", 
        "Total Reps", "Total Poses", "Body Battery Drain", "Min Temp", "Max Temp", 
        "Avg Resp", "Moving Time", "Elapsed Time", "Min Elevation", "Max Elevation"
    ]

    header_exists = False
    if all_rows and len(all_rows) > 0 and len(all_rows[0]) > 0:
        if all_rows[0][0] == "Date":
            header_exists = True

    if not header_exists:
        sheet.insert_row(headers, 1)
        all_rows = sheet.get_all_values()
    
    # Wir prüfen nun beide Spalten (Date + Time) für den Duplikat-Check
    existing_entries = {f"{row[0]} {row[1]}" for row in all_rows if len(row) > 1}

    new_entries = 0
    for act in reversed(activities):
        full_start_time = act.get('startTimeLocal', '') # Format: "2026-02-06 10:15:55"
