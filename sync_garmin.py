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
    print("Starting extended Garmin sync with auto-headers...")
    
    # Credentials laden
    garmin_email = os.environ.get('GARMIN_EMAIL')
    garmin_password = os.environ.get('GARMIN_PASSWORD')
    google_creds_json = os.environ.get('GOOGLE_CREDENTIALS')
    sheet_id = os.environ.get('SHEET_ID')
    session_base64 = os.environ.get('GARMIN_SESSION_BASE64')
    
    # Garmin Login (Session-Logik)
    print("Connecting to Garmin...")
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
        print("Attempting standard login...")
        garmin = Garmin(garmin_email, garmin_password)
        garmin.login()

    # Aktivitäten holen (letzte 20)
    activities = garmin.get_activities(0, 20)
    
    # Google Sheets Verbindung
    creds_dict = json.loads(google_creds_json)
    creds = Credentials.from_service_account_info(
        creds_dict, 
        scopes=['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
    )
    client = gspread.authorize(creds)
    sheet = client.open_by_key(sheet_id).sheet1

    # Vorhandene Daten abrufen
    existing_data = sheet.get_all_values()
    
    # Kopfzeilen definieren
    headers = [
        "Date", "Activity Type", "Title", "Distance (km)", "Calories", "Duration (min)", 
        "Avg HR", "Max HR", "Aerobic TE", "Avg Run Cadence", "Max Run Cadence", 
        "Avg Speed (km/h)", "Max Speed (km/h)", "Total Ascent", "Total Descent", 
        "Avg Stride Length (m)", "Vertical Ratio / VO2", "Avg Vertical Oscillation", 
        "Avg GCT", "Avg GAP", "Avg Power", "Max Power", "Steps", 
        "Body Battery Drain", "Min Temp", "Max Temp", "Moving Time", 
        "Elapsed Time", "Min Elevation", "Max Elevation"
    ]

    # Falls das Sheet leer ist: Header einfügen
    if not existing_data:
        sheet.insert_row(headers, 1)
        print("✅ Header-Zeile wurde neu erstellt.")
        existing_dates = set()
    else:
        # Bestehende Startzeiten sammeln, um Duplikate zu vermeiden
        existing_dates = {row[0] for row in existing_data if row}

    new_entries = 0
    # Wir drehen die Liste um (reversed), damit der älteste Lauf zuerst hinzugefügt wird
    for act in reversed(activities):
        full_start_time = act.get('startTimeLocal', '')
        
        # Check ob Eintrag schon existiert
        if full_start_time in existing_dates:
            continue

        # Daten-Extraktion
        row = [
            full_start_time,                                         # Date
            act.get('activityType', {}).get('typeKey', ''),          # Type
            act.get('activityName', ''),                             # Title
            round(act.get('distance', 0) / 1000, 2),                # Distance (km)
            act.get('calories', 0),                                  # Calories
            round(act.get('duration', 0) / 60, 2),                   # Duration (min)
            act.get('averageHR', 0),                                 # Avg HR
            act.get('maxHR', 0),                                     # Max HR
            act.get('aerobicTrainingEffect', 0),                     # Aerobic TE
            act.get('averageRunningCadenceInStepsPerMinute', 0),     # Avg Run Cadence
            act.get('maxRunningCadenceInStepsPerMinute', 0),         # Max Run Cadence
            round(act.get('averageSpeed', 0) * 3.6, 2),              # Avg Speed (km/h)
            round(act.get('maxSpeed', 0) * 3.6, 2),                  # Max Speed (km/h)
            round(act.get('elevationGain', 0), 1),                   # Total Ascent
            round(act.get('elevationLoss', 0), 1),                   # Total Descent
            round(act.get('avgStrideLength', 0) / 100, 2),           # Avg Stride Length (m)
            act.get('vO2MaxValue', 0),                               # Platzhalter
            act.get('avgVerticalOscillation', 0),                    # Avg Vertical Oscillation
            act.get('avgGroundContactTime', 0),                      # Avg GCT
            act.get('avgGradeAdjustedSpeed', 0),                     # Avg GAP
            act.get('avgPower', 0),                                  # Avg Power
            act.get('maxPower', 0),                                  # Max Power
            act.get('steps', 0),                                     # Steps
            act.get('bodyBatteryDrainValue', 0),                     # Body Battery Drain
            act.get('minTemperature', 0),                            # Min Temp
            act.get('maxTemperature', 0),                            # Max Temp
            round(act.get('movingDuration', 0) / 60, 2),             # Moving Time
            round(act.get('elapsedDuration', 0) / 60, 2),            # Elapsed Time
            round(act.get('minElevation', 0), 1),                    # Min Elevation
            round(act.get('maxElevation', 0), 1)                     # Max Elevation
        ]
        
        sheet.append_row(row)
        print(f"✅ Added: {full_start_time} - {act.get('activityName')}")
        new_entries += 1

    print(f"Fertig! {new_entries} neue Einträge hinzugefügt.")

if __name__ == "__main__":
    main()
