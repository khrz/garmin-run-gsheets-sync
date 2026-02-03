import os
import json
import base64
import zipfile
import io
import garth
from garminconnect import Garmin
from google.oauth2.service_account import Credentials
import gspread
from datetime import datetime, timedelta

# Load environment variables
if os.path.exists('.env'):
    try:
        from dotenv import load_dotenv
        load_dotenv()
    except ImportError:
        pass

def format_duration(seconds):
    """Convert seconds to minutes (rounded to 2 decimals)"""
    return round(seconds / 60, 2) if seconds else 0

def format_pace(distance_meters, duration_seconds):
    """Calculate pace in min/km"""
    if not distance_meters or not duration_seconds:
        return 0
    distance_km = distance_meters / 1000
    pace_seconds = duration_seconds / distance_km
    return round(pace_seconds / 60, 2)

def main():
    print("Starting Garmin running activities sync...")
    
    # Get credentials from environment variables
    garmin_email = os.environ.get('GARMIN_EMAIL')
    garmin_password = os.environ.get('GARMIN_PASSWORD')
    google_creds_json = os.environ.get('GOOGLE_CREDENTIALS')
    sheet_id = os.environ.get('SHEET_ID')
    session_base64 = os.environ.get('GARMIN_SESSION_BASE64') # Neues Secret
    
    if not all([google_creds_json, sheet_id]):
        print("❌ Missing required Google environment variables")
        return

    # --- GARMIN LOGIN START ---
    print("Connecting to Garmin...")
    garmin = None

    # Versuch 1: Login via Session-Token (Umgeht 2FA)
    if session_base64:
        try:
            print("Attempting login via Session-Token...")
            decoded_data = base64.b64decode(session_base64)
            with zipfile.ZipFile(io.BytesIO(decoded_data)) as z:
                z.extractall('session_data')
            
            garth.resume('session_data')
            garmin = Garmin()
            garmin.garth = garth.client
            print("✅ Login via Session successful")
        except Exception as e:
            print(f"⚠️ Session login failed: {e}")

    # Versuch 2: Fallback auf Passwort (falls Session fehlt oder abgelaufen ist)
    if not garmin or not garmin.garth.profile:
        print("Attempting standard login with password...")
        try:
            garmin = Garmin(garmin_email, garmin_password)
            garmin.login()
            print("✅ Standard login successful")
        except Exception as e:
            print(f"❌ All login attempts failed: {e}")
            return
    # --- GARMIN LOGIN ENDE ---
    
    # Get recent activities
    print("Fetching recent activities...")
    try:
        activities = garmin.get_activities(0, 20)
        print(f"Found {len(activities)} total activities")
    except Exception as e:
        print(f"❌ Failed to fetch activities: {e}")
        return
    
    # Filter for running activities only
    running_activities = [
        activity for activity in activities 
        if activity.get('activityType', {}).get('typeKey', '').lower() in ['running', 'treadmill_running', 'trail_running']
    ]
    
    print(f"Found {len(running_activities)} running activities")
    
    if not running_activities:
        print("No running activities found in recent data")
        return
    
    # Connect to Google Sheets
    print("Connecting to Google Sheets...")
    try:
        creds_dict = json.loads(google_creds_json)
        creds = Credentials.from_service_account_info(
            creds_dict,
            scopes=[
                'https://www.googleapis.com/auth/spreadsheets',
                'https://www.googleapis.com/auth/drive'
            ]
        )
        client = gspread.authorize(creds)
        sheet = client.open_by_key(sheet_id).sheet1 # Nutzt jetzt die SHEET_ID direkt
        print("✅ Connected to Google Sheets")
    except Exception as e:
        print(f"❌ Failed to connect to Google Sheets: {e}")
        return
    
    # Get existing dates to avoid duplicates
    try:
        existing_data = sheet.get_all_values()
        existing_dates = set()
        if len(existing_data) > 1:
            for row in existing_data[1:]:
                if row and row[0]:
                    existing_dates.add(row[0])
        print(f"Found {len(existing_dates)} existing entries")
    except Exception as e:
        print(f"Warning: Could not check existing data: {e}")
        existing_dates = set()
    
    # Process each running activity
    new_entries = 0
    for activity in running_activities:
        try:
            activity_date = activity.get('startTimeLocal', '')[:10]
            
            if activity_date in existing_dates:
                print(f"Skipping {activity_date} - already exists")
                continue
            
            activity_name = activity.get('activityName', 'Run')
            distance_meters = activity.get('distance', 0)
            distance_km = round(distance_meters / 1000, 2) if distance_meters else 0
            duration_seconds = activity.get('duration', 0)
            duration_min = format_duration(duration_seconds)
            avg_pace = format_pace(distance_meters, duration_seconds)
            avg_hr = activity.get('averageHR', 0) or 0
            max_hr = activity.get('maxHR', 0) or 0
            calories = activity.get('calories', 0) or 0
            avg_cadence = activity.get('averageRunningCadenceInStepsPerMinute', 0) or 0
            elevation_gain = round(activity.get('elevationGain', 0), 1) if activity.get('elevationGain') else 0
            activity_type = activity.get('activityType', {}).get('typeKey', 'running')
            
            row = [
                activity_date, activity_name, distance_km, duration_min,
                avg_pace, avg_hr, max_hr, calories, avg_cadence,
                elevation_gain, activity_type
            ]
            
            sheet.append_row(row)
            print(f"✅ Added: {activity_date} - {activity_name} ({distance_km} km)")
            new_entries += 1
            
        except Exception as e:
            print(f"❌ Error processing activity: {e}")
            continue
    
    if new_entries > 0:
        print(f"\n🎉 Successfully added {new_entries} new running activities!")
    else:
        print("\n✓ No new activities to add")

if __name__ == "__main__":
    main()
