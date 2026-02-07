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

    # CRITICAL FIX: Profil laden, um die User-ID für Health-Anfragen zu registrieren
    print("👤 Lade Benutzerprofil zur Identifikation...")
    try:
        garmin.display_name = garmin.get_display_name()
        print(f"✅
