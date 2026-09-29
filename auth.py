"""
OAuth 2.0 Desktop Client Authentication Module for Google Sheets / Drive.
Implements the local browser login flow using InstalledAppFlow and caches credentials in token.json.
"""

import os
import glob
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow

from google.oauth2 import service_account
import json

# Google API Scopes required for reading & writing to Google Sheets
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.file",
    "https://www.googleapis.com/auth/drive.readonly",
]

TOKEN_FILE = "token.json"


def find_service_account_file():
    """Locate any Service Account JSON file in 'json/' or root directory."""
    candidates = glob.glob("json/*.json") + glob.glob("*.json")
    for path in candidates:
        if "token" in os.path.basename(path).lower():
            continue
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
                if isinstance(data, dict) and data.get("type") == "service_account":
                    return path
        except Exception:
            continue
    return None


def find_client_secret_file():
    """Locate the OAuth 2.0 Desktop Client JSON file."""
    # Check json/ folder first
    candidates = glob.glob("json/client_secret*.json") + glob.glob("json/*.json")
    # Exclude token files and service account files
    clean_candidates = []
    for c in candidates:
        bname = os.path.basename(c).lower()
        if "token" in bname:
            continue
        try:
            with open(c, "r", encoding="utf-8") as f:
                d = json.load(f)
                if isinstance(d, dict) and d.get("type") == "service_account":
                    continue
        except Exception:
            pass
        clean_candidates.append(c)

    if clean_candidates:
        return clean_candidates[0]
    
    # Check root directory
    root_candidates = glob.glob("client_secret*.json")
    if root_candidates:
        return root_candidates[0]

    # Check Streamlit Cloud Secrets
    try:
        import streamlit as st
        if hasattr(st, "secrets") and ("gcp_service_account" in st.secrets or "google_oauth_token" in st.secrets):
            return "Streamlit Secrets"
    except Exception:
        pass
        
    return None


def get_auth_type():
    """Returns 'service_account', 'oauth', or None based on active configuration."""
    try:
        import streamlit as st
        if hasattr(st, "secrets") and "gcp_service_account" in st.secrets:
            return "service_account"
        if hasattr(st, "secrets") and "google_oauth_token" in st.secrets:
            return "oauth"
    except Exception:
        pass

    if find_service_account_file():
        return "service_account"
    if is_authenticated():
        return "oauth"
    return None


def create_auth_flow(redirect_uri="http://localhost:8080/"):
    """Creates a configured InstalledAppFlow instance."""
    secret_file = find_client_secret_file()
    if not secret_file:
        raise FileNotFoundError(
            "OAuth 2.0 Client Secret JSON file not found in 'json/' or root directory. "
            "Please place your client_secret_*.json inside the 'json' folder."
        )
    return InstalledAppFlow.from_client_secrets_file(secret_file, SCOPES, redirect_uri=redirect_uri)


def get_authorization_url(redirect_uri="http://localhost:8080/"):
    """
    Generates a direct Google OAuth authorization URL for the user to click.
    Returns:
        tuple: (auth_url, state, flow)
    """
    flow = create_auth_flow(redirect_uri=redirect_uri)
    url, state = flow.authorization_url(prompt="consent", access_type="offline")
    return url, state, flow


import urllib.parse


def complete_auth_with_code(code_or_url, flow=None, redirect_uri="http://localhost:8080/"):
    """
    Exchanges an authorization code or redirect URL for Google credentials and saves token.json.
    Handles both direct code and full redirect URL safely without state mismatch errors.
    """
    if flow is None:
        flow = create_auth_flow(redirect_uri=redirect_uri)
    
    clean_val = code_or_url.strip()
    
    # Try extracting the pure code parameter first to bypass OAuth state verification issues
    code_param = None
    if "code=" in clean_val:
        parsed = urllib.parse.urlparse(clean_val)
        qs = urllib.parse.parse_qs(parsed.query)
        if "code" in qs and qs["code"]:
            code_param = qs["code"][0]

    try:
        if code_param:
            flow.fetch_token(code=code_param)
        elif "http" in clean_val:
            flow.fetch_token(authorization_response=clean_val)
        else:
            flow.fetch_token(code=clean_val)
    except Exception:
        # Fallback: create fresh flow and fetch with code directly
        if code_param:
            fresh_flow = create_auth_flow(redirect_uri=redirect_uri)
            fresh_flow.fetch_token(code=code_param)
            flow = fresh_flow
        else:
            raise

    creds = flow.credentials
    with open(TOKEN_FILE, "w", encoding="utf-8") as token:
        token.write(creds.to_json())
    return creds


def get_google_credentials():
    """
    Load credentials from:
    1. Service Account JSON file (headless 24/7 background sync), OR
    2. Cached user credentials from token.json, OR
    3. Start the local browser OAuth flow.
    Returns:
        google.oauth2.credentials.Credentials or service_account.Credentials.
    """
    # 1. Prefer Service Account JSON if present (ideal for 24/7 mobile sync)
    sa_path = find_service_account_file()
    if sa_path:
        try:
            return service_account.Credentials.from_service_account_file(sa_path, scopes=SCOPES)
        except Exception as e:
            print(f"Error loading service account credentials from {sa_path}: {e}")

    # 2. Fall back to User OAuth token.json
    creds = None
    if os.path.exists(TOKEN_FILE):
        try:
            creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
        except Exception as e:
            print(f"Error loading token.json: {e}")
            creds = None

    # If there are no valid credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
            except Exception as e:
                print(f"Token refresh failed: {e}. Starting fresh auth flow.")
                creds = None

        if not creds:
            flow = create_auth_flow(redirect_uri="http://localhost:8080/")
            # Run local server on port 8080
            creds = flow.run_local_server(port=8080, prompt="consent", access_type="offline")

        # Save credentials for subsequent executions
        with open(TOKEN_FILE, "w", encoding="utf-8") as token:
            token.write(creds.to_json())

    return creds


def is_authenticated():
    """Check if valid service account or refreshable OAuth credentials currently exist."""
    if find_service_account_file():
        return True
    if not os.path.exists(TOKEN_FILE):
        return False
    try:
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
        return creds and (creds.valid or creds.refresh_token is not None)
    except Exception:
        return False


def logout():
    """Removes token.json to reset the authentication session."""
    if os.path.exists(TOKEN_FILE):
        try:
            os.remove(TOKEN_FILE)
            return True
        except Exception:
            return False
    return False
