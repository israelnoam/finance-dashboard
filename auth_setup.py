"""
Interactive Google OAuth 2.0 Desktop Authentication Setup
Runs in foreground, launches browser on Windows, and saves token.json.
"""

import os
import sys
from auth import find_client_secret_file, SCOPES, TOKEN_FILE
from google_auth_oauthlib.flow import InstalledAppFlow

def main():
    print("=" * 60)
    print(" Google Account Setup for Financial Dashboard")
    print("=" * 60)

    secret_file = find_client_secret_file()
    if not secret_file:
        print("❌ Error: Client secret JSON not found in 'json/' directory.")
        sys.exit(1)

    print(f"Found client secret: {secret_file}")
    print("Starting authentication flow...")

    flow = InstalledAppFlow.from_client_secrets_file(
        secret_file,
        SCOPES,
        redirect_uri="http://localhost:8080/"
    )

    auth_url, state = flow.authorization_url(
        prompt="consent",
        access_type="offline"
    )

    print("\n👉 Please open the following URL in your browser to sign in:")
    print("-" * 60)
    print(auth_url)
    print("-" * 60)

    # Open in Windows default browser
    try:
        os.system(f'start "" "{auth_url}"')
    except Exception as e:
        print(f"Note: Could not automatically launch browser ({e}). Please copy and paste the URL above.")

    print("\nWaiting for you to complete authorization in your browser...")
    try:
        # Run local server on port 8080
        flow.run_local_server(port=8080, prompt="consent", access_type="offline", open_browser=False)
        creds = flow.credentials

        with open(TOKEN_FILE, "w", encoding="utf-8") as f:
            f.write(creds.to_json())

        print("\n" + "=" * 60)
        print("✅ SUCCESS! Authentication token saved to token.json.")
        print("Your Financial Dashboard is now connected to Google Sheets!")
        print("=" * 60)

    except Exception as exc:
        print(f"\n❌ Local server error: {exc}")
        print("\nAlternative: If your browser redirected to http://localhost:8080/?code=...")
        code_input = input("Paste the full redirect URL or authorization code here: ").strip()
        if code_input:
            try:
                if "code=" in code_input:
                    flow.fetch_token(authorization_response=code_input)
                else:
                    flow.fetch_token(code=code_input)
                with open(TOKEN_FILE, "w", encoding="utf-8") as f:
                    f.write(flow.credentials.to_json())
                print("✅ SUCCESS! Token saved from pasted code.")
            except Exception as ex2:
                print(f"❌ Failed to exchange code: {ex2}")

if __name__ == "__main__":
    main()
