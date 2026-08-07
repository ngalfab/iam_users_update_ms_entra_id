import os
import pandas as pd
import requests
from msal import ConfidentialClientApplication

# Debug marker to verify script version in GitHub Actions logs
print("=== STARTING ENTRA ID USER CREATION WORKFLOW (POST /v1.0/users) ===")

TENANT_ID = os.getenv("TENANT_ID")
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
TENANT_DOMAIN = "Ngalemo182.onmicrosoft.com"

# Verify environment variables
if not all([TENANT_ID, CLIENT_ID, CLIENT_SECRET]):
    print("ERROR: Missing TENANT_ID, CLIENT_ID, or CLIENT_SECRET environment variables.")
    exit(1)

# Authenticate with MSAL
authority = f"https://login.microsoftonline.com/{TENANT_ID}"
app = ConfidentialClientApplication(
    CLIENT_ID, client_credential=CLIENT_SECRET, authority=authority
)
token_result = app.acquire_token_for_client(
    scopes=["https://graph.microsoft.com/.default"]
)

if "access_token" not in token_result:
    print(f"ERROR: Token acquisition failed - {token_result.get('error_description')}")
    exit(1)

headers = {
    "Authorization": f"Bearer {token_result['access_token']}",
    "Content-Type": "application/json",
}

# Load user list CSV
df = pd.read_csv("sc300_practice_users_dallas-v2.csv")

for _, row in df.iterrows():
    first_name = str(row["FirstName"]).strip()
    last_name = str(row["LastName"]).strip()
    mail_nickname = f"{first_name.lower()}.{last_name.lower()}"
    upn = f"{mail_nickname}@{TENANT_DOMAIN}"

    payload = {
        "accountEnabled": True,
        "displayName": f"{first_name} {last_name}",
        "mailNickname": mail_nickname,
        "userPrincipalName": upn,
        "passwordProfile": {
            "forceChangePasswordNextSignIn": True,
            "password": "TempPassword123!",
        },
    }

    # Base endpoint for user creation
    url = "https://graph.microsoft.com/v1.0/users"
    response = requests.post(url, headers=headers, json=payload)

    if response.status_code == 201:
        print(f"Successfully created user: {upn}")
    elif response.status_code == 409:
        print(f"User already exists: {upn}")
    else:
        print(f"Failed to create {upn}: {response.status_code} - {response.text}")