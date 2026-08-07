import os
import pandas as pd
import requests
from msal import ConfidentialClientApplication

print("=== STARTING ENTRA ID USER CREATION WORKFLOW (POST /v1.0/users) ===")

TENANT_ID = os.getenv("TENANT_ID")
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
TENANT_DOMAIN = "Ngalemo182.onmicrosoft.com"

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
csv_file = "sc300_practice_users_dallas-v2.csv"
df = pd.read_csv(csv_file)

# Clean header whitespace
df.columns = df.columns.str.strip()

for _, row in df.iterrows():
    given_name = str(row["GivenName"]).strip()
    surname = str(row["Surname"]).strip()
    
    # Use MailNickname from CSV if present, otherwise build from names
    if "MailNickname" in row and pd.notna(row["MailNickname"]):
        mail_nickname = str(row["MailNickname"]).strip()
    else:
        mail_nickname = f"{given_name.lower()}.{surname.lower()}"

    upn = f"{mail_nickname}@{TENANT_DOMAIN}"

    # Build user payload utilizing fields available in your CSV
    payload = {
        "accountEnabled": True,
        "displayName": str(row.get("DisplayName", f"{given_name} {surname}")).strip(),
        "givenName": given_name,
        "surname": surname,
        "mailNickname": mail_nickname,
        "userPrincipalName": upn,
        "jobTitle": str(row.get("JobTitle", "")).strip() if pd.notna(row.get("JobTitle")) else None,
        "department": str(row.get("Department", "")).strip() if pd.notna(row.get("Department")) else None,
        "city": str(row.get("City", "")).strip() if pd.notna(row.get("City")) else None,
        "state": str(row.get("State", "")).strip() if pd.notna(row.get("State")) else None,
        "country": str(row.get("Country", "")).strip() if pd.notna(row.get("Country")) else None,
        "passwordProfile": {
            "forceChangePasswordNextSignIn": True,
            "password": str(row.get("Password", "TempPassword123!")).strip() if pd.notna(row.get("Password")) else "TempPassword123!",
        },
    }

    # Remove keys with None values to avoid sending empty attributes
    payload = {k: v for k, v in payload.items() if v is not None}

    url = "https://graph.microsoft.com/v1.0/users"
    response = requests.post(url, headers=headers, json=payload)

    if response.status_code == 201:
        print(f"Successfully created user: {upn}")
    elif response.status_code == 409:
        print(f"User already exists: {upn}")
    else:
        print(f"Failed to create {upn}: {response.status_code} - {response.text}")