import os
import sys
import pandas as pd
import requests
from msal import ConfidentialClientApplication

# ==========================================
# 1. CONSTANTS & ENVIRONMENT CONFIGURATION
# ==========================================
TENANT_ID = os.getenv("TENANT_ID")
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
TENANT_DOMAIN = "Ngalemo182.onmicrosoft.com"
GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"

# Guard check for missing environment secrets
if not all([TENANT_ID, CLIENT_ID, CLIENT_SECRET]):
    print("ERROR: Missing required environment variables (TENANT_ID, CLIENT_ID, CLIENT_SECRET).")
    sys.exit(1)


# ==========================================
# 2. AUTHENTICATION FUNCTION
# ==========================================
def get_graph_access_token():
    """Acquires an app-only access token for Microsoft Graph API using MSAL."""
    authority = f"https://login.microsoftonline.com/{TENANT_ID}"
    app = ConfidentialClientApplication(
        CLIENT_ID, client_credential=CLIENT_SECRET, authority=authority
    )
    
    token_result = app.acquire_token_for_client(
        scopes=["https://graph.microsoft.com/.default"]
    )
    
    if "access_token" in token_result:
        return token_result["access_token"]
    else:
        error_msg = token_result.get("error_description", "Unknown error")
        raise ConnectionError(f"Token acquisition failed: {error_msg}")


# ==========================================
# 3. HELPER FUNCTIONS
# ==========================================
def make_graph_request(method, endpoint, headers, payload=None):
    """Executes HTTP requests against the Microsoft Graph API."""
    url = f"{GRAPH_BASE_URL}/{endpoint}"
    response = requests.request(method=method, url=url, headers=headers, json=payload)
    return response


# ==========================================
# 4. MAIN EXECUTION PIPELINE
# ==========================================
def main():
    # Step A: Authenticate
    try:
        access_token = get_graph_access_token()
        print("Successfully authenticated with Microsoft Entra ID.")
    except Exception as e:
        print(f"Authentication Error: {e}")
        sys.exit(1)

    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json",
    }

    # Step B: Load & Process Input Data
    csv_path = "sc300_practice_users_dallas-v2.csv"
    if not os.path.exists(csv_path):
        print(f"ERROR: Input file '{csv_path}' not found.")
        sys.exit(1)

    df = pd.read_csv(csv_path)
    df.columns = df.columns.str.strip()

    # Step C: Operations Loop
    for _, row in df.iterrows():
        given_name = str(row["GivenName"]).strip()
        surname = str(row["Surname"]).strip()
        
        # Use MailNickname from CSV if present, otherwise construct from first/last name
        if "MailNickname" in row and pd.notna(row["MailNickname"]):
            mail_nickname = str(row["MailNickname"]).strip()
        else:
            mail_nickname = f"{given_name.lower()}.{surname.lower()}"

        upn = f"{mail_nickname}@{TENANT_DOMAIN}"

        # Custom payload mapping CSV attributes directly
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

        # Filter out keys with None values to avoid sending null attributes to Microsoft Graph
        payload = {k: v for k, v in payload.items() if v is not None}

        # Step D: Call Microsoft Graph Endpoint
        response = make_graph_request("POST", "users", headers, payload)

        # Step E: Handle Response States
        if response.status_code == 201:
            print(f"[201 CREATED] User: {upn}")
        elif response.status_code == 409:
            print(f"[409 EXISTS] User already exists: {upn}")
        else:
            print(f"[{response.status_code} FAILED] {upn}: {response.text}")


if __name__ == "__main__":
    main()