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

    csv_path = "users.csv"
    if not os.path.exists(csv_path):
        print(f"ERROR: Input file '{csv_path}' not found.")
        sys.exit(1)

    df = pd.read_csv(csv_path)

    # Clean header spaces and UTF-8 BOM characters
    df.columns = df.columns.str.strip().str.replace("\ufeff", "")
    
    # Case-insensitive column map (maps lowercase header to actual CSV header)
    col_map = {col.lower().replace(" ", ""): col for col in df.columns}

    for idx, row in df.iterrows():
        # Helper function to query row values safely across multiple naming variations
        def get_val(*variations):
            for var in variations:
                key = var.lower().replace(" ", "")
                actual_col = col_map.get(key)
                if actual_col and pd.notna(row[actual_col]):
                    val = str(row[actual_col]).strip()
                    if val:
                        return val
            return None

        # Resolve primary attributes with fallbacks
        given_name = get_val("GivenName", "FirstName", "Given Name")
        surname = get_val("Surname", "LastName", "Last Name", "FamilyName")

        if not given_name or not surname:
            print(f"[SKIP Row {idx + 1}] Missing First/Given Name or Last/Surname.")
            continue

        mail_nickname = get_val("MailNickname", "Nickname") or f"{given_name.lower()}.{surname.lower()}"
        upn = f"{mail_nickname}@{TENANT_DOMAIN}"

        # Your exact payload mapping structure
        payload = {
            "accountEnabled": True,
            "displayName": get_val("DisplayName") or f"{given_name} {surname}",
            "givenName": given_name,
            "surname": surname,
            "mailNickname": mail_nickname,
            "userPrincipalName": upn,
            "jobTitle": get_val("JobTitle", "Title"),
            "department": get_val("Department"),
            "city": get_val("City"),
            "state": get_val("State"),
            "country": get_val("Country"),
            "passwordProfile": {
                "forceChangePasswordNextSignIn": True,
                "password": get_val("Password") or "TempPassword123!",
            },
        }

        # Remove keys with None values to keep payload clean for Graph API
        payload = {k: v for k, v in payload.items() if v is not None}

        # Send creation request
        response = make_graph_request("POST", "users", headers, payload)

        # Handle response status codes
        if response.status_code == 201:
            print(f"[201 CREATED] User: {upn}")
        elif response.status_code in (400, 409) and "ObjectConflict" in response.text:
            print(f"[EXISTS] User already exists in Entra ID: {upn}")
        else:
            print(f"[{response.status_code} FAILED] {upn}: {response.text}")


if __name__ == "__main__":
    main()