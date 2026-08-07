import os
import pandas as pd
import requests
import msal

# Load credentials from GitHub Actions environment variables
TENANT_ID = os.environ.get("TENANT_ID")
CLIENT_ID = os.environ.get("CLIENT_ID")
CLIENT_SECRET = os.environ.get("CLIENT_SECRET")

# Authenticate with MSAL
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPE = ["https://graph.microsoft.com/.default"]

app = msal.ConfidentialClientApplication(
    CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET
)
token_result = app.acquire_token_for_client(scopes=SCOPE)

if "access_token" not in token_result:
    raise Exception(f"Failed to authenticate: {token_result.get('error_description')}")

headers = {
    "Authorization": f"Bearer {token_result['access_token']}",
    "Content-Type": "application/json"
}

# Process CSV and execute API updates
df = pd.read_csv("sc300_practice_users_dallas-v2.csv")

for _, row in df.iterrows():
    upn = row["UserPrincipalName"]
    url = f"https://graph.microsoft.com/v1.0/users/{upn}"
    
    payload = {
        "jobTitle": str(row["JobTitle"]),
        "department": str(row["Department"]),
        "city": str(row["City"]),
        "state": str(row["State"]),
        "country": str(row["Country"]),
        "employeeId": str(row["EmployeeId"])
    }

    response = requests.patch(url, headers=headers, json=payload)

    if response.status_code == 204:
        print(f"Successfully updated: {upn}")
    else:
        print(f"Failed to update {upn}: {response.status_code} - {response.text}")