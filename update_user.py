import json
import os
import pandas as pd
import requests
from msal import ConfidentialClientApplication

# Environment variables
TENANT_ID = os.getenv("TENANT_ID")
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
DOMAIN_NAME = "Ngalemo182.onmicrosoft.com"  # Replace with your actual domain

# Acquire Access Token
authority = f"https://login.microsoftonline.com/{TENANT_ID}"
app = ConfidentialClientApplication(
    CLIENT_ID, client_credential=CLIENT_SECRET, authority=authority
)
token_result = app.acquire_token_for_client(
    scopes=["https://graph.microsoft.com/.default"]
)

if "access_token" not in token_result:
    print(
        f"Failed to acquire token: {token_result.get('error_description')}"
    )
    exit(1)

headers = {
    "Authorization": f"Bearer {token_result['access_token']}",
    "Content-Type": "application/json",
}

# Read CSV and create users
df = pd.read_csv("sc300_practice_users_dallas-v2.csv")

for _, row in df.iterrows():
    # Construct required fields
    mail_nickname = (
        f"{str(row['FirstName']).lower()}.{str(row['LastName']).lower()}"
    )
    upn = f"{mail_nickname}@{DOMAIN_NAME}"

    user_payload = {
        "accountEnabled": True,
        "displayName": f"{row['FirstName']} {row['LastName']}",
        "mailNickname": mail_nickname,
        "userPrincipalName": upn,
        "passwordProfile": {
            "forceChangePasswordNextSignIn": True,
            "password": "TempPassword123!",  # Temporary password meeting complexity requirements
        },
    }

    # POST endpoint creates a new user object
    url = "https://graph.microsoft.com/v1.0/users"
    response = requests.post(url, headers=headers, json=user_payload)

    if response.status_code == 201:
        print(f"Successfully created user: {upn}")
    elif response.status_code == 409:
        print(f"User already exists: {upn}")
    else:
        print(f"Failed to create {upn}: {response.status_code} - {response.text}")