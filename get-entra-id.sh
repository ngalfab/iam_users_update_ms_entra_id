#!/bin/bash

# Ensure Azure CLI is installed
if ! command -v az &> /dev/null; then
    echo "Error: Azure CLI ('az') is not installed. Please install it first."
    exit 1
fi

# Check if user is logged in
az account show &> /dev/null
if [ $? -ne 0 ]; then
    echo "Logging into Azure..."
    az login
fi

echo "=========================================="
echo "      Microsoft Entra ID Credentials      "
echo "=========================================="

# 1. Retrieve Current Tenant ID
TENANT_ID=$(az account show --query tenantId -o tsv)
echo "Tenant ID: $TENANT_ID"

# Prompt for App Registration Name (or use default)
DEFAULT_APP_NAME="SC300-Entra-Automation"
read -p "Enter App Registration Name [$DEFAULT_APP_NAME]: " APP_NAME
APP_NAME=${APP_NAME:-$DEFAULT_APP_NAME}

# 2. Query Client ID for the given App Registration
CLIENT_ID=$(az ad app list --display-name "$APP_NAME" --query "[0].appId" -o tsv)

if [ -z "$CLIENT_ID" ] || [ "$CLIENT_ID" == "null" ]; then
    echo ""
    echo "App Registration '$APP_NAME' not found."
    read -p "Would you like to create it now? (y/n): " CREATE_APP
    if [[ "$CREATE_APP" =~ ^[Yy]$ ]]; then
        CLIENT_ID=$(az ad app create --display-name "$APP_NAME" --query appId -o tsv)
        echo "Successfully created '$APP_NAME'!"
    else
        echo "Exiting."
        exit 0
    fi
fi

echo "Client ID: $CLIENT_ID"
echo "=========================================="
echo ""
echo "Output for GitHub Actions Variables:"
echo "TENANT_ID  = $TENANT_ID"
echo "CLIENT_ID  = $CLIENT_ID"