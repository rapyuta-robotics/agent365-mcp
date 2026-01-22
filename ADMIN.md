# Agent 365 MCP - IT Administrator Setup Guide

This guide covers the complete setup process for IT administrators to enable Agent 365 MCP for their organization.

## Prerequisites

- **Azure/Entra ID Global Administrator** or **Application Administrator** role
- **Azure CLI** installed (`az` command)
- **PowerShell** (for optional setup script)
- Users must have **Copilot for Microsoft 365** license assigned

## Overview

The setup involves:
1. Creating the Agent 365 service principal (Microsoft's backend)
2. Creating your organization's app registration
3. Configuring permissions
4. Granting admin consent
5. Distributing credentials to users

## Step 1: Create Agent 365 Service Principal

This creates the connection to Microsoft's Agent 365 backend service.

### Option A: Using PowerShell Script (Recommended)

```powershell
# Download the official setup script
Invoke-WebRequest -Uri "https://raw.githubusercontent.com/microsoft/Agent365-devTools/main/scripts/cli/Auth/New-Agent365ToolsServicePrincipalProdPublic.ps1" -OutFile "New-Agent365ToolsServicePrincipal.ps1"

# Run the script (requires admin consent)
./New-Agent365ToolsServicePrincipal.ps1
```

### Option B: Using Azure CLI

```bash
# Sign in to Azure
az login

# The Agent 365 API application ID (Microsoft-owned)
AGENT365_API_ID="ea9ffc3e-8a23-4a7d-836d-234d7c7565c1"

# Create service principal for Agent 365 in your tenant
az ad sp create --id $AGENT365_API_ID

# Verify it was created
az ad sp show --id $AGENT365_API_ID --query "{displayName:displayName, appId:appId, id:id}"
```

Expected output:
```json
{
  "displayName": "Agent365Tools",
  "appId": "ea9ffc3e-8a23-4a7d-836d-234d7c7565c1",
  "id": "<your-tenant-specific-id>"
}
```

## Step 2: Create Your App Registration

Create an app registration that your users will authenticate against.

```bash
# Create the app registration
az ad app create \
  --display-name "Agent 365 MCP" \
  --sign-in-audience "AzureADMyOrg"

# Get the app ID (save this - you'll need it)
APP_ID=$(az ad app list --display-name "Agent 365 MCP" --query "[0].appId" -o tsv)
echo "Your App ID: $APP_ID"

# Enable public client flows (required for device code auth)
az ad app update --id $APP_ID --is-fallback-public-client true

# Add localhost redirect URI for device code flow
az ad app update --id $APP_ID \
  --public-client-redirect-uris "http://localhost"

# Create service principal for the app
az ad sp create --id $APP_ID

# Get your Tenant ID (save this - users need it)
TENANT_ID=$(az account show --query tenantId -o tsv)
echo "Your Tenant ID: $TENANT_ID"
```

## Step 3: Add API Permissions

Add the required Agent 365 scopes to your app.

```bash
# Agent 365 API ID
AGENT365_API="ea9ffc3e-8a23-4a7d-836d-234d7c7565c1"

# Permission scope IDs (from Microsoft documentation)
# Teams scope
az ad app permission add --id $APP_ID --api $AGENT365_API \
  --api-permissions 5efd4b9c-e459-40d4-a524-35db033b072f=Scope

# Copilot scope
az ad app permission add --id $APP_ID --api $AGENT365_API \
  --api-permissions 342a1c02-975d-4f22-bae9-56a200ca9fd0=Scope

# Mail scope
az ad app permission add --id $APP_ID --api $AGENT365_API \
  --api-permissions be685e8e-277f-43ec-aff6-087fdca57ca3=Scope

# Calendar scope
az ad app permission add --id $APP_ID --api $AGENT365_API \
  --api-permissions 75c3a580-2c8f-4906-adc6-ffa8601d78dc=Scope

# Me (profile) scope
az ad app permission add --id $APP_ID --api $AGENT365_API \
  --api-permissions 2ce6ce0f-4701-4b11-8087-5031f87ad5b9=Scope

# OneDrive/SharePoint scope
az ad app permission add --id $APP_ID --api $AGENT365_API \
  --api-permissions 45b74cfc-7a12-4589-8d26-781de38fbfcc=Scope

# Word scope
az ad app permission add --id $APP_ID --api $AGENT365_API \
  --api-permissions 6f7b3c3c-d822-4164-b9ec-8bf520399d24=Scope

# Excel scope
az ad app permission add --id $APP_ID --api $AGENT365_API \
  --api-permissions c80101f0-b7e0-4a85-b8c2-0c855aeced97=Scope

# Files scope
az ad app permission add --id $APP_ID --api $AGENT365_API \
  --api-permissions 965326d7-c2b4-4ae5-9833-a7286013fb2c=Scope
```

### Permission Reference Table

| Scope Name | Scope ID | Description |
|------------|----------|-------------|
| McpServers.Teams.All | 5efd4b9c-e459-40d4-a524-35db033b072f | Teams chats, channels, meetings |
| McpServers.CopilotMCP.All | 342a1c02-975d-4f22-bae9-56a200ca9fd0 | M365 Copilot search |
| McpServers.Mail.All | be685e8e-277f-43ec-aff6-087fdca57ca3 | Outlook email |
| McpServers.Calendar.All | 75c3a580-2c8f-4906-adc6-ffa8601d78dc | Calendar events |
| McpServers.Me.All | 2ce6ce0f-4701-4b11-8087-5031f87ad5b9 | User profile |
| McpServers.OneDriveSharepoint.All | 45b74cfc-7a12-4589-8d26-781de38fbfcc | SharePoint/OneDrive |
| McpServers.Word.All | 6f7b3c3c-d822-4164-b9ec-8bf520399d24 | Word documents |
| McpServers.Excel.All | c80101f0-b7e0-4a85-b8c2-0c855aeced97 | Excel spreadsheets |
| McpServers.Files.All | 965326d7-c2b4-4ae5-9833-a7286013fb2c | File operations |

## Step 4: Grant Admin Consent

Grant organization-wide consent for the permissions.

### Option A: Using Azure Portal (Recommended)

1. Go to Azure Portal > Entra ID > App registrations
2. Find "Agent 365 MCP"
3. Go to "API permissions"
4. Click "Grant admin consent for [Your Organization]"
5. Confirm

### Option B: Using Azure CLI (Graph API)

```bash
# Get the service principal IDs
SP_ID=$(az ad sp list --filter "appId eq '$APP_ID'" --query "[0].id" -o tsv)
AGENT365_SP=$(az ad sp list --filter "appId eq 'ea9ffc3e-8a23-4a7d-836d-234d7c7565c1'" --query "[0].id" -o tsv)

echo "App Service Principal ID: $SP_ID"
echo "Agent 365 Service Principal ID: $AGENT365_SP"

# Create the permission grant for ALL users in the organization
az rest --method POST \
  --uri "https://graph.microsoft.com/v1.0/oauth2PermissionGrants" \
  --headers "Content-Type=application/json" \
  --body "{
    \"clientId\": \"$SP_ID\",
    \"consentType\": \"AllPrincipals\",
    \"resourceId\": \"$AGENT365_SP\",
    \"scope\": \"McpServers.Teams.All McpServers.Mail.All McpServers.Calendar.All McpServers.Me.All McpServers.CopilotMCP.All McpServers.OneDriveSharepoint.All McpServers.Word.All McpServers.Files.All McpServers.Excel.All\"
  }"
```

### Option C: Grant for Specific Users Only

If you want to limit access to specific users instead of the entire organization:

```bash
# For each user, create an individual grant
USER_ID="user@yourcompany.com"

az rest --method POST \
  --uri "https://graph.microsoft.com/v1.0/oauth2PermissionGrants" \
  --headers "Content-Type=application/json" \
  --body "{
    \"clientId\": \"$SP_ID\",
    \"consentType\": \"Principal\",
    \"principalId\": \"$(az ad user show --id $USER_ID --query id -o tsv)\",
    \"resourceId\": \"$AGENT365_SP\",
    \"scope\": \"McpServers.Teams.All McpServers.Mail.All McpServers.Calendar.All McpServers.Me.All McpServers.CopilotMCP.All McpServers.OneDriveSharepoint.All McpServers.Word.All McpServers.Files.All McpServers.Excel.All\"
  }"
```

## Step 5: Verify Setup

```bash
# List the permission grants
az rest --method GET \
  --uri "https://graph.microsoft.com/v1.0/oauth2PermissionGrants?\$filter=clientId eq '$SP_ID'"
```

## Step 6: Distribute to Users

Provide users with:

| Information | How to Get |
|-------------|------------|
| Tenant ID | `az account show --query tenantId -o tsv` |
| Client ID | The App ID from Step 2 |

### User Setup Instructions

Share this with your users:

```bash
# 1. Set environment variables
export AGENT365_TENANT_ID="<tenant-id-from-admin>"
export AGENT365_CLIENT_ID="<client-id-from-admin>"

# 2. Authenticate (one-time)
npx github:rapyuta-robotics/agent365-mcp auth

# 3. Add to your MCP client config
# See README.md for configuration examples
```

## Updating Permissions

If you need to add or modify permissions later:

```bash
# Get existing grant ID
GRANT_ID=$(az rest --method GET \
  --uri "https://graph.microsoft.com/v1.0/oauth2PermissionGrants?\$filter=clientId eq '$SP_ID'" \
  --query "value[0].id" -o tsv)

# Update with new scopes
az rest --method PATCH \
  --uri "https://graph.microsoft.com/v1.0/oauth2PermissionGrants/$GRANT_ID" \
  --headers "Content-Type=application/json" \
  --body "{
    \"scope\": \"McpServers.Teams.All McpServers.Mail.All McpServers.Calendar.All McpServers.Me.All McpServers.CopilotMCP.All McpServers.OneDriveSharepoint.All McpServers.Word.All McpServers.Files.All McpServers.Excel.All\"
  }"
```

## Revoking Access

### Revoke for All Users

```bash
# Delete the permission grant
az rest --method DELETE \
  --uri "https://graph.microsoft.com/v1.0/oauth2PermissionGrants/$GRANT_ID"
```

### Revoke for Specific User

```bash
# Find the user's grant
USER_ID="user@yourcompany.com"
USER_OBJ_ID=$(az ad user show --id $USER_ID --query id -o tsv)

USER_GRANT=$(az rest --method GET \
  --uri "https://graph.microsoft.com/v1.0/oauth2PermissionGrants?\$filter=clientId eq '$SP_ID' and principalId eq '$USER_OBJ_ID'" \
  --query "value[0].id" -o tsv)

# Delete it
az rest --method DELETE \
  --uri "https://graph.microsoft.com/v1.0/oauth2PermissionGrants/$USER_GRANT"
```

## Restricting Access to Specific Users

By default, the AllPrincipals consent allows **everyone in your org** to use Agent 365 MCP. To restrict access to specific users or groups:

### Option 1: Azure Portal

1. Go to **Entra ID** → **Enterprise applications**
2. Find **Agent 365 MCP**
3. Go to **Properties**
4. Set **Assignment required?** to **Yes**
5. Save
6. Go to **Users and groups**
7. Click **Add user/group**
8. Select the users or groups who should have access

### Option 2: Azure CLI

```bash
# Get the enterprise app (service principal) object ID
SP_OBJECT_ID=$(az ad sp list --filter "appId eq '$APP_ID'" --query "[0].id" -o tsv)

# Enable assignment requirement
az rest --method PATCH \
  --uri "https://graph.microsoft.com/v1.0/servicePrincipals/$SP_OBJECT_ID" \
  --headers "Content-Type=application/json" \
  --body '{"appRoleAssignmentRequired": true}'

# Add a user (get user object ID first)
USER_EMAIL="user@yourcompany.com"
USER_ID=$(az ad user show --id $USER_EMAIL --query id -o tsv)

# Assign user to app (empty appRoleId = default access)
az rest --method POST \
  --uri "https://graph.microsoft.com/v1.0/servicePrincipals/$SP_OBJECT_ID/appRoleAssignments" \
  --headers "Content-Type=application/json" \
  --body "{
    \"principalId\": \"$USER_ID\",
    \"resourceId\": \"$SP_OBJECT_ID\",
    \"appRoleId\": \"00000000-0000-0000-0000-000000000000\"
  }"
```

### Option 3: Add a Security Group

```bash
# Create or use existing security group
GROUP_NAME="Agent365-Users"
GROUP_ID=$(az ad group show --group $GROUP_NAME --query id -o tsv)

# Assign group to app
az rest --method POST \
  --uri "https://graph.microsoft.com/v1.0/servicePrincipals/$SP_OBJECT_ID/appRoleAssignments" \
  --headers "Content-Type=application/json" \
  --body "{
    \"principalId\": \"$GROUP_ID\",
    \"resourceId\": \"$SP_OBJECT_ID\",
    \"appRoleId\": \"00000000-0000-0000-0000-000000000000\"
  }"
```

Users not assigned will see: `AADSTS50105: Your administrator has configured the application to block users`.

## Troubleshooting

### "AADSTS65001: User hasn't consented"

Admin consent was not granted. Run Step 4 again.

### "AADSTS700016: Application not found"

The service principal wasn't created. Run Step 1 again.

### "Scope not present in request"

The permission grant is missing scopes. Update the grant with all required scopes using the update command above.

### Users without Copilot license

Users without Copilot for M365 license will see reduced functionality. The proxy will automatically disable servers that require Copilot. To explicitly disable servers:

```json
{
  "env": {
    "AGENT365_DISABLED_SERVERS": "copilot"
  }
}
```

### Limiting to specific services

To only enable certain services:

```json
{
  "env": {
    "AGENT365_DISABLED_SERVERS": "copilot,excel,word"
  }
}
```

## Security Considerations

1. **Least Privilege**: Only grant permissions that are needed
2. **Audit Logs**: Monitor Azure AD sign-in logs for unusual activity
3. **Conditional Access**: Consider applying conditional access policies
4. **User Assignment**: Consider requiring user assignment to the app
5. **Token Lifetime**: Tokens expire after 1 hour by default

## Support

- Microsoft Agent 365 Documentation: https://learn.microsoft.com/en-us/microsoft-365-copilot/extensibility/
- Azure AD App Registration: https://learn.microsoft.com/en-us/azure/active-directory/develop/
- Report issues: https://github.com/rapyuta-robotics/agent365-mcp/issues
