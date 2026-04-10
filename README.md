# Microsoft Teams App

Keeper workflow and EPM integration with Microsoft Teams.

## Overview

The **Keeper Teams App** helps achieve zero standing privilege and streamlines credential workflow requests and approvals directly from Teams. The customer hosts the Teams agent and Commander Service Mode, ensuring that zero knowledge is maintained with end-to-end encryption.

This document describes the installation of the Keeper Teams App using a streamlined setup method that requires the use of Keeper Secrets Manager. If you don't have a Secrets Manager or KeeperPAM license, please contact your Keeper account manager.

---

## Features

| Feature | Description |
|---------|-------------|
| **Record Access Requests** | Request access to specific Keeper records with justification, custom permissions and access time limits. This includes standard vault records and KeeperPAM resources. |
| **Folder Access Requests** | Request access to specific Keeper Shared Folders with justification, custom permissions and access time limits. |
| **One-Time Share Requests** | Request for a one-time share, password reset or other dynamic password generation with a self-destructing share link. The one-time share can also be editable, offering bi-directional sharing capabilities. |
| **Endpoint Privilege Manager Approvals** | Keeper Endpoint Privilege Manager (KEPM) just-in-time elevation approvals in realtime through a dedicated Teams channel. |
| **SSO Cloud Device Approvals** | Perform approvals of SSO Cloud devices directly through Teams, if the Keeper Automator service is not deployed. |

---

## Prerequisites

### System Requirements

To maintain zero knowledge and full end-to-end encryption, the Keeper Teams App and Commander Service Mode containers are hosted by each customer on their own infrastructure to interact with the Microsoft Teams cloud service. Commander is used locally to help set everything up.

| Requirement | Details |
|-------------|---------|
| Linux VM | Any VM in the cloud or on-prem which can establish https/443 outbound connections to Microsoft and Keeper services. |
| Docker | Docker is the recommended method for setting up the service |
| Keeper Commander | Service Mode running and accessible |
| Keeper Secrets Manager | Either Keeper Secrets Manager or KeeperPAM license used for retrieving the secret configuration data |
| Microsoft 365 Tenant | Requires admin access to register applications in Azure Active Directory |
| Microsoft Teams | Teams workspace with permissions to install custom apps |

> **Important:** The `teams-app-setup` command requires Keeper Secrets Manager (KSM) to be activated. If KSM is not available, please contact your account manager.

---

## Setup Steps

Follow these seven steps to configure the Teams app:

1. [Register Azure AD Application](#step-1-register-azure-ad-application)
2. [Create Azure Bot Registration](#step-2-create-azure-bot-registration)
3. [Create Approvals Channel](#step-3-create-approvals-channel)
4. [Commander Service Mode Setup](#step-4-commander-service-mode-setup)
5. [Run Teams App Setup Command](#step-5-run-teams-app-setup-command)
6. [Deploy to Docker Environment](#step-6-deploy-to-docker-environment)
7. [Upload Teams App Package](#step-7-upload-teams-app-package)

---

### Step 1. Register Azure AD Application

In this section, you will create an Azure AD Application in your Microsoft 365 tenant to authenticate the Teams bot.

1. Sign in to the [Azure Portal](https://portal.azure.com) as a Global Administrator or Application Administrator

2. Navigate to **Azure Active Directory** → **App registrations**

3. Click **New registration**

4. Configure the application:
   - **Name:** `Keeper Security Bot`
   - **Supported account types:** "Accounts in this organizational directory only (Single tenant)"
   - **Redirect URI:** Leave blank for now

5. Click **Register**

6. After creation, note the following values from the **Overview** page:
   - **Application (client) ID** - Save this as `AZURE_CLIENT_ID`
   - **Directory (tenant) ID** - Save this as `AZURE_TENANT_ID`

7. Configure **API Permissions**:
   - Go to **API permissions** → **Add a permission**
   - Select **Microsoft Graph** → **Application permissions**
   - Add the following permissions:
     - `User.Read.All`
     - `ChannelMessage.Send`
     - `TeamsActivity.Send`
   - Click **Grant admin consent for [Your Organization]**

8. Create a **Client Secret**:
   - Go to **Certificates & secrets** → **Client secrets**
   - Click **New client secret**
   - Description: `Keeper Teams App`
   - Expiration: Select appropriate duration (recommend 24 months)
   - Click **Add**
   - **Copy the Value immediately** - Save this as `AZURE_CLIENT_SECRET`

> **Important:** Save the **Application (client) ID**, **Directory (tenant) ID**, and **Client Secret** for Step 5.

---

### Step 2. Create Azure Bot Registration

1. In the Azure Portal, click **Create a resource**

2. Search for **Azure Bot** and click **Create**

3. Configure the bot:
   - **Bot handle:** `keeper-security-bot` (must be unique)
   - **Subscription:** Select your subscription
   - **Resource group:** Create new or use existing
   - **Pricing tier:** Free (F0) for testing, Standard (S1) for production
   - **Type of App:** Single Tenant
   - **Creation type:** Use existing app registration
   - **App ID:** Enter the Application (client) ID from Step 1

4. Click **Review + create** → **Create**

5. After deployment, go to the Bot resource:
   - Navigate to **Configuration**
   - Set **Messaging endpoint:** `https://<your-domain>/api/messages`
   - Click **Apply**

6. Enable the **Microsoft Teams** channel:
   - Go to **Channels**
   - Click **Microsoft Teams** (Commercial)
   - Accept the Terms of Service
   - Click **Apply**

7. Note the **Microsoft App ID** (same as Application ID from Step 1) - Save this as `BOT_ID`

> **Important:** Save the **Bot ID** for Step 5.

---

### Step 3. Create Approvals Channel

1. In Microsoft Teams, create a new **Private channel** for approvals:
   - Right-click on your Team → **Add channel**
   - **Channel name:** `keeper-vault-approvers`
   - **Privacy:** Private - Only specific teammates can access
   - Click **Create**

2. Get the **Team ID** and **Channel ID**:
   
   **Option A: Using Teams Web**
   - Open Teams in a web browser
   - Navigate to the approvals channel
   - The URL will look like: `https://teams.microsoft.com/l/channel/19%3A...@thread.tacv2/...?groupId=<TEAM_ID>&tenantId=...`
   - Extract the `groupId` as **Team ID**
   - The channel portion `19:...@thread.tacv2` (URL decoded) is the **Channel ID**

   **Option B: Using Graph Explorer**
   - Go to [Graph Explorer](https://developer.microsoft.com/en-us/graph/graph-explorer)
   - Run: `GET https://graph.microsoft.com/v1.0/me/joinedTeams`
   - Find your team and note the `id` as **Team ID**
   - Run: `GET https://graph.microsoft.com/v1.0/teams/{team-id}/channels`
   - Find your channel and note the `id` as **Channel ID**

3. **Add the bot to the channel** (after Step 6):
   - In the channel, click **+** to add a tab
   - Or mention the bot: `@Keeper Security`

> **Important:** Save the **Team ID** and **Channel ID** for Step 5.

---

### Step 4. Commander Service Mode Setup

To enable the service to authenticate and execute commands within the Keeper tenant, an authorized **Keeper Commander configuration file** must be created. This configuration can be generated on a host computer or workstation.

1. [Install Keeper Commander](https://docs.keeper.io/en/keeperpam/commander-cli/commander-installation-setup) locally on your machine

2. If required, create a new Keeper service account dedicated to this integration, ensuring it has access to the relevant records and folders and the ability to perform record and folder sharing.

3. Login to Commander with the Keeper Service account:

```bash
keeper shell
My Vault> login serviceuser@company.com
```

4. Complete the authentication process including any 2FA requirements. Once you are fully authenticated, proceed to Step 5.

---

### Step 5. Run Teams App Setup Command

The `teams-app-setup` command generates a `docker-compose.yml` file which you will use to operate the Teams App and Commander Service Mode services.

From the Commander shell, type:

```
teams-app-setup
```

#### Command Line Options

The `teams-app-setup` command supports the following optional flags for customization:

| Parameter | Description | Default Value |
|-----------|-------------|---------------|
| `--folder-name` (optional) | Name for the shared folder | Commander Service Mode - Teams App |
| `--app-name` (optional) | Name for the Secrets Manager app | Commander Service Mode - KSM App |
| `--config-record-name` (optional) | Name for the Commander config record | Commander Service Mode Docker Config |
| `--teams-record-name` (optional) | Name for the Teams config record | Commander Service Mode Teams App Config |
| `--config-path` (optional) | Path to config.json file | ~/.keeper/config.json |
| `--timeout` (optional) | Device timeout setting | 30d |
| `--skip-device-setup` (optional) | Skip device registration if already configured | false |

Example with Custom Names:

```bash
teams-app-setup --folder-name "My Teams Integration" --timeout 7d
```

The command will guide you through the following prompts:

#### Phase 1: Docker Service Mode Setup

It automatically configures KSM and uploads the config file required for setting up service mode via Docker.

```
Phase 1: Running Docker Service Mode Setup
═══════════════════════════════════════════════════════════
    Docker Setup
═══════════════════════════════════════════════════════════

[1/7] Checking device settings...
  ✓  Device already registered
  ✓  Persistent login already enabled
  ✓  Setting logout timeout to 30d...

[2/7] Creating shared folder 'Commander Service Mode - Teams App'...
  ✓  Shared folder created successfully

[3/7] Creating record 'Commander Service Mode Docker Config'...
  ✓  Record created successfully

[4/7] Uploading config.json attachment...
  ✓  Config file uploaded successfully

[5/7] Creating Secrets Manager app 'Commander Service Mode - KSM App'...
  ✓  App created successfully

[6/7] Sharing folder with app...
  ✓  Folder shared with app

[7/7] Creating client device and generating config...
  ✓  Client device created successfully

✓ Docker Setup Complete!
```

#### Service Configuration

Configure the Commander Service port:

| Prompt | Description | Example |
|--------|-------------|---------|
| Port | Port number for Commander Service Mode (1024-65535) | 8900 |

#### Phase 2: Teams App Integration Setup

Enter the Azure credentials obtained from **Steps 1, 2, and 3**:

| Prompt | Description | Example |
|--------|-------------|---------|
| Azure Client ID (required) | Application (client) ID from Step 1 | xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx |
| Azure Client Secret (required) | Client secret value from Step 1 | xxxxxxxxxxxxxxxxxxxxxxxxxxxxxx |
| Azure Tenant ID (required) | Directory (tenant) ID from Step 1 | xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx |
| Bot ID (required) | Microsoft App ID from Step 2 (same as Client ID) | xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx |
| Bot Password (required) | Same as Azure Client Secret | xxxxxxxxxxxxxxxxxxxxxxxxxxxxxx |
| Teams Team ID (required) | Team ID from Step 3 | xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx |
| Teams Channel ID (required) | Channel ID from Step 3 | 19:xxxxxx@thread.tacv2 |
| Enable EPM? (optional) | Enable Endpoint Privilege Manager approvals (y/n) | y |
| EPM Polling Interval (optional) | How often to check for EPM requests in seconds. Default: 120 | 120 |
| Enable Device Approvals? (optional) | Enable SSO Cloud device approvals (y/n) | y |
| Device Approval Polling Interval (optional) | How often to check for device approvals in seconds. Default: 120 | 120 |

> **Note:** In order to process Endpoint Privilege Manager approvals and SSO Cloud approvals, the Teams App service user must have administrative permissions "Manage Endpoint Privilege" and "Managing the Keeper Admin Console".

After the command executes successfully, it automatically performs the following actions:

- Configures persistent device authentication
- Creates a Shared Folder named **"Commander Service Mode – Teams App"**
- Creates a KSM application with access to the shared folder
- Creates a client device and generates a Base64-encoded configuration value
- Creates a Docker Config record and uploads the `config.json` file from the `.keeper` directory
- Creates a Teams App Config record containing the Azure/Bot credentials

```
✓ Teams App Integration Setup Complete!

Resources Created:
  Phase 1 - Commander Service:
    • Shared Folder: Commander Service Mode - Teams App
    • KSM App: Commander Service Mode - KSM App (with edit permissions)
    • Config Record: XXXXXX
    • KSM Base64 Config: ✓ Generated
  Phase 2 - Teams App:
    • Teams Config Record: XXXXXX
    • Approvals Channel: XXXXXX
    • EPM Integration: true
    • Device Approval: true
```

Upon successful execution, a `docker-compose.yml` is generated containing both the Commander Service Mode and Teams App services, ready for deployment.

```yaml
services:
  commander:
    container_name: keeper-service
    ports:
    - 127.0.0.1:<port>:<port>
    image: keeper/commander:latest
    command: service-create -p <port> -c 'search,share-record,share-folder,share-report,record-add,tree,one-time-share,pedm,device-approve,get' -f json -q y -ur <CONFIG_RECORD_UID> --ksm-config <KSM_CONFIG_BASE64_VALUE> --record <CONFIG_RECORD_UID>
    healthcheck:
      test:
      - CMD-SHELL
      - python -c "import sys, urllib.request; sys.exit(0 if urllib.request.urlopen('http://localhost:<port>/health', timeout=2).status == 200 else 1)"
      interval: 60s
      timeout: 3s
      start_period: 10s
      retries: 30
    restart: unless-stopped
    
  teams-app:
    container_name: keeper-teams-app
    image: keeper/teams-app:latest
    ports:
    - "3978:3978"
    environment:
      KSM_CONFIG: <KSM_CONFIG_BASE64_VALUE>
      COMMANDER_RECORD: <CONFIG_RECORD_UID>
      TEAMS_RECORD: <TEAMS_CONFIG_RECORD_UID>
    depends_on:
      commander:
        condition: service_healthy
    restart: unless-stopped
```

Once setup is complete, ensure that the Commander session is terminated and the local `.keeper/config.json` file is deleted to prevent device token conflicts.

```bash
My Vault> quit
$ rm ~/.keeper/config.json
```

---

### Step 6. Deploy to Docker Environment

In this section, you will set up a Docker Compose environment on a Linux virtual machine or host where the Commander Service will run.

1. Launch a Linux VM or prepare a Linux host and connect to it via SSH.

2. Install `docker` and `docker-compose` (refer to the installation instructions [here](https://docs.keeper.io/en/keeperpam/privileged-access-manager/references/installing-docker-on-linux))

3. Transfer the generated `docker-compose.yml` file from Step 5 to the target Linux server.

4. Start up the services on the host machine:

```bash
docker compose up -d
```

#### Service Startup Sequence

The services start sequentially:

1. Commander Service starts first, generates an API key, and saves it along with the service URL to the vault record
2. Health checks validate the Commander service is running
3. Teams App starts after health checks pass, automatically retrieving the API key and service URL from the vault record

#### Verify Successful Startup

Monitor the logs to make sure everything starts up.

**Check container status:**

```bash
$ docker ps
NAME                STATUS                    PORTS
keeper-service      Up (healthy)              127.0.0.1:<port> -> <port>/tcp
keeper-teams-app    Up                        0.0.0.0:3978 -> 3978/tcp
```

**View Commander Service logs:**

```bash
$ docker logs keeper-service
[2026-01-21 10:00:00] Starting Commander Service Mode...
Generated API key: ****nQ= (stored in vault record: <CONFIG_VAULT_RECORD>)
Commander Service starting on <SERVICE_URL>/api/v2
Keeper Commander Service initialization complete
```

> **Note:** The API key is redacted in Docker logs for security. Both services communicate securely via the shared vault record.

**View Teams App logs:**

```bash
$ docker logs keeper-teams-app
```

If everything is successful, you'll see the messages below:

```
============================================================
Starting Keeper Teams App
============================================================
[INFO] Config loaded from KSM vault
[INFO] Initializing Keeper Commander Teams App...
[OK] Configuration loaded
[OK] Keeper client initialized: http://commander:<port>/api/v2
[OK] Bot Framework adapter initialized
[OK] All handlers registered

============================================================
Starting Keeper Commander Teams App
============================================================
[OK] Bot listening on port 3978
[INFO] Approval channel configured: <channel-id>
[OK] EPM poller initialized (enabled, interval: 120s)
[OK] Cloud SSO Device Approval poller initialized (enabled, interval: 120s)

⚡ Teams App is running!
```

---

### Step 7. Upload Teams App Package

After the Docker services are running, upload the Teams app package to your Microsoft Teams environment.

#### Create the App Package

1. Download the app package template from the [releases page](https://github.com/Keeper-Security/teams-app/releases)

2. Extract the ZIP file, which contains:
   ```
   keeper-teams-app/
   ├── manifest.json
   ├── color.png
   └── outline.png
   ```

3. Edit `manifest.json` and replace the placeholders:
   - Replace `${{TEAMS_APP_ID}}` with your **Azure Client ID** (from Step 1)
   - Replace `${{BOT_ID}}` with your **Bot ID** (same as Azure Client ID)
   - Remove `${{APP_NAME_SUFFIX}}` or replace with empty string

4. Repackage the files into a ZIP:
   ```bash
   zip keeper-teams-app.zip manifest.json color.png outline.png
   ```

#### Upload to Teams Admin Center (Recommended for Organization-wide)

1. Sign in to [Microsoft Teams Admin Center](https://admin.teams.microsoft.com)

2. Navigate to **Teams apps** → **Manage apps**

3. Click **Upload new app** → **Upload**

4. Select your `keeper-teams-app.zip` file

5. After upload, find the app and click **Publish**

6. Configure app policies to allow users to install the app

#### Upload for Testing (Personal/Team only)

1. In Microsoft Teams client, click **Apps** in the sidebar

2. Click **Manage your apps** → **Upload an app**

3. Choose **Upload a custom app**

4. Select your `keeper-teams-app.zip` file

5. Click **Add** to install for yourself or **Add to a team** for a specific team

---

## Command Reference for Requesting User

### keeper-request-record

Request access to a specific Keeper record.

**Syntax:**

```
keeper-request-record <record-uid-or-description> <justification>
```

**Examples:**

```
keeper-request-record kR3cF9Xm2Lp8NqT1uV6w Emergency server access
keeper-request-record "prod db EU region" Need to run migration
keeper-request-record AWS-Production Deployment ticket #12345
```

### keeper-request-folder

Request access to a shared folder.

**Syntax:**

```
keeper-request-folder <folder-uid-or-description> <justification>
```

**Examples:**

```
keeper-request-folder kF8zQ2Nm5Wx9PtR3sY7a Need staging access
keeper-request-folder "Staging Team Folder" Need staging access for deployment
```

### keeper-one-time-share

Request a one-time share link for a record.

**Syntax:**

```
keeper-one-time-share <record-uid-or-description> <justification>
```

**Examples:**

```
keeper-one-time-share kR3cF9Xm2Lp8NqT1uV6w Need to share with contractor John
keeper-one-time-share "AWS Production Password" Sharing with vendor for audit
```

### keeper-create-secret

Create a new login record directly in your Keeper vault. The record is saved to a shared folder that you have access to, with optional subfolder selection.

**Syntax:**

```
keeper-create-secret ["<title>"] ["<notes>"]
```

Both title and notes are optional — if provided, they pre-fill the form. Wrap values in quotes if they contain spaces.

**Examples:**

```
keeper-create-secret
keeper-create-secret "Staging new cred's"
keeper-create-secret "Staging new cred's" Testing Note
```

**Features:**
- Shared folder selection (required) with subfolder support
- Optional auto-generate password via checkbox
- Zero-knowledge disclaimer on password field
- Notification sent to approvers channel on successful creation

**Aliases:** `create-secret`, `createsecret`, `kcs`

### help

Display available commands and usage information.

```
help
```

---

## Configuration Reference

### Environment Variables

| Variable | Description | Required |
|----------|-------------|----------|
| `MICROSOFT_APP_ID` | Azure Bot Application ID | Yes |
| `MICROSOFT_APP_PASSWORD` | Azure Bot Application Password (Client Secret) | Yes |
| `AZURE_TENANT_ID` | Azure Active Directory Tenant ID | Yes |
| `AZURE_CLIENT_ID` | Azure AD Application Client ID | Yes |
| `AZURE_CLIENT_SECRET` | Azure AD Application Client Secret | Yes |
| `TEAMS_TEAM_ID` | Team ID containing the approvals channel | Yes |
| `TEAMS_CHANNEL_ID` | Approvals channel ID | Yes |
| `COMMANDER_URL` | Commander Service Mode URL | Yes |
| `COMMANDER_API_KEY` | Commander Service Mode API Key | Auto-configured |
| `PORT` | Teams App listening port | No (default: 3978) |
| `LOG_LEVEL` | Logging level (debug, info, warn, error) | No (default: info) |
| `EPM_ENABLED` | Enable Endpoint Privilege Manager polling | No (default: false) |
| `EPM_POLL_INTERVAL` | EPM polling interval in seconds | No (default: 120) |
| `DEVICE_APPROVAL_ENABLED` | Enable SSO Cloud device approval polling | No (default: false) |
| `DEVICE_APPROVAL_POLL_INTERVAL` | Device approval polling interval in seconds | No (default: 120) |

### Configuration via KSM Record

When using Keeper Secrets Manager, the configuration is stored in a vault record with the following fields:

| Field | Description |
|-------|-------------|
| `microsoft_app_id` | Azure Bot Application ID |
| `microsoft_app_password` | Azure Bot Password |
| `tenant_id` | Azure AD Tenant ID |
| `azure_client_id` | Azure AD Client ID |
| `azure_client_secret` | Azure AD Client Secret |
| `approvals_team_id` | Teams Team ID for approvals channel |
| `approvals_channel_id` | Teams Approvals Channel ID |
| `pedm_enabled` | EPM polling enabled (true/false) |
| `pedm_polling_interval` | EPM polling interval in seconds |
| `device_approval_enabled` | Device approval enabled (true/false) |
| `device_approval_polling_interval` | Device approval polling interval in seconds |

---

## Updates

### Updating the Commander Service Mode and Teams App Containers

To update to the latest version of Commander or the Teams App, follow the steps below to stop the service, update the containers and start up the new containers.

```bash
docker compose down
docker compose pull
docker compose up -d
```

---

## Troubleshooting

### Startup Errors

| Error | Cause | Solution |
|-------|-------|----------|
| Commander Service Mode is prompting for master password | Multiple config.json files are attached to the Vault record | Follow steps 4-5 to run the `teams-app-setup` command with new folder name again to create a new JSON config file. |
| Cannot reach Keeper Service Mode | Service Mode not running or wrong URL | Verify the service URL in the vault record is as expected |
| BotFrameworkAdapter initialization failed | Invalid bot credentials | Verify MICROSOFT_APP_ID and MICROSOFT_APP_PASSWORD |
| Azure AD authentication error | Invalid tenant or client credentials | Verify AZURE_TENANT_ID, AZURE_CLIENT_ID, and AZURE_CLIENT_SECRET |

### Teams API Errors

| Error | Cause | Solution |
|-------|-------|----------|
| Conversation not found | Invalid team or channel ID | Verify TEAMS_TEAM_ID and TEAMS_CHANNEL_ID |
| Authorization denied (401) | Bot not properly configured or token expired | Regenerate client secret and update configuration |
| Forbidden (403) | Missing API permissions | Ensure Graph API permissions are granted admin consent |
| Channel not found | Bot not added to channel | Add the bot to the approvals channel |
| Activity not found | Message was deleted or activity ID invalid | This may occur when updating old cards; can be ignored |

### Access Grant Errors

| Error | Cause | Solution |
|-------|-------|----------|
| Record Not Found | Invalid UID or record deleted | Verify the record UID exists in Keeper vault |
| Folder Not Found | Invalid folder UID | Verify the folder UID exists in Keeper vault |
| Invalid UID Type (record vs folder) | Used wrong command for item type | Use `keeper-request-folder` for folders, `keeper-request-record` for records |
| User already has time-limited access | Conflict with existing share | Revoke existing access first, then grant new permission |
| Share permissions require permanent access | Trying to use duration with Can Share/Edit & Share | Share permissions (Can Share, Edit & Share, Change Owner) are always permanent |
| User share failed | Permission conflict on folder | User may have incompatible existing access; revoke and re-grant |
| Cannot grant access to record owner | User is the record owner | Record owners already have full access; no action needed |

### Search & Approval Card Errors

| Error | Cause | Solution |
|-------|-------|----------|
| No records found matching... | Search query too specific or no matches | Try broader search terms; check record exists in vault |
| Search command timed out | Service Mode slow or vault very large | Increase timeout or use more specific search |
| Something went wrong (on card) | Operation took longer than Teams timeout | Card will update automatically when operation completes |
| Card not updating | Activity ID mismatch or permissions | Check bot has permission to update messages in channel |

### One-Time Share Errors

| Error | Cause | Solution |
|-------|-------|----------|
| One-time share links cannot be created for PAM records | PAM records don't support OTS | Use `keeper-request-record` for PAM records instead |
| Share link created but URL not found | Unexpected Service Mode response | Check Service Mode version; verify one-time-share command registered |
| Failed to create one-time share | Record may not be shareable | Verify user has share permissions on the record |

### EPM Errors

| Error | Cause | Solution |
|-------|-------|----------|
| No data returned | EPM feature not enabled | Enable EPM in your Keeper enterprise settings. Ensure service user has admin permissions. |
| EPM sync failed | Service Mode can't reach EPM server | Check network connectivity and EPM configuration |
| Failed to approve/deny EPM request | Request may have expired | Check if request is still pending; it may have auto-expired |

### Device Approval Errors

| Error | Cause | Solution |
|-------|-------|----------|
| No pending device approvals | No devices waiting for approval | Normal if no users are pending SSO device approval |
| Failed to approve device | Service user lacks admin permissions | Ensure service account has "Managing the Keeper Admin Console" permission |

---

## Security Considerations

### Zero Knowledge Architecture

The Keeper Teams App maintains zero knowledge principles:

- All credential data remains encrypted end-to-end
- The Teams App never has access to decrypted secrets
- Commander Service Mode handles all vault operations locally
- Microsoft Teams only sees approval workflow metadata, not actual credentials

### Network Security

- Configure firewall rules to restrict access to Commander Service Mode
- Use HTTPS for all external communications
- Consider placing services behind a reverse proxy with TLS termination

### Credential Rotation

- Rotate Azure AD client secrets periodically (recommended: every 12 months)
- Monitor for expired secrets in Azure AD
- Update Docker configuration when rotating secrets

---

## References

- [Commander CLI Overview](https://docs.keeper.io/en/keeperpam/commander-cli)
- [Commander Service Mode](https://docs.keeper.io/en/keeperpam/commander-cli/service-mode-rest-api)
- [Endpoint Privilege Manager](https://docs.keeper.io/en/keeperpam/endpoint-privilege-manager)
- [SSO Connect Cloud](https://docs.keeper.io/en/enterprise-guide/sso-connect-cloud)
- [Microsoft Teams Bot Documentation](https://docs.microsoft.com/en-us/microsoftteams/platform/bots/what-are-bots)
- [Azure Bot Service Documentation](https://docs.microsoft.com/en-us/azure/bot-service/)
- [Microsoft Graph API](https://docs.microsoft.com/en-us/graph/overview)
