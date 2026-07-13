# =============================================================================
# PMO Project Initiation Portal — Local Deployment Guide
# =============================================================================
# Self-hosted internal portal for Windows / LAN environments.
# =============================================================================

## Prerequisites

- Windows 10/11 or Windows Server 2019+
- Node.js >= 18.0.0 ([Download](https://nodejs.org/))
- npm >= 9.0.0 (included with Node.js)
- Git for Windows ([Download](https://git-scm.com/download/win))
- Azure AD App Registration with `Sites.ReadWrite.All` and `User.Read.All`
- SharePoint Online site with Client Master and Main Tracker lists

## Installation

1. Clone or download this repository to your machine:
   ```powershell
   git clone https://github.com/your-org/cyber-pmo-project-initiation.git
   cd cyber-pmo-project-initiation
   ```

2. Install dependencies:
   ```powershell
   npm install
   ```

3. Copy `.env.example` to `.env` and fill in your Azure AD and SharePoint values:
   ```powershell
   Copy-Item .env.example .env
   notepad .env
   ```

## Environment Variables

| Variable | Required | Description |
|----------|----------|-------------|
| `PORT` | No | Server port (default: **45821**) |
| `NODE_ENV` | No | `production` or `development` |
| `TENANT_ID` | Yes | Azure AD tenant GUID |
| `CLIENT_ID` | Yes | Azure AD app registration client ID |
| `CLIENT_SECRET` | Yes | Azure AD app registration client secret |
| `USER_EMAIL` | No | Primary user / service account email |
| `SECRET_ID` | No | Azure AD secret identifier |
| `OBJECT_ID` | No | Azure AD app registration object ID |
| `SHAREPOINT_SITE_URL` | Yes | SharePoint site URL |
| `ALLOWED_ORIGINS` | No | CORS origins (default: `http://localhost:45821`) |
| `RATE_LIMIT_WINDOW_MS` | No | Rate limit window (default: 60000) |
| `RATE_LIMIT_MAX` | No | Max requests per window (default: 100) |

## Starting the Application

### Production Mode

```powershell
npm start
```

Output:
```
========================================
  PMO Project Initiation Portal
========================================
  Server Started Successfully

  Local:
  http://localhost:45821

  Network:
  http://192.168.1.50:45821

========================================
```

### Development Mode (with auto-reload)

```powershell
npm run dev
```

## Accessing from Another Device on Wi-Fi

1. Find your computer's local IPv4 address:
   ```powershell
   ipconfig
   ```
   Look for `IPv4 Address` under your active Wi-Fi/Ethernet adapter (e.g. `192.168.1.50`).

2. From another device on the same network, open:
   ```
   http://192.168.1.50:45821
   ```

3. Ensure Windows Firewall allows incoming connections on port **45821** (see below).

## Windows Firewall Configuration

The application listens on port **45821**. You must allow inbound traffic on this port.

### Using PowerShell (Administrator)

```powershell
New-NetFirewallRule -DisplayName "PMO Portal" -Direction Inbound -Protocol TCP -LocalPort 45821 -Action Allow
```

### Using Windows Defender Firewall GUI

1. Open **Windows Defender Firewall with Advanced Security**
2. Click **Inbound Rules** > **New Rule**
3. Select **Port** > **TCP** > **Specific local ports**: `45821`
4. Select **Allow the connection**
5. Apply to **Domain**, **Private**, **Public** (or at least Private for home/office networks)
6. Name: `PMO Project Initiation Portal`

### Verify Firewall Rule

```powershell
Get-NetFirewallRule -DisplayName "PMO Portal" | Format-Table -AutoSize
```

## Running as a Background Service (PM2)

[PM2](https://pm2.keymetrics.io/) is a production process manager that keeps your app running in the background and restarts it automatically on failure or reboot.

### Install PM2 Globally

```powershell
npm install -g pm2
```

### Start the Application

```powershell
pm2 start server.js --name project-initiation
```

### Save PM2 State

```powershell
pm2 save
```

### Configure PM2 to Start on Windows Boot

```powershell
pm2 startup
```

Follow the instructions printed by the command. Typically:

```powershell
pm2 save
$Env:PM2_HOME = "C:\Users\YourUser\.pm2"
pm2 startup
```

### Common PM2 Commands

```powershell
pm2 status                    # List running apps
pm2 logs project-initiation   # View logs
pm2 restart project-initiation # Restart app
pm2 stop project-initiation   # Stop app
pm2 delete project-initiation # Remove from PM2
```

## Health Check

The application exposes a health endpoint:

```
GET http://localhost:45821/health
```

Response:
```json
{
  "status": "OK",
  "uptime": 12345,
  "version": "1.0.0"
}
```

## API Endpoints

| Method | Path | Description |
|--------|------|-------------|
| GET | `/health` | Health check |
| GET | `/api/metadata` | Portal metadata |
| GET | `/api/users` | Search users |
| GET | `/api/bootstrap` | Load SharePoint data |
| POST | `/api/initiate` | Create Main Tracker items |

## Troubleshooting

### Port Already in Use

If port 45821 is occupied:

```powershell
netstat -ano | findstr :45821
taskkill /PID <PID> /F
```

Or change the port in `.env`:
```
PORT=45822
```

### Cannot Access from Another Device

1. Verify the server is listening on all interfaces:
   ```powershell
   netstat -ano | findstr :45821
   ```
   Should show `0.0.0.0:45821` or `[::]:45821`.

2. Check Windows Firewall allows port 45821 inbound.

3. Ensure both devices are on the same network/subnet.

4. Temporarily disable firewall to test (not recommended for production):
   ```powershell
   Set-NetFirewallProfile -Profile Private -Enabled False
   ```

### Azure AD / SharePoint Errors

- Verify `TENANT_ID`, `CLIENT_ID`, `CLIENT_SECRET` are correct in `.env`.
- Ensure the Azure AD app has `Sites.ReadWrite.All` and `User.Read.All` permissions.
- Check that the SharePoint site URL is correct and accessible.

### Application Crashes on Startup

- Run `npm start` manually to see the error output.
- Check for missing environment variables (the app exits if `TENANT_ID`, `CLIENT_ID`, or `CLIENT_SECRET` are missing).
- Ensure Node.js >= 18.0.0 is installed: `node --version`.

## Updating the Application

### From Git

```powershell
git pull origin main
npm install
pm2 restart project-initiation
```

### Manual Update

1. Stop the application:
   ```powershell
   pm2 stop project-initiation
   ```

2. Replace the application files with the updated version.

3. Install any new dependencies:
   ```powershell
   npm install
   ```

4. Start the application:
   ```powershell
   pm2 start server.js --name project-initiation
   ```

## Microsoft Graph Permissions

The Azure AD app registration requires:

- `Sites.ReadWrite.All` — Read/write SharePoint lists
- `User.Read.All` — Search users for the people picker

## Power Automate Integration

The `/api/initiate` endpoint writes items to the SharePoint Main Tracker list. Power Automate flows can be triggered on item creation in that list.

## License

Internal — Techdefence Labs Solutions Pvt. Ltd.
