# =============================================================================
# PMO Project Initiation Portal
# =============================================================================
# A metadata-driven project initiation portal built on Node.js, Express,
# Microsoft Graph, and SharePoint Online.
# =============================================================================

## Prerequisites

- Node.js >= 18.0.0
- npm >= 9.0.0
- Azure AD App Registration with `Sites.ReadWrite.All` and `User.Read.All`
- SharePoint Online site with Client Master and Main Tracker lists

## Quick Start

1. Clone the repository.
2. Copy `.env.example` to `.env` and fill in your Azure AD and SharePoint values.
3. Install dependencies:
   ```bash
   npm ci
   ```
4. Start the development server:
   ```bash
   npm run dev
   ```
5. Open `http://localhost:3000`.

## Environment Variables

| Variable | Required | Description |
|----------|----------|-------------|
| `PORT` | No | Server port (default: 3000) |
| `NODE_ENV` | No | Environment mode (`production` or `development`) |
| `TENANT_ID` | Yes | Azure AD tenant GUID |
| `CLIENT_ID` | Yes | Azure AD app registration client ID |
| `CLIENT_SECRET` | Yes | Azure AD app registration client secret |
| `USER_EMAIL` | No | Primary user / service account email |
| `SECRET_ID` | No | Azure AD app registration secret identifier |
| `OBJECT_ID` | No | Azure AD app registration object ID |
| `SHAREPOINT_SITE_URL` | Yes | SharePoint site URL (e.g. `https://contoso.sharepoint.com/sites/pmo`) |
| `ALLOWED_ORIGINS` | No | Comma-separated CORS origins |
| `RATE_LIMIT_WINDOW_MS` | No | Rate limit window in ms (default: 60000) |
| `RATE_LIMIT_MAX` | No | Max requests per window (default: 100) |
| `RATE_LIMIT_INITIATE_WINDOW_MS` | No | Initiate rate limit window in ms (default: 60000) |
| `RATE_LIMIT_INITIATE_MAX` | No | Max initiate requests per window (default: 10) |
| `HEALTH_CHECK_TIMEOUT_MS` | No | Health check downstream timeout in ms (default: 5000) |

## Deployment

### Docker

```bash
docker compose up -d
```

### PM2 (Ubuntu VPS)

```bash
pm2 start ecosystem.config.js
pm2 save
pm2 startup
```

### Azure App Service

1. Create a new Node.js App Service.
2. Set startup command to `pm2 start ecosystem.config.js` (or configure PM2 via `.deployment` file).
3. Add Application Settings for all environment variables.
4. Enable "Always On".

### Render / Railway

1. Connect the repository.
2. Set environment variables in the dashboard.
3. Set build command to `npm ci` and start command to `node server.js`.
4. Ensure the service listens on the port provided by `PORT`.

## API Endpoints

| Method | Path | Description |
|--------|------|-------------|
| GET | `/api/health` | Health check |
| GET | `/api/metadata` | Portal metadata (departments, services, project types) |
| GET | `/api/users` | Search users by query and/or department |
| GET | `/api/bootstrap` | Load clients and initiated departments from SharePoint |
| POST | `/api/initiate` | Create Main Tracker items for selected departments |

## Microsoft Graph Permissions

The Azure AD app registration requires:

- `Sites.ReadWrite.All` — Read/write SharePoint lists
- `User.Read.All` — Search users for the people picker

## Power Automate Integration

The `/api/initiate` endpoint writes items to the SharePoint Main Tracker list. Power Automate flows can be triggered on item creation in that list.

## Logs

- Application logs are output to stdout/stderr.
- In Docker, logs are managed by the container runtime.
- In PM2, logs are written to `logs/pm2-error.log` and `logs/pm2-out.log`.

## License

Internal — Techdefence Labs Solutions Pvt. Ltd.
