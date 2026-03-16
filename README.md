# BSC Inspection Server

Backend server for Bharat Steel (Chennai) Inspection Portal.
Receives form submissions and saves to OneDrive automatically.

## Environment Variables (set in Render.com)

| Variable | Value |
|---|---|
| `CLIENT_ID` | Azure App Client ID |
| `CLIENT_SECRET` | Azure App Client Secret |
| `TENANT_ID` | Azure Tenant ID |
| `USER_ID` | pdqc@bharatsteels.in |

## Endpoint

`POST /submit` — receives form data, saves PDF + Excel row to OneDrive
