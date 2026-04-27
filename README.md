# Uniware Automations API

FastAPI wrapper for Unicommerce/Uniware — inventory fetch and PO upload as REST endpoints.
Designed to be triggered from Excel (VBA) or Google Sheets (Apps Script).

## Structure

```
app/
├── main.py
├── routers/
│   ├── inventory.py     # GET  /inventory/fetch
│   └── po.py            # GET  /po/template, POST /po/upload
└── services/
    ├── auth.py          # Token management, api_post
    ├── inventory.py     # fetch + Excel builder
    └── po.py            # Excel parser, PO creator, results builder, template builder
```

## Endpoints

| Method | Path | Description |
|--------|------|-------------|
| GET | `/health` | Health check |
| GET | `/inventory/fetch?format=json` | Inventory as JSON |
| GET | `/inventory/fetch?format=excel` | Inventory as .xlsx download |
| GET | `/po/template` | Download blank PO Excel template |
| POST | `/po/upload` | Upload filled template, create POs |
| POST | `/po/upload?approve=true` | Create POs in APPROVED status |
| POST | `/po/upload?dry_run=true` | Validate Excel without hitting Unicommerce |

## Local Development

```bash
pip install fastapi "uvicorn[standard]" requests openpyxl python-dotenv
uvicorn app.main:app --reload --port 8080
# Swagger UI → http://localhost:8080/docs
```

## .env

```env
UNIWARE_TENANT_URL=https://yourcompany.unicommerce.com
UNIWARE_USERNAME=your@email.com
UNIWARE_PASSWORD=yourpassword
UNIWARE_CLIENT_ID=my-trusted-client
UNIWARE_FACILITY_CODE=
UNIWARE_ACCESS_TOKEN=
UNIWARE_REFRESH_TOKEN=
```

> On Cloud Run, set these as environment variables or Secret Manager secrets.
> Token persistence to .env is skipped automatically when no .env file is present (e.g. Cloud Run).

## Deploy to GCP Cloud Run

```bash
# Build & push
gcloud builds submit --tag gcr.io/YOUR_PROJECT_ID/uniware-api

# Deploy
gcloud run deploy uniware-api \
  --image gcr.io/YOUR_PROJECT_ID/uniware-api \
  --platform managed \
  --region asia-south1 \
  --allow-unauthenticated \
  --set-env-vars UNIWARE_TENANT_URL=https://yourcompany.unicommerce.com \
  --set-env-vars UNIWARE_USERNAME=your@email.com \
  --set-env-vars UNIWARE_PASSWORD=yourpassword \
  --set-env-vars UNIWARE_CLIENT_ID=my-trusted-client
```
