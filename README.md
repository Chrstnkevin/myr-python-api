# MYR Python API

Python API untuk generate file Excel (.xlsx) dari template.

## Deploy ke Railway

1. Push repo ini ke GitHub
2. Buka [railway.app](https://railway.app) → New Project → Deploy from GitHub
3. Pilih repo ini → Railway otomatis detect Python dan deploy
4. Setelah deploy, copy URL yang diberikan Railway
5. Set di Vercel project (Next.js) → Settings → Environment Variables:
   ```
   PYTHON_API_URL = https://your-app.railway.app/generate-xls
   ```

## Endpoint

`POST /generate-xls`

Body (JSON):
```json
{
  "header": { "judulDokumen": "...", "keterangan": "...", ... },
  "entries": [ ... ],
  "template_b64": "<base64 of xlsx template>"
}
```
