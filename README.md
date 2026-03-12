# Project Schedule Forecast Dashboard

React + TypeScript dashboard for weekly hours forecasting from an Excel workbook.

## Features

- Upload or replace a `.xlsx` workbook in-browser
- Uses `Hours_03-05-26.xlsx` from `public/` as initial development data
- Parses task rows (`Name`, `Work`, `Start`, `Finish`, `Resource Names`, and `Project`/`Proje`)
- Distributes each task's `Work` hours evenly across days from `Start` to `Finish`
- Rolls up daily values into Monday-based weekly buckets
- Optional weekend inclusion toggle
- Stacked weekly forecast bars grouped by `Resource Names` or `Project`
- Capacity line overlay (global capacity)
- Editable per-week capacity values in the table (nice-to-have)
- Highlights over-capacity weeks in chart and table
- Filters for resource, project, and date range
- Summary cards for total hours, capacity, variance, and count of over-capacity weeks
- Sortable weekly table
- Export weekly forecast table to CSV
- Loading and parse error states

## Tech stack

- React 19 + TypeScript
- Vite
- Recharts (charting)
- SheetJS `xlsx` (Excel parsing)
- `date-fns` (date normalization / Monday week bucketing)

## Folder structure

- `src/App.tsx`: app shell, controls, summary, state orchestration
- `src/components/ForecastChart.tsx`: stacked bars + capacity line + tooltips
- `src/components/PivotPlanningTable.tsx`: interactive planning pivot/editor
- `src/components/ForecastTable.tsx`: sortable weekly summary table
- `src/components/MultiSelectProjects.tsx`: searchable multi-select project filter
- `src/utils/excel.ts`: workbook parsing and column normalization
- `src/utils/planner.ts`: planning data model, week aggregation, overrides
- `src/utils/reportExport.ts`: client-side Excel export helpers (fallback path)
- `src/utils/reportExportApi.ts`: backend chart-export API client
- `api/export-report.py`: Vercel serverless API route for embedded chart export
- `backend/export_api.py`: optional Python API for embedded Excel chart export
- `backend/requirements.txt`: Python dependencies for export API
- `requirements.txt`: Python dependencies for Vercel serverless API
- `src/types.ts`: shared TypeScript models
- `public/Hours_03-05-26.xlsx`: initial workbook used for development
- `vercel.json`: Vercel build config
- `netlify.toml`: Netlify build + SPA redirect config

## Run locally

1. Install dependencies:

```bash
npm install
```

2. Start dev server:

```bash
npm run dev
```

3. Open the local Vite URL shown in terminal.

4. Build for production:

```bash
npm run build
```

5. Optional local production preview:

```bash
npm run preview
```

## Embedded Excel Chart Export (Optional Backend)

The app can export with a real embedded Excel chart when the local Python export API is running.

1. Install Python dependencies:

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r backend/requirements.txt
```

2. Start the export API:

```bash
python backend/export_api.py
```

3. Keep it running at `http://127.0.0.1:8000`.

4. In the dashboard, click `Export Report Excel`.
   - If API is running: exported workbook includes embedded chart in `Weekly Capacity Chart`.
   - If API is not running: app falls back to client-side export (chart data sheet without embedded chart object).

## Shared Workbook Storage (PC Source of Truth)

The app now supports backend-backed shared workbook storage so uploads are not browser-local only.

- Primary workbook (`main`) and Sales workbook (`sales`) are stored on disk by the backend.
- Shared planning state (manual overrides, enabled resources, weekly caps, filters, weekend settings) is also stored on disk.
- Frontend loads shared data from backend APIs at startup and on refresh.
- Uploads are validated (`.xlsx`, size limit, basic signature check), then written atomically.

### Backend APIs used for shared data

- `GET /api/workbook-state?dataset=main|sales`
- `GET /api/workbook-file?dataset=main|sales`
- `POST /api/upload-workbook?dataset=main|sales`
- `GET /api/shared-state`
- `PUT /api/shared-state`
- `GET /api/shared-health`

### Shared store location (on your PC)

By default:

- `backend/shared_store/workbooks/main.xlsx`
- `backend/shared_store/workbooks/sales.xlsx`
- `backend/shared_store/manifest.json`
- `backend/shared_store/shared_state.json`

You can override the storage directory with:

- `CAPACITY_SHARED_DATA_DIR`

### Environment variables

- Frontend:
  - `VITE_SHARED_DATA_API_URL` (for example `http://192.168.1.10:8000`)
  - `VITE_EXPORT_API_URL` (optional override for export API)
- Backend:
  - `CAPACITY_SHARED_DATA_DIR` (optional custom shared store path)
  - `CAPACITY_MAX_UPLOAD_BYTES` (optional upload size limit in bytes)
  - `CAPACITY_API_HOST` (default `0.0.0.0`)
  - `CAPACITY_API_PORT` (default `8000`)
  - `CAPACITY_API_DEBUG` (`true`/`false`)

### Run shared mode locally

1. Start backend API:

```bash
python backend/export_api.py
```

2. Ensure frontend points to backend:

```bash
# .env.local
VITE_SHARED_DATA_API_URL=http://127.0.0.1:8000
```

3. Start frontend:

```bash
npm run dev
```

If users connect over your LAN, use your PC LAN IP in `VITE_SHARED_DATA_API_URL` and make sure firewall/router rules allow access to backend port 8000.

## Embedded Chart Export on Public Deploy (Vercel)

This repo includes `api/export-report.py`, so Vercel can host the chart-export API at the same domain:

- `/api/export-report`

After deploying to Vercel, `Export Report Excel` will call this hosted endpoint automatically and generate an Excel file with embedded chart from your public website.

## Deploying the App Publicly

This repo includes a static frontend plus optional Python backend services.

- If you deploy only the frontend, shared workbook source-of-truth storage will not exist.
- For shared uploads/data across users, deploy the backend API somewhere always-on (your PC server, VM, or cloud) and point frontend at it via `VITE_SHARED_DATA_API_URL`.

### Option A (Recommended): Vercel

1. Push the project to GitHub.
2. Go to https://vercel.com and sign in.
3. Click **Add New > Project** and import your GitHub repo.
4. Vercel will detect Vite automatically.
5. Confirm build settings:
   - Build Command: `npm run build`
   - Output Directory: `dist`
6. Click **Deploy**.
7. Share the generated URL, for example:
   - `https://your-app.vercel.app`

Auto-redeploy:
- Every push to your connected branch triggers a new deployment automatically.

Custom domain later:
- In Vercel project settings, open **Domains** and add your domain.

### Option B: Netlify

1. Push the project to GitHub.
2. Go to https://app.netlify.com and sign in.
3. Click **Add new site > Import an existing project**.
4. Select your GitHub repo.
5. Netlify will use `netlify.toml`:
   - Build Command: `npm run build`
   - Publish Directory: `dist`
6. Deploy and share your URL:
   - `https://your-app.netlify.app`

Auto-redeploy:
- New commits to the connected branch auto-trigger deploys.

Custom domain later:
- Use **Site settings > Domain management**.

### Option C: GitHub Pages

Use this if you want hosting directly from GitHub. This requires a GitHub Actions workflow and Vite base-path config for your repo name.

High-level steps:
1. Push repo to GitHub.
2. Add a GitHub Actions workflow to build and publish `dist`.
3. Enable GitHub Pages in repo settings.
4. Share the Pages URL.

### Build command reference

```bash
npm run build
```

This outputs the deployable static site into `dist/`.

### How to update the deployed site

1. Commit your changes.
2. Push to the connected branch (for example, `main`).
3. Hosting platform automatically rebuilds and redeploys.
4. Share the same public URL; it always serves latest successful deployment.

## Assumptions

- `Work` is interpreted as hours.
- Work distribution is even across active days in each task range.
- Default active days are weekdays only (weekends excluded unless toggled on).
- Weekly buckets always start on Monday.
- Rows missing valid `Work`, `Start`, or `Finish` are skipped.
- If `Start > Finish`, dates are swapped and still processed.
- Unknown/missing grouping values are normalized to `Unassigned` (resource) or `Unspecified` (project).
