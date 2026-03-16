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
- Revenue tab with shared per-project `Revenue/hour` and `Gross Profit/hour` rates
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
- `src/components/RevenueWorkspace.tsx`: revenue rate editor + revenue/profit charts
- `src/components/PivotPlanningTable.tsx`: interactive planning pivot/editor
- `src/components/ForecastTable.tsx`: sortable weekly summary table
- `src/components/MultiSelectProjects.tsx`: searchable multi-select project filter
- `src/utils/excel.ts`: workbook parsing and column normalization
- `src/utils/planner.ts`: planning data model, week aggregation, overrides
- `src/utils/reportExport.ts`: client-side Excel export helpers (fallback path)
- `src/utils/reportExportApi.ts`: backend chart-export API client
- `src/utils/activeWorkbookApi.ts`: active workbook upload/load API client
- `src/utils/planningStateApi.ts`: shared planning override state API client
- `src/utils/revenueRatesApi.ts`: shared revenue rates API client
- `src/utils/revenue.ts`: revenue/profit calculation utilities
- `api/export-report.py`: Vercel serverless API route for embedded chart export
- `api/workbook-state.py`: Vercel serverless active workbook metadata route
- `api/workbook-file.py`: Vercel serverless active workbook download route
- `api/upload-workbook.py`: Vercel serverless active workbook upload route
- `api/planning-state.py`: Vercel serverless shared planning overrides state route
- `api/revenue-rates.py`: Vercel serverless shared revenue rates state route
- `api/_workbook_store.py`: shared workbook storage helper for `/api` routes
- `backend/export_api.py`: Python API for chart export + persistent active workbook storage
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

## Backend API (Export + Persistent Active Workbook)

The Python backend now handles both:

- Excel report export (`/api/export-report`)
- Active workbook storage used by the app (`/api/upload-workbook`, `/api/workbook-file`, `/api/workbook-state`)
- Shared planning override storage used by the app (`/api/planning-state`)
- Shared revenue rates storage used by the app (`/api/revenue-rates`)

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

3. Keep it running at `http://127.0.0.1:8000` (or your LAN IP if sharing internally).

4. Set frontend API URL in `.env.local`:

```bash
VITE_SHARED_DATA_API_URL=http://127.0.0.1:8000
```

5. In the dashboard, upload a workbook in `Upload or Replace Workbook (.xlsx)`.
   - The backend saves it to the persistent active workbook location.
   - Refreshing the app loads this uploaded workbook again.
   - Other users pointed to the same backend see the same workbook data.

6. In the dashboard, click `Export Report Excel`.
   - If API is running: exported workbook includes embedded chart in `Weekly Capacity Chart`.
   - If API is not running: app falls back to client-side export (chart data sheet without embedded chart object).

### Active workbook storage details

- Stored file paths (default): `backend/shared_store/active_main.xlsx`, `backend/shared_store/active_sales.xlsx`
- Metadata paths (default): `backend/shared_store/manifests/main.json`, `backend/shared_store/manifests/sales.json`
- Storage root can be overridden with: `CAPACITY_SHARED_DATA_DIR`
- Upload size limit (default 30MB) can be overridden with: `CAPACITY_MAX_UPLOAD_BYTES`
- Planning override entry limit (default 100000) can be overridden with: `CAPACITY_MAX_PLANNING_OVERRIDES`
- Per-cell override max hours (default 1000000) can be overridden with: `CAPACITY_MAX_OVERRIDE_HOURS`
- Revenue rates project limit (default 10000) can be overridden with: `CAPACITY_MAX_RATE_PROJECTS`
- Revenue/Gross Profit per-hour max value (default 1000000) can be overridden with: `CAPACITY_MAX_RATE_PER_HOUR`
- Optional frontend input cap can be set with: `VITE_MAX_RATE_PER_HOUR`

### Active workbook API endpoints

- `GET /api/workbook-state?dataset=main|sales`
- `GET /api/workbook-file?dataset=main|sales`
- `POST /api/upload-workbook?dataset=main|sales`
- `GET /api/planning-state?dataset=main|sales`
- `POST /api/planning-state?dataset=main|sales`
- `GET /api/revenue-rates?dataset=main|sales`
- `POST /api/revenue-rates?dataset=main|sales`
- `GET /api/shared-health`

## API Routes on Public Deploy (Vercel)

This repo includes serverless Python routes in `/api`, including:

- `/api/export-report`
- `/api/workbook-state`
- `/api/workbook-file`
- `/api/upload-workbook`
- `/api/planning-state`
- `/api/revenue-rates`

So the deployed host can serve workbook upload/load endpoints directly without requiring a separate API URL.

Important limitation for serverless storage:
- If `BLOB_READ_WRITE_TOKEN` is set and the Python `vercel` package is installed, `/api` workbook routes store data in Vercel Blob (durable shared storage).
- If Blob is not configured, routes fall back to serverless runtime filesystem (`/tmp`), which can reset on cold starts/redeploys.
- For fully reliable persistence outside Blob, run `backend/export_api.py` with a stable `CAPACITY_SHARED_DATA_DIR`.

## Deploying the App Publicly

The frontend can still be deployed statically, but persistent workbook replacement requires the Python backend to be running at a stable URL.

If you deploy only static frontend files, uploads cannot replace the active workbook persistently for all users.

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
