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
- `src/utils/csv.ts`: CSV export helpers
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

## Deploying the App Publicly

This app is a static frontend and is ready for one-click hosting on Vercel, Netlify, or GitHub Pages.

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
