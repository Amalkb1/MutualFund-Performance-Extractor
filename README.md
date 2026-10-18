# Fundfolio

A React app for importing Mutual Sahi Hai Excel reports, maintaining a mutual
fund performance master, searching schemes, and exporting results.

## Run locally

```bash
npm install
npm run dev
```

Open the local URL printed by Vite. The bundled dataset in `data/` is loaded on
first use. Imported updates are saved to this browser's local storage.

The master and search tables contain only these selected workbook columns:
Scheme Name, Benchmark, 1/3/5-year Regular and Benchmark returns, and Daily
AUM (Cr.). Import reads the underlying Excel numbers rather than rounded
display text, and table/CSV exports retain their decimal precision.
For the master AUM field, uploads with `Previous Month Average AUM (Cr.)` are
also supported; that source value is stored in the `Daily AUM (Cr.)` column.

## Deploy to Vercel

1. Push this repository to GitHub and import it in Vercel.
2. Keep the project root as the repository root. Vercel reads `vercel.json`;
   the build command is `npm run build` and the output directory is `dist`.
3. Deploy. No Python runtime, environment variables, or API service are needed.

## Data storage

This deployment is a static, client-side React app. Imported or replaced data
is saved in the current browser only; it is not uploaded to Vercel or shared
between visitors, devices, or browsers. Export the master CSV to back it up or
move it to another browser. Shared team data needs a database and authenticated
API, which are not part of this static deployment.

## Optional Python utilities

`build_master.py` and `query_funds.py` remain available for offline workflows.
Install their dependencies with `pip install -r requirements.txt`.
