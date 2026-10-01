# AutoTask

A web-based automation workspace. The first module updates the Zona Uniform Daily Sales Tracker from supplier Excel reports.

## Current MVP

- Upload the current Uniform Daily Sales Tracker.
- Upload either a complete combined Zona Sales Report or the four individual RDXB, RAB, FRY and ROSE reports.
- Detect source columns by header rather than fixed Excel letters.
- Aggregate daily sales by campus.
- Aggregate Exchange / Shipping Value across the four campuses by date.
- Match each date to the appropriate monthly tracker sheet and date row.
- Return an updated `.xlsx` file without changing the original local workbook.

## Next phase

Add Microsoft 365 sign-in and Microsoft Graph email retrieval for supplier `shanawaz@zonatrading.net`, including flexible DAILY SALES REPORT subject matching and automatic attachment selection.

## Development

```bash
npm install
npm run dev
```

Python dependencies for the Vercel API are listed in `requirements.txt`.
