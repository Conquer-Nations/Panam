# Pan American Wire MFG — Customer Portal

A single-page web dashboard for [Conquer Nation Logistics](https://conquernation1.sharepoint.com) to share Pan American Wire MFG shipment and inventory data with the customer, with role-based access.

## Features

- **Two-role login**
  - **Developer** — full access to all 8 tabs (Dashboard, Current Inventory, Monthly Invoices, Weekly Invoices, 10 G, 11 G, 14 G, Snapshot History)
  - **Customer** — Dashboard + Current Inventory only
- **Live updates** — polls `data.json` every 30 seconds; flashes a "LIVE" pill when new data is detected
- **Current On Hand** parsed from the latest "Current on hand" block in the Balance Sheet
- **Date filter** — shows data from `2026-01-01` onward (configurable via `MIN_DATE` in `index.html`)
- **Charts** — total-on-hand trend line, received vs shipped bars, gauge composition donut
- **Sortable, searchable tables** for every detail sheet
- Mobile-responsive industrial dark theme

## Files

| File | Purpose |
|---|---|
| `index.html` | The dashboard — single file, opens directly in any browser |
| `data.json` | Extracted workbook data |
| `refresh-data.cmd` | Double-click to regenerate `data.json` from the source `.xlsx` |
| `refresh-data.ps1` | What the .cmd calls |
| `README.txt` | Detailed deployment + change-password notes |

## Quick start

1. Open `index.html` in any browser (or serve the folder with any static host)
2. Log in with the developer or customer password (set near the top of the `<script>` block in `index.html`)
3. To update data: run `refresh-data.cmd` after editing the source workbook, then re-upload `data.json`

## Deployment

The whole site is static — no backend needed. Drop the folder onto:

- [Netlify Drop](https://app.netlify.com/drop) (easiest)
- GitHub Pages
- Vercel
- OneDrive / SharePoint embed
- Any S3 bucket / static host

For production we recommend pairing with server-side auth (Netlify Pro password protection, Cloudflare Access, etc.) since the in-page password check can be inspected by anyone who views source.

## Security note

The two passwords in `index.html` are plain-text strings — anyone who can read the source can see them. This is fine for non-sensitive view-only data behind an obscure URL, but **do not** push this repo public unless you've rotated the passwords or removed them.
