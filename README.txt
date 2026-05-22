PAN AMERICAN WIRE MFG — CUSTOMER PORTAL
========================================

A single-page web dashboard your customer logs into to view shipment data.

FILES IN THIS FOLDER
--------------------
  index.html         The dashboard itself (open in any browser to test locally)
  data.json          Extracted data from the customer workbook
  refresh-data.cmd   Double-click to regenerate data.json when shipments change
  refresh-data.ps1   Script the .cmd calls
  README.txt         This file


HOW TO USE LOCALLY
------------------
1. Double-click index.html in this folder. It opens in your browser.
2. Two logins, two views:
     Developer password: Developer2025
       -> Full access: Dashboard, Current Inventory, Monthly Invoices,
          Weekly Invoices, 10 G, 11 G, 14 G, Snapshot History
     Customer password:  PanAmerican2025
       -> Customer view: Dashboard + Current Inventory only
3. Live updates: the dashboard polls data.json every 30 seconds.
   When you re-run refresh-data.cmd, the customer browser shows
   the new data automatically (amber "LIVE" pill flashes).


HOW TO CHANGE THE PASSWORDS
---------------------------
Open index.html in Notepad, find this block near the top of the <script> tag:

    const PASSWORDS = {
      developer: "Developer2025",
      customer:  "PanAmerican2025"
    };

Replace either password with your chosen value. Save. Re-upload index.html.

NOTE: this is client-side auth — anyone determined enough can view the page source
and see the passwords. It's appropriate for non-sensitive view-only data. For real
security use one of the hosting options below that adds server-side auth.


HOW TO UPDATE DATA WHEN NEW SHIPMENTS COME IN
----------------------------------------------
1. Open the customer workbook:
     C:\Users\LisaSandoval\Downloads\PAN AMERICAN WIRE MFG NOV 2025 - CUSTOMER.xlsx
   (You'll need to unprotect the sheet to edit — Excel will prompt; there's no password.)
2. Add your new shipment rows to the relevant sheet (INVOICE MONTHLY, 10 G, etc.).
3. Save and close the workbook.
4. Double-click refresh-data.cmd in this folder. It regenerates data.json.
5. Re-upload data.json to your host.


HOW TO DEPLOY (PICK ONE)
------------------------

OPTION A — NETLIFY DROP (easiest, free, takes 2 minutes)
  1. Go to https://app.netlify.com/drop
  2. Drag this entire panam-dashboard folder onto the page.
  3. Netlify gives you a URL like https://random-name-12345.netlify.app
  4. Optionally claim it with a free account to rename it (e.g. panam-portal.netlify.app)
  5. Send the URL + password to your customer.

  To add real password protection (better than the client-side check):
  - In your Netlify site → Site settings → Visitor access → "Password protection"
    (this is a $19/mo Netlify Pro feature — gives proper HTTP auth at the edge)

OPTION B — GITHUB PAGES (free, slightly more setup)
  1. Create a free GitHub account and a new repo (e.g. "panam-portal").
  2. Upload all files from this folder to the repo.
  3. Repo Settings → Pages → Branch: main, folder: / (root). Save.
  4. URL becomes https://your-username.github.io/panam-portal/

OPTION C — ONEDRIVE / SHAREPOINT (use your existing Microsoft 365)
  1. Upload this folder to OneDrive.
  2. Right-click index.html → "Embed" or "Share with link".
  3. Microsoft serves it as a webpage. Same client-side password applies.
  Limitation: OneDrive doesn't give you a clean .com URL — Netlify is nicer for that.

OPTION D — VERCEL (free tier, similar to Netlify)
  1. Sign up at https://vercel.com
  2. "Add New" → "Project" → drag this folder, or connect a GitHub repo.
  3. Deploy. Gets you https://panam-portal.vercel.app or similar.


WHAT THE CUSTOMER SEES
----------------------
- Login screen (password gate)
- Header: Pan American Wire MFG — Shipment Portal
- Tabs: Summary | Monthly Invoices | Weekly Invoices | 10 Gauge | 11 Gauge | 14 Gauge | Balance Sheet
- Summary tab: Big stat tiles (received / shipped / balance), bar chart, donut chart, summary table
- Detail tabs: searchable, sortable tables of every shipment row
- "Customer view — read only" badge


RECOMMENDED STACK
-----------------
For your situation I'd go with Netlify Drop (Option A) for the demo, then if Pan American
loves it and you want real auth, upgrade to Netlify Pro ($19/mo) for password protection,
or build a small Next.js + Auth.js login if you want full control.
