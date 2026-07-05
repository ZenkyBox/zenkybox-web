# ZenkyBox v4 — Sales Reports, Source Data & Access Management

## What's New

### 1. ZenkyBox Sales Report (new tab, visible to everyone)
Five report types, each showing SKU details + Combo details + Cost + Earning:
- **Overall** — all-time totals since your first upload
- **Month-over-Month** — grouped by calendar month
- **Weekly** — grouped by Sun–Sat week
- **Buyer-wise** — grouped by buyer name/email
- **Location-wise** — grouped by ship-city + ship-state

Every report card has its own **Export CSV** button.

**Important:** these reports only populate Revenue/Cost/Earning/Buyer/Location once your
Upload Sales file contains those columns. A stock-only file (just sku + quantity) still
updates inventory correctly, but won't show up in Sales Report until you upload a file
with price/buyer/city/state columns too. The Upload Sales screen tells you which columns
it detected on every upload.

### 2. Source Data (admin-only tab)
- Shows your connected database URL (or "not connected — local only" if you haven't set up Supabase)
- **Day-wise Update Record** — every SKU/combo change, import, sales upload, and PIN change,
  grouped by day, with time, action, detail, and who made it (admin/staff)
- Export the whole log to CSV

### 3. Access Management (admin-only tab)
- **Admin**: full control — edit/delete SKUs & combos, Clear All, Bulk Import Replace mode,
  Costing & Pricing, Source Data, Access Management
- **Other Users (Staff)**: can add new SKUs & combos, upload sales reports, view Dashboard,
  Combo Readiness, Reports, and Sales Report — but cannot edit/delete existing records,
  clear data, or see Source Data
- Change the admin PIN from this tab

**Please read this honestly:** this is a lightweight PIN gate stored with your workspace
data — good enough to stop accidental changes by staff, but it is **not** real multi-account
security (no individual logins, no server-side enforcement). Anyone who knows the PIN gets
full admin access on any device. If you need genuine per-person accounts later, the
correct upgrade is Supabase Auth — a bigger addition, ask if you want it built.

**Default PIN is `2468`** until you change it in Access Management. Change this before
sharing your app link with anyone.

---

## Deploying This Version

Same process as every previous update:

```bash
# 1. Extract this ZIP, overwrite your existing project folder
# 2. In that folder:
git add .
git commit -m "v4: Sales Reports, Source Data, Access Management"
git push
```

Vercel auto-deploys in 2–3 minutes. No new environment variables are required —
this version reuses your existing Supabase setup if you have one.

### First thing to do after deploying
1. Open the app → sidebar → click **"Staff mode — unlock admin"**
2. Enter the default PIN `2468`
3. Go to **Access Management** → set your own PIN
4. Share the app link — new visitors start in Staff mode automatically

---

## How to Get Full Data Into Sales Report

Your existing Amazon exports (like the one with the column-shift issue we fixed earlier)
already contain `item-price`, `ship-city`, `ship-state`, and `purchase-date` — the app
now auto-detects these column names. If a future export doesn't show the "detected" banner
on the Upload Sales screen, the column names differ from what we scan for; send a sample
and it can be added.
