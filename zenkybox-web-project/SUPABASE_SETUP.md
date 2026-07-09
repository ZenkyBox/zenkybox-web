# ZenkyBox — Supabase Setup (Cross-Device Sync, 100% Free)

Supabase's free tier requires **no credit card** and **cannot auto-bill you** — if you ever exceed limits, the project simply pauses (your data is safe). For ZenkyBox's data size, you'll use less than 1% of the free tier.

**Free tier includes:** 500 MB database · 5 GB bandwidth/month · Real-time sync · Unlimited API requests

---

## Step 1: Create Supabase Account (2 minutes)

1. Go to: **https://supabase.com**
2. Click **"Start your project"**
3. Sign in with **GitHub** (you already have an account!)
4. No payment info asked — ever.

---

## Step 2: Create a Project (2 minutes)

1. Click **"New project"**
2. Name: `zenkybox`
3. Database Password: create any strong password (save it somewhere)
4. Region: **South Asia (Mumbai)** — closest to you
5. Click **"Create new project"**
6. Wait ~2 minutes while it provisions

---

## Step 3: Create the Data Table (1 minute)

1. In your project, click **"SQL Editor"** (left sidebar)
2. Click **"New query"**
3. Paste this exactly:

```sql
create table zenkybox_data (
  id int primary key,
  data jsonb,
  updated_at timestamptz default now()
);

-- Allow public read/write (single-user app)
alter table zenkybox_data enable row level security;

create policy "Allow all access"
  on zenkybox_data for all
  using (true)
  with check (true);

-- Enable realtime
alter publication supabase_realtime add table zenkybox_data;
```

4. Click **"Run"** (bottom right)
5. You should see: `Success. No rows returned` ✅

---

## Step 4: Get Your API Keys (1 minute)

1. Click **⚙️ Settings** (left sidebar, bottom)
2. Click **"API"**
3. Copy these two values:

| What | Where to find it |
|------|------------------|
| **Project URL** | Top of page — looks like `https://abcdefgh.supabase.co` |
| **anon public key** | Under "Project API keys" — the long string labeled `anon` `public` |

⚠️ Copy the **anon** key, NOT the `service_role` key.

---

## Step 5: Add Keys to Vercel (2 minutes)

1. Go to: **https://vercel.com/dashboard**
2. Click your **zenkybox-web** project
3. **Settings** → **Environment Variables**
4. Add these two:

| Key | Value |
|-----|-------|
| `NEXT_PUBLIC_SUPABASE_URL` | your Project URL |
| `NEXT_PUBLIC_SUPABASE_ANON_KEY` | your anon public key |

5. Click **Save**
6. Go to **Deployments** tab → click **⋯** on latest → **Redeploy**

---

## Step 6: Verify Sync Works ✅

1. Open your app on your **computer** — sidebar should show green dot **"Synced across devices"**
2. Add a test SKU
3. Open the same URL on your **phone**
4. The SKU appears instantly — real-time sync is live! 🎉

---

## Troubleshooting

**Grey dot "Local only"?**
- Both env variables must be set in Vercel, then **redeploy**
- Variable names must start with `NEXT_PUBLIC_` exactly

**Data not syncing between devices?**
- Verify Step 3 SQL ran successfully (check Table Editor → zenkybox_data exists)
- Make sure realtime was enabled (last line of the SQL)

**"Row level security" error in console?**
- Re-run the policy part of the SQL in Step 3

**Project paused after inactivity?**
- Free-tier projects pause after 1 week of no use — just click "Restore" in the Supabase dashboard (takes 1 minute, data is untouched)

---

## Why Supabase over Firebase?

| | Supabase Free | Firebase Spark |
|---|---|---|
| Credit card required | Never | Never |
| Can auto-bill you | No — pauses instead | No — pauses instead |
| Database | 500 MB Postgres | 1 GB NoSQL |
| Real-time sync | ✅ | ✅ |
| Open source | ✅ | ❌ |
| Dashboard to view data | Full SQL table editor | JSON tree |

Both are genuinely free for your scale. Supabase gives you a proper table editor to see your inventory data directly.
