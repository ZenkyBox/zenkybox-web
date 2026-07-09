# First Admin Unlock — Step-by-Step

## Once your app is deployed and live:

### Step 1: Open your app
Go to your live URL (e.g. `https://zenkybox-web-2295.vercel.app`).
By default, every visitor — including you — starts in **Staff mode**.

### Step 2: Find the unlock control
Look at the **bottom of the purple sidebar**, just above "Thoughtful Gifts. Joyful Moments."
You'll see:
```
🔒 Staff mode — unlock admin
```

### Step 3: Click it, enter the default PIN
A small password box appears. Type:
```
2468
```
Press **Enter** or click **Unlock**.

### Step 4: Confirm you're in
- The sidebar footer now shows: `🔓 Admin mode` with a **Lock** button
- New tabs appear that weren't visible before: **Bulk Import, Costing & Pricing, Source Data, Access Management**

### Step 5: Change the PIN immediately
1. Click **Access Management** in the sidebar
2. Scroll to **"Change Admin PIN"**
3. Enter a new PIN (4+ characters) in both boxes
4. Click **Update PIN**
5. You'll see a success toast — the default `2468` no longer works from this point on

**Do this before sharing your app link with anyone.** Anyone who knows the PIN gets full admin access on whatever device they're using.

---

## How staff will experience it

When your team opens the same URL:
- They see Dashboard, Combo Readiness, SKU Catalog, Gift Combos, Upload Sales, Reports, and ZenkyBox Sales Report
- In Catalog/Combos, they can **add** new items but there's no pencil/trash icon to edit or delete existing ones
- No "Clear All" buttons anywhere
- No Bulk Import, Costing & Pricing, Source Data, or Access Management in their sidebar at all

They never see a PIN prompt unless they click "unlock admin" themselves — which they can do too, if they know your PIN. Give the PIN only to people you want to have full control.

---

## Important limits of this system (please read)

This is a **PIN gate**, not real user accounts:
- Everyone shares the same "admin" identity — the activity log will show `role: admin` for anyone who unlocked, not *which* specific person
- Unlocking is per-browser-tab-session (closing the tab locks it again automatically) — this is intentional, so admin access doesn't stay open forever on a shared computer
- There's no password recovery — if you forget the new PIN, you'd need to reset it directly in your Supabase database (or ask for help clearing just that one field)

If you eventually want real named logins (e.g., "Priya added this SKU" instead of just "admin"), that requires adding Supabase Auth — a separate, bigger upgrade. Let me know if you want that built next.

---

## Locking back to staff mode

Click **"Lock"** next to "Admin mode" in the sidebar footer any time you want to hand the device back to someone in staff mode, without them needing to know anything changed.
