/**
 * ZenkyBox Database Layer — Supabase Edition
 * - Supabase (free tier) for cross-device real-time sync
 * - Falls back to localStorage automatically if not configured
 * - No credit card ever required; free tier cannot auto-bill
 *
 * CRITICAL: loadData() distinguishes "no data exists yet" from "fetch failed" —
 * these must never be treated the same, or a transient network/Supabase hiccup
 * (most likely right after a fresh deploy) causes the app to silently overwrite
 * real remote data with an empty local state.
 */

const LOCAL_KEY = "zenkybox-web-v3";

let supabase = null;

async function getSupabase() {
  if (supabase) return supabase;
  const url = process.env.NEXT_PUBLIC_SUPABASE_URL;
  const key = process.env.NEXT_PUBLIC_SUPABASE_ANON_KEY;
  if (!url || !key) return null;
  try {
    const { createClient } = await import("@supabase/supabase-js");
    supabase = createClient(url, key);
    return supabase;
  } catch (e) {
    console.warn("Supabase client could not initialize", e);
    return null;
  }
}

/* ─── READ ───
   Returns { ok, data, source }.
   ok=true  → safe to trust `data` (may legitimately be null if this is a
              brand-new install with nothing saved yet).
   ok=false → the fetch itself FAILED (network/Supabase error). Caller must
              NOT proceed to auto-save in this case — data may exist remotely
              and simply couldn't be retrieved right now. */
export async function loadData() {
  const hasSupabaseConfigured = !!(
    process.env.NEXT_PUBLIC_SUPABASE_URL && process.env.NEXT_PUBLIC_SUPABASE_ANON_KEY
  );

  if (hasSupabaseConfigured) {
    try {
      const sb = await getSupabase();
      if (!sb) {
        // Configured but client failed to initialize — treat as a failed fetch,
        // NOT as "no data". Do not let the caller assume empty state.
        return { ok: false, data: null, source: "supabase-init-failed" };
      }
      const { data, error } = await sb
        .from("zenkybox_data")
        .select("data, updated_at")
        .eq("id", 1)
        .single();

      // PGRST116 = "no rows found" — this IS a legitimate empty state
      // (first-ever run, table exists but nothing saved yet).
      if (error && error.code !== "PGRST116") {
        return { ok: false, data: null, source: "supabase-error", error };
      }
      return { ok: true, data: data?.data ?? null, updatedAt: data?.updated_at ?? null, source: "supabase" };
    } catch (e) {
      // Any thrown exception (network failure, timeout, etc.) — genuinely
      // unknown state. Refuse to let the caller treat this as "empty".
      return { ok: false, data: null, source: "supabase-exception", error: e };
    }
  }

  // No Supabase configured at all — local-only mode. Empty localStorage here
  // truly does mean "nothing saved yet on this device", which is safe.
  try {
    const raw = localStorage.getItem(LOCAL_KEY);
    return { ok: true, data: raw ? JSON.parse(raw) : null, source: "local" };
  } catch (e) {
    return { ok: true, data: null, source: "local-error" };
  }
}

/* ─── WRITE ───
   Returns the updated_at timestamp actually used for this write, so the caller
   can remember "the newest version I myself am responsible for" and use it to
   ignore stale/self-echoed realtime events (see subscribeToData below).

   CRITICAL: this MERGES with whatever is currently saved remotely rather than
   blindly overwriting it. Without this, an older cached browser tab — one
   built before a newer field (e.g. investors/expenses/income) existed in the
   payload shape — would silently WIPE that field the moment it saved anything,
   since it simply doesn't know to include it. Merging means a field is only
   ever removed when the caller explicitly includes it as empty/absent in
   THEIR OWN payload, never as a side effect of a different, older client. */
export async function saveData(payload) {
  const timestamp = new Date().toISOString();

  try {
    localStorage.setItem(LOCAL_KEY, JSON.stringify(payload));
  } catch (e) {}

  try {
    const sb = await getSupabase();
    if (sb) {
      let merged = payload;
      try {
        const { data: existing } = await sb
          .from("zenkybox_data")
          .select("data")
          .eq("id", 1)
          .single();
        if (existing?.data) merged = { ...existing.data, ...payload };
      } catch (e) {
        // No existing row (first-ever save) — just write the payload as-is.
      }
      await sb
        .from("zenkybox_data")
        .upsert({ id: 1, data: merged, updated_at: timestamp });
    }
  } catch (e) {
    console.error("saveData: Supabase write failed", e);
  }

  return timestamp;
}

/* ─── REAL-TIME SUBSCRIPTION ───
   Passes (data, updatedAt) to the callback. The caller is responsible for
   comparing updatedAt against the timestamp of its own most recent saveData()
   call and ignoring anything that isn't genuinely newer — otherwise a self-echo
   or an out-of-order delivery of an older write can overwrite fresher local
   state (the exact bug that caused "add SKU, count flickers back down"). */
export async function subscribeToData(callback) {
  try {
    const sb = await getSupabase();
    if (sb) {
      const channel = sb
        .channel("zenkybox-sync")
        .on(
          "postgres_changes",
          { event: "*", schema: "public", table: "zenkybox_data", filter: "id=eq.1" },
          (payload) => {
            if (payload.new?.data) callback(payload.new.data, payload.new.updated_at);
          }
        )
        .subscribe();
      return () => sb.removeChannel(channel);
    }
  } catch (e) {}
  return null;
}

export const hasSync = () =>
  !!(process.env.NEXT_PUBLIC_SUPABASE_URL && process.env.NEXT_PUBLIC_SUPABASE_ANON_KEY);
