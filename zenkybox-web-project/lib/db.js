/**
 * ZenkyBox Database Layer — Supabase Edition
 * - Supabase (free tier) for cross-device real-time sync
 * - Falls back to localStorage automatically if not configured
 * - No credit card ever required; free tier cannot auto-bill
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
    console.warn("Supabase not available, using localStorage", e);
    return null;
  }
}

/* ─── READ ─── */
export async function loadData() {
  try {
    const sb = await getSupabase();
    if (sb) {
      const { data, error } = await sb
        .from("zenkybox_data")
        .select("data")
        .eq("id", 1)
        .single();
      if (!error && data?.data) return data.data;
    }
  } catch (e) {}

  try {
    const raw = localStorage.getItem(LOCAL_KEY);
    if (raw) return JSON.parse(raw);
  } catch (e) {}

  return null;
}

/* ─── WRITE ─── */
export async function saveData(payload) {
  try {
    localStorage.setItem(LOCAL_KEY, JSON.stringify(payload));
  } catch (e) {}

  try {
    const sb = await getSupabase();
    if (sb) {
      await sb
        .from("zenkybox_data")
        .upsert({ id: 1, data: payload, updated_at: new Date().toISOString() });
    }
  } catch (e) {}
}

/* ─── REAL-TIME SUBSCRIPTION ─── */
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
            if (payload.new?.data) callback(payload.new.data);
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
