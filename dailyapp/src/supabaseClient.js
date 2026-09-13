import { createClient } from "@supabase/supabase-js";

const SUPABASE_URL = "https://xranddncmkfxxuqssche.supabase.co";
const SUPABASE_KEY = "sb_publishable_Tnam3lbx0S02crLB_kv7hQ_ZSAFGM6i";

export const supabase = createClient(SUPABASE_URL, SUPABASE_KEY);
