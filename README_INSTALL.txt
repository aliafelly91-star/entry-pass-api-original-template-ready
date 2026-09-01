FINAL BUILD — Admin + AI Analytics + Per-Key Usage
===================================================

1) SUPABASE
-----------
Open SQL Editor and run once:
  ai_analytics_final.sql

This creates/updates:
- gemma_api_keys
- ai_usage_logs
- indexes + RLS policies
- per-key analytics support

2) RENDER
---------
Replace main.py in:
  entry-pass-api-original-template-ready

Keep these existing Environment Variables:
  GEMMA_API_KEY                 (fallback key only)
  SUPABASE_SERVICE_ROLE_KEY     (required for Gemma key pool + server analytics)

Then Save / Rebuild / Deploy.

After deploy, Gemma G button behavior is:
- Uses active keys from gemma_api_keys in Supabase, in order.
- If no active dashboard key exists, uses GEMMA_API_KEY from Render as fallback.
- Logs exact key_id, requests, tokens, latency, success/failure.

3) FLUTTER
----------
Replace:
  lib/screens/stage_two_review_screen.dart
  lib/screens/gemini_passport_reader.dart

Keep/add:
  lib/screens/gemma4_passport_reader.dart

Then run:
  flutter run

4) ADMIN DASHBOARD
------------------
Replace the current dashboard HTML with:
  admin_dashboard.html

The final dashboard includes:
- Full FC Barcelona dark/glass theme on login + every admin page.
- Barça crest watermark background.
- AI market-style analytics.
- Time filters: Today, Yesterday, 24h, 7d, 30d, This Month, Previous Month, Custom dates.
- Gemma / Gemini / OpenAI / Hugging Face key management.
- Per-key request count, input/output/total tokens, estimated cost and mini time chart.
- Gemma keys are now managed from the dashboard; Render key remains fallback.

5) IMPORTANT
------------
- Never put SUPABASE_SERVICE_ROLE_KEY inside Flutter or HTML.
- Do not expose full API keys publicly.
- Hugging Face cards are ready to display per-key analytics, but its reader must write key_id/token usage to ai_usage_logs for those numbers to move.
