#!/bin/sh
# Runs `npm run setup` against a fake wrangler with typed-in answers, then checks the results.
# Run with: sh test/setup.test.sh
set -e
cd "$(dirname "$0")/.."
TMP=$(mktemp -d)
export FAKE_LOG="$TMP/calls" OUTLOOK_MCP_WRANGLER="$PWD/test/fake-wrangler.sh"
[ -f wrangler.jsonc ] && cp wrangler.jsonc "$TMP/wrangler.jsonc.saved"
restore() { if [ -f "$TMP/wrangler.jsonc.saved" ]; then cp "$TMP/wrangler.jsonc.saved" wrangler.jsonc; else rm -f wrangler.jsonc; fi; rm -f wrangler.jsonc.backup; }
trap restore EXIT

# Answers: a bad email first (should re-ask), then real answers.
printf 'not-an-email\nTest@Outlook.com\nEurope/London\n\n1a2b3c4d-1234-5678-9abc-1234567890ab\nsuper~secret~value~1234567890\n' \
  | node scripts/setup.mjs > "$TMP/out" 2>&1 || { cat "$TMP/out"; echo "FAIL: setup exited with an error"; exit 1; }

check() { if grep -q -- "$2" "$1"; then echo "ok   $3"; else echo "FAIL $3"; cat "$TMP/out"; exit 1; fi; }
check "$TMP/out" "doesn't look like an email" "re-asks on a bad email"
check wrangler.jsonc '"ALLOWED_EMAILS": "test@outlook.com"' "saves the email in lowercase"
check wrangler.jsonc '"TIME_ZONE": "Europe/London"' "saves the time zone"
check wrangler.jsonc '"id": "abcdefabcdefabcdefabcdefabcdef12"' "saves the new storage ID"
check wrangler.jsonc '"workers_dev": true' "uses workers.dev when no domain is given"
check "$FAKE_LOG.secrets" "MS_CLIENT_ID=1a2b3c4d-1234-5678-9abc-1234567890ab" "saves the client ID"
check "$FAKE_LOG.secrets" "MS_CLIENT_SECRET=super~secret~value~1234567890" "saves the client secret"
check "$TMP/out" "https://outlook-mcp.tester.workers.dev/callback" "shows the right redirect URI"
check "$TMP/out" "https://outlook-mcp.tester.workers.dev/mcp" "shows the URL to paste into Claude"
if grep -q "<KV_NAMESPACE_ID>\|<YOUR_EMAIL>\|<TIME_ZONE>" wrangler.jsonc; then echo "FAIL placeholders left"; exit 1; else echo "ok   no placeholders left"; fi

# Second run with a custom domain and --kv-id: should skip storage creation.
: > "$FAKE_LOG"
printf 'me@hotmail.com\nAmerica/Chicago\nhttps://mail.example.com/\n1a2b3c4d-1234-5678-9abc-1234567890ab\nsuper~secret~value~1234567890\n' \
  | node scripts/setup.mjs --kv-id=99999999999999999999999999999999 > "$TMP/out" 2>&1 || { cat "$TMP/out"; exit 1; }
check wrangler.jsonc '"pattern": "mail.example.com"' "uses the custom domain"
check wrangler.jsonc '"workers_dev": false' "turns off workers.dev for a custom domain"
check wrangler.jsonc '"id": "99999999999999999999999999999999"' "reuses --kv-id"
if grep -q "kv namespace create" "$FAKE_LOG"; then echo "FAIL made new storage despite --kv-id"; exit 1; else echo "ok   skips making storage with --kv-id"; fi
check "$TMP/out" "https://mail.example.com/mcp" "shows the custom domain URL"
echo "All setup tests passed."
