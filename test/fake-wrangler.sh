#!/bin/sh
# Pretends to be wrangler so test/setup.test.sh can run the setup script offline.
# Logs every call to $FAKE_LOG.
echo "$*" >> "$FAKE_LOG"
case "$*" in
  whoami) echo "You are logged in with an OAuth Token, associated with the email test@example.com." ;;
  "kv namespace list") echo '[{"id":"11111111111111111111111111111111","title":"something-else"}]' ;;
  "kv namespace create outlook-mcp-oauth") printf 'Creating namespace\n{\n  "kv_namespaces": [{ "binding": "outlook_mcp_oauth", "id": "abcdefabcdefabcdefabcdefabcdef12" }]\n}\n' ;;
  deploy) echo "Uploaded outlook-mcp"; echo "Deployed outlook-mcp triggers"; echo "  https://outlook-mcp.tester.workers.dev" ;;
  "secret list") echo "[]" ;;
  "secret put "*) read -r value; echo "$3=$value" >> "$FAKE_LOG.secrets"; echo "Success! Uploaded secret $3" ;;
  *) echo "unknown command: $*" >&2; exit 1 ;;
esac
