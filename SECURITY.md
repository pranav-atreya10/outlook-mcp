# Security

This connector handles access to people's email, so security reports are taken seriously.

## Reporting a problem
Please **don't** open a public issue. Report it privately here:
https://github.com/pranav-atreya10/outlook-mcp/security/advisories/new

Say what the problem is and how to reproduce it. You'll get a reply within a week.

## How your data is protected
- **You host your own copy.** Your email only flows between Microsoft, your Cloudflare Worker, and Claude. Nobody else's server is involved, including the author's.
- **Only your email can sign in.** `ALLOWED_EMAILS` is required. If it's empty, nobody can sign in.
- **Tokens are encrypted.** Microsoft tokens are stored encrypted in your Cloudflare KV by `@cloudflare/workers-oauth-provider`, and can only be unlocked with a token Claude holds.
- **Sign-in uses PKCE** and a consent page that can't be loaded inside another site's frame.
- **Your secrets stay secret.** `MS_CLIENT_SECRET` is stored as a Cloudflare secret, never in the repo. `wrangler.jsonc` (with your email) is git-ignored.

## Revoking access
- In Claude: Settings → Connectors → disconnect Outlook.
- In Microsoft: https://account.live.com/consent/Manage → remove "Claude Outlook connector".
