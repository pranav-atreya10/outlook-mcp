# Outlook connector for Claude (personal accounts)

An [MCP](https://modelcontextprotocol.io) server that lets Claude read, search, send and reply to your Outlook email, and read, add, change and delete your Outlook calendar events. It works with personal Microsoft accounts (@outlook.com, @hotmail.com, @live.com) and runs for free on [Cloudflare Workers](https://workers.cloudflare.com).

Each person hosts their own copy with **one command**, so your email never passes through anyone else's server. Free to run.

[![CI](https://github.com/pranav-atreya10/outlook-mcp/actions/workflows/ci.yml/badge.svg)](https://github.com/pranav-atreya10/outlook-mcp/actions/workflows/ci.yml) [![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)

**What you can ask Claude:**
- *"Summarize my unread emails from today"*
- *"Reply to Jake's email and say I can do Thursday"*
- *"What's on my calendar this week?"*
- *"Add a dentist appointment Friday at 3pm"*
- *"Make a weekly study session every Tuesday at 7"*

**16 tools:** whoami, search_emails, read_email, send_email, draft_email, reply_email, forward_email, update_email, list_mail_folders, list_events, get_event, create_event, update_event, delete_event, stop_repeating_event, respond_to_event

---

## Set up your own copy (about 10 minutes)

You need [Node.js](https://nodejs.org) 18+ and a free [Cloudflare account](https://dash.cloudflare.com/sign-up). Then paste this into a terminal:

```sh
git clone https://github.com/pranav-atreya10/outlook-mcp && cd outlook-mcp && npm install && npm run setup
```

The setup asks for your email and time zone, puts the connector online, and walks you through the one part it can't do for you: registering a free app with Microsoft (about 5 clicks). At the end it prints the link to paste into Claude:

1. In Claude, go to **Settings → Connectors → Add custom connector**.
2. Name: `Outlook`. URL: the link setup gave you (it ends in `/mcp`).
3. Click **Connect**, click **Continue to Microsoft**, sign in, and accept.

Try asking: *"What's on my Outlook calendar this week?"*

Want to change something later (email, time zone, domain)? Just run `npm run setup` again. It reuses your storage and Microsoft app.

<details>
<summary><b>Prefer to set it up by hand?</b></summary>

1. `npx wrangler login`, then `npx wrangler kv namespace create outlook-mcp-oauth` and copy the `id`.
2. Copy `wrangler.example.jsonc` to `wrangler.jsonc` and fill in the `<...>` parts (storage `id`, your time zone, your email). For your own domain, follow the comment in the file.
3. `npm run deploy` and note your address (`YOUR-URL`).
4. At [entra.microsoft.com](https://entra.microsoft.com) → **App registrations** → **New registration**: account type **Personal Microsoft accounts only**, redirect URI **Web** → `YOUR-URL/callback`. Copy the **Application (client) ID**.
5. **Certificates & secrets** → **New client secret** → copy the **Value**.
6. **API permissions** → Microsoft Graph → **Delegated**: `Mail.ReadWrite`, `Mail.Send`, `Calendars.ReadWrite`, `User.Read`, `offline_access`.
7. `npx wrangler secret put MS_CLIENT_ID` and `npx wrangler secret put MS_CLIENT_SECRET`.
8. Add `YOUR-URL/mcp` to Claude as above.
</details>

---

## If something breaks
- **"redirect_uri mismatch" from Microsoft:** the redirect URI in your Microsoft app must be exactly your address + `/callback`.
- **Deploy says the domain isn't on your account:** make sure your domain shows up under **Websites** in the same Cloudflare account you logged in to, or switch to a workers.dev address.
- **"This account isn't allowed":** run `npm run setup` again and enter the email you sign in with.
- **Wrong Microsoft app details:** run `npm run setup -- --reset-microsoft`.
- **Tools stop working after a long time:** disconnect and reconnect the connector in Claude.
- **See errors:** run `npx wrangler tail` while you use it.

## How it works
1. Claude connects to `/mcp` and is told to sign in.
2. You approve on a small page, then sign in with Microsoft.
3. Microsoft sends you back to `/callback`. The connector keeps your Microsoft tokens (encrypted) and gives Claude its own token for this server.
4. Each tool call uses your Microsoft token to talk to Outlook through [Microsoft Graph](https://learn.microsoft.com/graph/overview).

| File | What it does |
| --- | --- |
| `src/index.ts` | Sign-in pages and the OAuth flow |
| `src/microsoft.ts` | Talking to Microsoft's sign-in service |
| `src/tools.ts` | All 16 tools |
| `scripts/setup.mjs` | The one-command setup |
| `test/` | Tests (they never touch real email) |

## Testing
```sh
npm run typecheck
npm test
```
`npm test` checks the tools and the setup script using fake accounts, so it never touches real email. It runs automatically on every pull request.

## Safety notes
- Your Microsoft tokens are stored encrypted on your own Cloudflare account.
- Only the emails you entered can sign in. If none are set, nobody can.
- Never share `MS_CLIENT_SECRET`. Your `wrangler.jsonc` is git-ignored so your email stays out of the repo.
- More in [SECURITY.md](SECURITY.md), including how to report a problem.

## Contributing
Help is welcome! See [CONTRIBUTING.md](CONTRIBUTING.md) for how to get started and some ideas for first changes.

## License
[MIT](LICENSE) © 2026 Pranav Atreya
