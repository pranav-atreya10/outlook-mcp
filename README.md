# Outlook connector for Claude (personal accounts)

An [MCP](https://modelcontextprotocol.io) server that lets Claude read, search, send and reply to your Outlook email, and read, add, change and delete your Outlook calendar events. It works with personal Microsoft accounts (@outlook.com, @hotmail.com, @live.com) and runs for free on [Cloudflare Workers](https://workers.cloudflare.com).

Each person hosts their own copy, so your email never passes through anyone else's server.

**16 tools:** whoami, search_emails, read_email, send_email, draft_email, reply_email, forward_email, update_email, list_mail_folders, list_events, get_event, create_event, update_event, delete_event, stop_repeating_event, respond_to_event

> **Note:** `wrangler.jsonc` in this repo holds the author's own settings (domain, storage ID, allowed email). Swap in your own in step 2 below before you deploy.

---

## Set up your own copy (about 15 minutes)

You'll need [Node.js](https://nodejs.org) and a free Cloudflare account.

```sh
git clone https://github.com/pranav-atreya10/outlook-mcp
cd outlook-mcp
npm install
npx wrangler login     # opens your browser to sign in to Cloudflare
```

### 1. Pick your address
Your connector needs a web address. Either:
- **Use a free workers.dev address** (easiest): it will look like `https://outlook-mcp.YOUR-NAME.workers.dev`, or
- **Use your own domain** if it's already on Cloudflare, like `https://outlook-mcp.example.com`.

Below, `YOUR-URL` means whichever address you picked.

### 2. Edit `wrangler.jsonc`
1. Make the storage the connector needs, and copy the `id` it prints:
   ```sh
   npx wrangler kv namespace create OAUTH_KV
   ```
   Paste that `id` into `kv_namespaces`.
2. Set your address:
   - **workers.dev:** delete the `routes` line and set `"workers_dev": true`.
   - **Your own domain:** change the `pattern` in `routes` to your domain.
3. Put your Outlook address in `ALLOWED_EMAILS` (for example `"you@outlook.com"`) so only you can sign in.
4. Change `TIME_ZONE` if you're not on New York time (use a name like `"Europe/London"`).

### 3. Register an app with Microsoft
This gives your connector permission to sign in with Microsoft.

1. Go to **https://entra.microsoft.com** and sign in with your Outlook account. (If it asks you to make a tenant, use **https://portal.azure.com** instead and go to App registrations.)
2. Open **App registrations**, then click **New registration**.
3. Fill it in:
   - **Name:** `Claude Outlook connector`
   - **Supported account types:** **Personal Microsoft accounts only**
   - **Redirect URI:** choose **Web** and enter `YOUR-URL/callback`
4. Click **Register**. Copy the **Application (client) ID**.
5. Go to **Certificates & secrets**, then **New client secret**. Pick 24 months, click Add, and copy the **Value** right away. It only shows once.
6. Go to **API permissions**, then **Add a permission**, then **Microsoft Graph**, then **Delegated**. Add:
   `Mail.ReadWrite`, `Mail.Send`, `Calendars.ReadWrite`, `User.Read`, `offline_access`

### 4. Put it online
```sh
npx wrangler secret put MS_CLIENT_ID      # paste the Application (client) ID
npx wrangler secret put MS_CLIENT_SECRET  # paste the secret Value
npm run deploy
```
Wait a minute, then open `YOUR-URL`. You should see "Outlook connector is running ✅".

### 5. Add it to Claude
1. In Claude, go to **Settings → Connectors → Add custom connector**.
2. Name: `Outlook`. URL: `YOUR-URL/mcp`
3. Click **Connect**, click **Continue to Microsoft**, sign in, and accept.

Done. Try asking: *"What's on my Outlook calendar this week?"*

---

## If something breaks
- **"redirect_uri mismatch" from Microsoft:** the redirect URI must be exactly `YOUR-URL/callback`.
- **Deploy says the domain isn't on your account:** make sure your domain shows up under **Websites** in the same Cloudflare account you logged in to, or switch to a workers.dev address.
- **"This account isn't allowed":** fix the email in `ALLOWED_EMAILS` and run `npm run deploy` again.
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
| `test/listtools.ts` | A quick self-test |

## Testing
```sh
npm run typecheck
npm test
```
`npm test` checks that all the tools load. It uses a fake account, so it won't touch your real email.

## Safety notes
- Your Microsoft tokens are stored encrypted on Cloudflare. Only your connector can unlock them.
- Never share `MS_CLIENT_SECRET` or commit `.dev.vars`.
- Keep `ALLOWED_EMAILS` set. If it's empty, anyone who finds your address could connect their own account to it.

## Contributing
Help is welcome! See [CONTRIBUTING.md](CONTRIBUTING.md) for how to get started and some ideas for first changes.

## License
[MIT](LICENSE) © 2026 Pranav Atreya
