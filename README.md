# Outlook connector for Claude (personal accounts)

Lets Claude read, search, send and reply to your Outlook email, and read, add, change and delete your Outlook calendar events. It works with @outlook.com, @hotmail.com and @live.com accounts.

**15 tools:** whoami, search_emails, read_email, send_email, draft_email, reply_email, forward_email, update_email, list_mail_folders, list_events, get_event, create_event, update_event, delete_event, respond_to_event

The storage it needs (`OAUTH_KV`) is already set up on your Cloudflare account and filled in for you in `wrangler.jsonc`.

---

## Setup (about 15 minutes)

You'll need Node.js on your computer. Unzip this folder, open a terminal in it, and run `npm install`.

### 1. Register an app with Microsoft
This gives your connector permission to sign in with Microsoft.

1. Go to **https://entra.microsoft.com** and sign in with your Outlook account. (If it asks you to make a tenant, you can also use **https://portal.azure.com**, then go to App registrations.)
2. Open **App registrations**, then click **New registration**.
3. Fill it in:
   - **Name:** `Claude Outlook connector`
   - **Supported account types:** **Personal Microsoft accounts only**
   - **Redirect URI:** choose **Web** and enter `https://outlook-mcp.<your-subdomain>.workers.dev/callback`
     (You'll get the exact address in step 3. You can come back and fix it.)
4. Click **Register**. Copy the **Application (client) ID**.
5. Go to **Certificates & secrets**, then **New client secret**. Pick 24 months, click Add, and copy the **Value** right away. It only shows once.
6. Go to **API permissions**, then **Add a permission**, then **Microsoft Graph**, then **Delegated**. Add:
   `Mail.ReadWrite`, `Mail.Send`, `Calendars.ReadWrite`, `User.Read`, `offline_access`

### 2. Lock it to your email
In `wrangler.jsonc`, put your Outlook address in `ALLOWED_EMAILS`, for example `"sam@outlook.com"`. This way only you can sign in to it. Change `TIME_ZONE` too if you're not on New York time.

### 3. Put it online
```sh
npx wrangler login                        # opens your browser to sign in to Cloudflare
npx wrangler secret put MS_CLIENT_ID      # paste the Application (client) ID
npx wrangler secret put MS_CLIENT_SECRET  # paste the secret Value
npm run deploy
```
When it finishes, it prints your web address, like `https://outlook-mcp.sam123.workers.dev`. Check that the redirect URI in step 1 matches it exactly, with `/callback` on the end.

Open that address in your browser. You should see "Outlook connector is running ✅".

### 4. Add it to Claude
1. In Claude, go to **Settings → Connectors → Add custom connector**.
2. Name: `Outlook`. URL: `https://outlook-mcp.<your-subdomain>.workers.dev/mcp`
3. Click **Connect**, click **Continue to Microsoft**, sign in, and accept.

Done. Try asking: *"What's on my Outlook calendar this week?"*

---

## If something breaks
- **"redirect_uri mismatch" from Microsoft:** the redirect URI in step 1 doesn't exactly match your address + `/callback`.
- **"This account isn't allowed":** fix the email in `ALLOWED_EMAILS` and run `npm run deploy` again.
- **Tools stop working after a long time:** disconnect and reconnect the connector in Claude.
- **See errors:** run `npx wrangler tail` while you use it.

## Testing
`npx tsx test/listtools.ts` checks that all 15 tools load (it uses a fake account, so it won't touch your real email).

## Safety notes
- Your Microsoft tokens are stored encrypted on Cloudflare. Only the connector can unlock them.
- Never share `MS_CLIENT_SECRET` or commit `.dev.vars`.
