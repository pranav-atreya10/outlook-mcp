# Outlook connector for Claude (personal accounts)

Lets Claude read, search, send and reply to your Outlook email, and read, add, change and delete your Outlook calendar events. It works with @outlook.com, @hotmail.com and @live.com accounts. It runs on **Vercel** at **mcp.atreyapranav.com**.

**15 tools:** whoami, search_emails, read_email, send_email, draft_email, reply_email, forward_email, update_email, list_mail_folders, list_events, get_event, create_event, update_event, delete_event, respond_to_event

---

## Setup (about 15 minutes)

### 1. Put it on Vercel
1. On vercel.com, click **Add New → Project** and import this GitHub repo (`outlook-mcp`). Keep the default settings and click **Deploy**. The first deploy will show an error page until the next steps are done. That's normal.
2. In the project, open **Storage → Create Database → Upstash (Redis)**. Pick the free plan and connect it to this project. This is where sign-ins are kept.
3. Open **Settings → Domains**, add `mcp.atreyapranav.com`. Your domain is already on Vercel, so it should connect by itself.

### 2. Register an app with Microsoft
1. Go to **https://entra.microsoft.com** and sign in with your Outlook account. (If that doesn't work, use **https://portal.azure.com**, then App registrations.)
2. Open **App registrations**, then click **New registration**:
   - **Name:** `Claude Outlook connector`
   - **Supported account types:** **Personal Microsoft accounts only**
   - **Redirect URI:** choose **Web** and enter `https://mcp.atreyapranav.com/callback`
3. Click **Register**. Copy the **Application (client) ID**.
4. Go to **Certificates & secrets**, then **New client secret**. Pick 24 months, click Add, and copy the **Value** right away. It only shows once.
5. Go to **API permissions**, then **Add a permission**, then **Microsoft Graph**, then **Delegated**. Add:
   `Mail.ReadWrite`, `Mail.Send`, `Calendars.ReadWrite`, `User.Read`, `offline_access`

### 3. Add your settings in Vercel
Go to **Settings → Environment Variables** and add:

| Name | Value |
|---|---|
| `MS_CLIENT_ID` | the Application (client) ID |
| `MS_CLIENT_SECRET` | the secret Value (mark it **Sensitive**) |
| `ALLOWED_EMAILS` | your Outlook email, so only you can sign in |
| `PUBLIC_URL` | `https://mcp.atreyapranav.com` |
| `TIME_ZONE` | `America/New_York` (optional) |

Then go to **Deployments**, open the latest one, click **⋯ → Redeploy**.

Open **https://mcp.atreyapranav.com**. You should see "Outlook connector is running ✅".

### 4. Add it to Claude
1. In Claude, go to **Settings → Connectors → Add custom connector**.
2. Name: `Outlook`. URL: `https://mcp.atreyapranav.com/mcp`
3. Click **Connect**, click **Continue to Microsoft**, sign in, and accept.

Try asking: *"What's on my Outlook calendar this week?"*

From now on, every push to GitHub goes live on its own.

---

## If something breaks
- **"redirect_uri mismatch" from Microsoft:** the redirect URI must be exactly `https://mcp.atreyapranav.com/callback`.
- **"Storage is missing":** connect Upstash Redis to the project (step 1.2), then redeploy.
- **"This account isn't allowed":** fix `ALLOWED_EMAILS`, then redeploy.
- **Tools stop working:** disconnect and reconnect the connector in Claude.
- **See errors:** Vercel project → **Logs**.

## Testing (on your computer)
```sh
npm install
npm run test:e2e    # full sign-in + tools test with a FAKE Microsoft
npm run test:tools  # checks all 15 tools load
```

## How it works
- `app/[[...path]]/route.ts` sends every request to `lib/server.ts`.
- `lib/server.ts` handles sign-in (OAuth) using Cloudflare's open-source OAuth library, and serves the MCP tools at `/mcp`.
- `lib/kv.ts` stores sign-ins in Upstash Redis. Microsoft tokens are encrypted, so the database alone can't read them.
- `lib/tools.ts` has the 15 tools, which call Microsoft Graph (`lib/microsoft.ts`).
