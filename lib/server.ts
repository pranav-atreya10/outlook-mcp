// Outlook connector for Claude (personal Microsoft accounts).
//
// How it works:
//   1. Claude connects to https://mcp.atreyapranav.com/mcp and gets told to sign in.
//   2. You approve on a small page here, then sign in with Microsoft.
//   3. Microsoft sends you back to /callback; we keep your Microsoft tokens (encrypted)
//      and Claude gets its own token for this server.
//   4. Every tool call uses your Microsoft token to talk to Outlook.

import OAuthProvider, {
  AuthorizationError,
  CimdFetchError,
  GrantType,
  OAuthError,
  type AuthRequest,
  type ConsentDescription,
  type OAuthHelpers,
} from "@cloudflare/workers-oauth-provider";
import { createMcpHandler } from "@modelcontextprotocol/server";
import { waitUntil } from "@vercel/functions";
import { getKV, type KVLike } from "./kv";
import { exchangeCode, getMe, microsoftAuthorizeUrl, MicrosoftTokenError, refreshTokens } from "./microsoft";
import { buildServer, type Props } from "./tools";

interface Env {
  OAUTH_KV: KVLike;
  OAUTH_PROVIDER: OAuthHelpers;
  MS_CLIENT_ID: string;
  MS_CLIENT_SECRET: string;
  /** Comma-separated Microsoft emails allowed to use this connector. Empty = anyone. */
  ALLOWED_EMAILS?: string;
  TIME_ZONE?: string;
}

// ---------------- the MCP endpoint (/mcp) ----------------

const mcpApi = {
  async fetch(request: Request, env: Env, ctx: { props: Props }) {
    const props = ctx.props;
    const handler = createMcpHandler(() => buildServer(props, env.TIME_ZONE || "America/New_York"));
    return handler.fetch(request);
  },
};

// ---------------- sign-in pages ----------------

const escape = (v: string) => v.replace(/[&<>"']/g, (c) => `&#${c.charCodeAt(0)};`);

function page(title: string, inner: string, headers = new Headers(), status = 200) {
  headers.set("Content-Type", "text/html; charset=utf-8");
  const html = `<!doctype html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>${escape(title)}</title>
<style>
 body{font-family:system-ui,-apple-system,sans-serif;background:#f4f5f7;color:#1b1d21;margin:0;display:grid;place-items:center;min-height:100vh;padding:16px;box-sizing:border-box}
 .card{background:#fff;border-radius:14px;box-shadow:0 2px 12px rgba(0,0,0,.08);padding:28px;max-width:420px;width:100%}
 h1{font-size:20px;margin:0 0 12px} p{line-height:1.5;margin:8px 0} ul{padding-left:20px;line-height:1.6}
 .warn{background:#fff4e5;border-radius:8px;padding:10px 12px}
 .row{display:flex;gap:10px;margin-top:20px} button{flex:1;font:inherit;padding:10px;border-radius:8px;border:1px solid #ccd;cursor:pointer;background:#fff}
 button.primary{background:#0f6cbd;border-color:#0f6cbd;color:#fff}
 @media (prefers-color-scheme:dark){body{background:#16181c;color:#e8e9ec}.card{background:#22252b}.warn{background:#3a2e1c}button{background:#2c3038;color:#e8e9ec;border-color:#444}}
</style></head><body><div class="card">${inner}</div></body></html>`;
  return new Response(html, { status, headers });
}

function consentPage(d: ConsentDescription, handle: string, headers: Headers) {
  const name = escape(d.clientName);
  const origin = d.clientDomain
    ? `Published by <strong>${escape(d.clientDomain)}</strong>.`
    : "This app registered itself, so its name is not verified.";
  return page(
    `Connect ${d.clientName} to Outlook`,
    `<h1>Let ${name} use your Outlook?</h1>
<p>${origin} Access goes to <strong>${escape(d.redirectHost)}</strong>.</p>
${d.redirectIsLoopback ? '<p class="warn"><strong>This sends access to an app on your computer.</strong> Only continue if you just started connecting from it.</p>' : ""}
<p>It will be able to:</p>
<ul><li>Read, search, send and reply to your email</li><li>Read, add, change and delete calendar events</li></ul>
<form method="post"><input type="hidden" name="handle" value="${escape(handle)}">
<div class="row"><button name="decision" value="deny">Deny</button><button class="primary" name="decision" value="approve">Continue to Microsoft</button></div></form>`,
    headers,
  );
}

async function s256(verifier: string) {
  const digest = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(verifier));
  return btoa(String.fromCharCode(...new Uint8Array(digest)))
    .replace(/\+/g, "-")
    .replace(/\//g, "_")
    .replace(/=+$/, "");
}

const redirect = (headers: Headers, location: string) => {
  headers.set("Location", location);
  return new Response(null, { status: 302, headers });
};

async function handleAuthorize(req: Request, env: Env): Promise<Response> {
  const oauth = env.OAUTH_PROVIDER;

  if (req.method === "GET") {
    const authReq: AuthRequest = await oauth.parseAuthRequest(req);
    const details = await oauth.describeConsent(authReq);
    const consent = await oauth.beginConsent(authReq);
    return consentPage(details, consent.handle, consent.headers);
  }

  // POST: the user clicked a button
  const form = await req.formData();
  const handle = String(form.get("handle"));
  if (form.get("decision") !== "approve") {
    const denied = await oauth.denyConsent(req, handle);
    return redirect(denied.headers, denied.redirectTo);
  }
  const approved = await oauth.approveConsent(req, handle);
  const verifier = crypto.randomUUID() + crypto.randomUUID();
  const { state, headers } = await oauth.beginUpstream(approved.request, {
    data: { verifier },
    headers: approved.headers,
  });
  return redirect(
    headers,
    microsoftAuthorizeUrl({
      clientId: env.MS_CLIENT_ID,
      redirectUri: new URL("/callback", req.url).toString(),
      state,
      codeChallenge: await s256(verifier),
    }),
  );
}

async function handleCallback(req: Request, env: Env): Promise<Response> {
  const oauth = env.OAUTH_PROVIDER;
  const url = new URL(req.url);
  const { request: original, data, headers } = await oauth.finishUpstream<{ verifier: string }>(req);

  if (url.searchParams.get("error")) {
    // You cancelled on Microsoft's page (or it failed): tell Claude "access denied".
    const r = new URL(original.redirectUri);
    r.searchParams.set("error", "access_denied");
    if (original.state) r.searchParams.set("state", original.state);
    return redirect(headers, r.toString());
  }

  const code = url.searchParams.get("code");
  if (!code) return page("Sign-in failed", "<h1>Sign-in failed</h1><p>Microsoft didn't send a code back. Try connecting again.</p>", headers, 400);

  const tokens = await exchangeCode(env.MS_CLIENT_ID, env.MS_CLIENT_SECRET, code, new URL("/callback", req.url).toString(), data.verifier);
  const me = await getMe(tokens.accessToken);

  const allowed = (env.ALLOWED_EMAILS ?? "")
    .split(",")
    .map((s) => s.trim().toLowerCase())
    .filter(Boolean);
  if (allowed.length && !allowed.includes(me.email)) {
    return page(
      "Not allowed",
      `<h1>This account isn't allowed</h1><p>${escape(me.email)} isn't on this connector's allowed list.</p>`,
      headers,
      403,
    );
  }

  const props: Props = {
    email: me.email,
    name: me.name,
    accessToken: tokens.accessToken,
    refreshToken: tokens.refreshToken,
    expiresAt: tokens.expiresAt,
  };
  const { redirectTo } = await oauth.completeAuthorization({
    request: original,
    userId: me.id,
    metadata: { email: me.email },
    scope: original.scope,
    props,
  });
  return redirect(headers, redirectTo);
}

const defaultHandler = {
  async fetch(req: Request, env: Env): Promise<Response> {
    const { pathname } = new URL(req.url);
    try {
      if (pathname === "/authorize") return await handleAuthorize(req, env);
      if (pathname === "/callback") return await handleCallback(req, env);
      if (pathname === "/")
        return page("Outlook connector", "<h1>Outlook connector is running ✅</h1><p>Add <code>/mcp</code> to this address and paste it into Claude as a custom connector.</p>");
      return new Response("Not found", { status: 404 });
    } catch (error) {
      if (error instanceof AuthorizationError && error.redirectTo) return Response.redirect(error.redirectTo, 302);
      if (error instanceof AuthorizationError || error instanceof CimdFetchError || error instanceof MicrosoftTokenError) {
        const msg =
          error instanceof AuthorizationError
            ? error.description
            : error instanceof MicrosoftTokenError
              ? `Microsoft said: ${error.message}`
              : "This app could not be verified.";
        return page("Something went wrong", `<h1>Something went wrong</h1><p>${escape(msg)}</p><p>Go back to Claude and try connecting again.</p>`, new Headers(), 400);
      }
      throw error;
    }
  },
};

// ---------------- wire it all together ----------------

// The provider needs this server's public address: PUBLIC_URL if set, otherwise the request's.
let provider: OAuthProvider<Env> | undefined;

function getEnv(): Env {
  return {
    OAUTH_KV: getKV(),
    OAUTH_PROVIDER: undefined as unknown as OAuthHelpers, // filled in by the library per request
    MS_CLIENT_ID: process.env.MS_CLIENT_ID ?? "",
    MS_CLIENT_SECRET: process.env.MS_CLIENT_SECRET ?? "",
    ALLOWED_EMAILS: process.env.ALLOWED_EMAILS,
    TIME_ZONE: process.env.TIME_ZONE,
  };
}

/** Handles every request to the site: sign-in pages, OAuth endpoints and /mcp. */
export async function handle(request: Request): Promise<Response> {
  const origin = (process.env.PUBLIC_URL || new URL(request.url).origin).replace(/\/+$/, "");
  // Make sure the library sees the public address even behind Vercel's proxy.
  const url = new URL(request.url);
  const publicUrl = new URL(url.pathname + url.search, origin);
  if (publicUrl.href !== url.href) request = new Request(publicUrl, request);

  provider ??= makeProvider(origin);
  const ctx = { waitUntil, passThroughOnException() {}, props: {} } as any;
  return provider.fetch(request, getEnv() as any, ctx);
}

const makeProvider = (origin: string) => new OAuthProvider<Env>({
  resourceMetadata: {
    resource: `${origin}/mcp`,
    authorization_servers: [origin],
    resource_name: "Outlook (personal)",
  },
  apiRoute: "/mcp",
  apiHandler: mcpApi as any,
  defaultHandler: defaultHandler as any,
  authorizeEndpoint: "/authorize",
  tokenEndpoint: "/token",
  clientRegistrationEndpoint: "/register",
  // Claude signs up through /register (dynamic client registration).
  clientIdMetadataDocumentEnabled: false,

  // Keep Claude's token lifetime in step with Microsoft's, and refresh Microsoft's
  // token whenever Claude refreshes ours. This is what keeps you signed in.
  tokenExchangeCallback: async ({ grantType, props, env }) => {
    const p = props as Props;
    const now = Math.floor(Date.now() / 1000);
    if (grantType === GrantType.AUTHORIZATION_CODE) {
      return { accessTokenTTL: Math.max(60, p.expiresAt - now - 120) };
    }
    if (grantType !== GrantType.REFRESH_TOKEN) return;
    try {
      const t = await refreshTokens(env.MS_CLIENT_ID, env.MS_CLIENT_SECRET, p.refreshToken);
      const newProps: Props = { ...p, accessToken: t.accessToken, refreshToken: t.refreshToken, expiresAt: t.expiresAt };
      return { newProps, accessTokenTTL: Math.max(60, t.expiresAt - now - 120) };
    } catch (e) {
      if (e instanceof MicrosoftTokenError && (e.code === "invalid_grant" || e.code === "interaction_required")) {
        throw new OAuthError("invalid_grant", { description: "Microsoft access was revoked or expired. Reconnect." });
      }
      throw new OAuthError("temporarily_unavailable", { description: "Microsoft sign-in is unavailable right now", statusCode: 503 });
    }
  },
});
