// Talking to Microsoft: sign-in (OAuth) and the Graph API for mail + calendar.

// "consumers" = personal Microsoft accounts (@outlook.com, @hotmail.com, @live.com).
// Change to "common" if you ever want work/school accounts too.
const AUTHORITY = "https://login.microsoftonline.com/consumers/oauth2/v2.0";
const GRAPH = "https://graph.microsoft.com/v1.0";

export const MS_SCOPES = [
  "openid",
  "profile",
  "email",
  "offline_access", // lets us get a refresh token so you stay signed in
  "User.Read",
  "Mail.ReadWrite",
  "Mail.Send",
  "Calendars.ReadWrite",
].join(" ");

export interface MsTokens {
  accessToken: string;
  refreshToken: string;
  expiresAt: number; // unix seconds
}

export function microsoftAuthorizeUrl(opts: {
  clientId: string;
  redirectUri: string;
  state: string;
  codeChallenge: string;
}): string {
  const url = new URL(`${AUTHORITY}/authorize`);
  url.searchParams.set("client_id", opts.clientId);
  url.searchParams.set("response_type", "code");
  url.searchParams.set("redirect_uri", opts.redirectUri);
  url.searchParams.set("response_mode", "query");
  url.searchParams.set("scope", MS_SCOPES);
  url.searchParams.set("state", opts.state);
  url.searchParams.set("code_challenge", opts.codeChallenge);
  url.searchParams.set("code_challenge_method", "S256");
  url.searchParams.set("prompt", "select_account");
  return url.toString();
}

export class MicrosoftTokenError extends Error {
  constructor(
    public readonly code: string,
    message: string,
    public readonly status: number,
  ) {
    super(message);
  }
}

async function tokenRequest(clientId: string, clientSecret: string, params: Record<string, string>): Promise<MsTokens> {
  const res = await fetch(`${AUTHORITY}/token`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({ client_id: clientId, client_secret: clientSecret, scope: MS_SCOPES, ...params }),
  });
  const json = (await res.json().catch(() => ({}))) as Record<string, any>;
  if (!res.ok || !json.access_token) {
    throw new MicrosoftTokenError(
      String(json.error ?? "unknown_error"),
      String(json.error_description ?? `Microsoft token request failed (${res.status})`),
      res.status,
    );
  }
  return {
    accessToken: json.access_token,
    refreshToken: json.refresh_token ?? params.refresh_token ?? "",
    expiresAt: Math.floor(Date.now() / 1000) + Number(json.expires_in ?? 3600),
  };
}

export function exchangeCode(clientId: string, clientSecret: string, code: string, redirectUri: string, verifier: string) {
  return tokenRequest(clientId, clientSecret, {
    grant_type: "authorization_code",
    code,
    redirect_uri: redirectUri,
    code_verifier: verifier,
  });
}

export function refreshTokens(clientId: string, clientSecret: string, refreshToken: string) {
  return tokenRequest(clientId, clientSecret, { grant_type: "refresh_token", refresh_token: refreshToken });
}

// ---------- Graph API ----------

export class GraphError extends Error {
  constructor(
    public readonly status: number,
    message: string,
  ) {
    super(message);
  }
}

export interface GraphOptions {
  method?: string;
  query?: Record<string, string | undefined>;
  body?: unknown;
  headers?: Record<string, string>;
}

export async function graph<T = any>(accessToken: string, path: string, opts: GraphOptions = {}): Promise<T> {
  const url = new URL(GRAPH + path);
  for (const [k, v] of Object.entries(opts.query ?? {})) if (v !== undefined) url.searchParams.set(k, v);
  const res = await fetch(url, {
    method: opts.method ?? "GET",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      ...(opts.body !== undefined ? { "Content-Type": "application/json" } : {}),
      ...opts.headers,
    },
    body: opts.body !== undefined ? JSON.stringify(opts.body) : undefined,
  });
  if (res.status === 202 || res.status === 204) return undefined as T;
  const text = await res.text();
  const json = text ? JSON.parse(text) : undefined;
  if (!res.ok) {
    const msg = json?.error?.message ?? text ?? res.statusText;
    if (res.status === 401) throw new GraphError(401, "Your Microsoft sign-in expired. Disconnect and reconnect the connector in Claude.");
    throw new GraphError(res.status, `Microsoft Graph error ${res.status}: ${msg}`);
  }
  return json as T;
}

export async function getMe(accessToken: string): Promise<{ id: string; email: string; name: string }> {
  const me = await graph<{ id: string; mail?: string; userPrincipalName?: string; displayName?: string }>(accessToken, "/me");
  return { id: me.id, email: (me.mail ?? me.userPrincipalName ?? "").toLowerCase(), name: me.displayName ?? "" };
}
