// End-to-end test with a FAKE Microsoft: sign in, get a token, call tools, refresh.
// Run with: npm run test:e2e   (never touches your real account)
process.env.MS_CLIENT_ID = "test-client";
process.env.MS_CLIENT_SECRET = "test-secret";
process.env.ALLOW_MEMORY_KV = "1";
process.env.ALLOWED_EMAILS = "sam@outlook.com";
delete process.env.KV_REST_API_URL;

const B = "https://mcp.example.com";
let msTokenCount = 0;
const realFetch = globalThis.fetch;
const graphCalls: string[] = [];

// Pretend to be Microsoft.
globalThis.fetch = (async (input: any, init?: any) => {
  const url = new URL(typeof input === "string" ? input : input.url ?? String(input));
  if (url.hostname === "login.microsoftonline.com") {
    msTokenCount++;
    return Response.json({ access_token: `ms-access-${msTokenCount}`, refresh_token: `ms-refresh-${msTokenCount}`, expires_in: 3600 });
  }
  if (url.hostname === "graph.microsoft.com") {
    const auth = new Headers(init?.headers).get("authorization");
    graphCalls.push(`${init?.method ?? "GET"} ${url.pathname} (${auth})`);
    if (url.pathname === "/v1.0/me") return Response.json({ id: "user-1", mail: "Sam@outlook.com", displayName: "Sam" });
    if (url.pathname.endsWith("/messages")) return Response.json({ value: [{ id: "m1", subject: "Hi", from: { emailAddress: { name: "A", address: "a@x.com" } }, isRead: false, bodyPreview: "hey" }] });
    if (url.pathname === "/v1.0/me/events") return Response.json({ id: "e1", subject: JSON.parse(init.body).subject, start: JSON.parse(init.body).start, end: JSON.parse(init.body).end }, { status: 201 });
    return Response.json({});
  }
  return realFetch(input, init);
}) as typeof fetch;

const { handle } = await import("../lib/server.ts");

const jar = new Map<string, string>();
async function req(path: string, init: RequestInit = {}) {
  const headers = new Headers(init.headers);
  if (jar.size) headers.set("cookie", [...jar].map(([k, v]) => `${k}=${v}`).join("; "));
  const res = await handle(new Request(B + path, { ...init, headers, redirect: "manual" }));
  for (const c of res.headers.getSetCookie()) {
    const [kv] = c.split(";");
    const [k, ...v] = kv.split("=");
    jar.set(k.trim(), v.join("="));
  }
  return res;
}
const check = (ok: unknown, msg: string) => {
  if (!ok) throw new Error("FAIL: " + msg);
  console.log("✓", msg);
};

// 1. Claude registers itself
const reg = await (await req("/register", {
  method: "POST",
  headers: { "content-type": "application/json" },
  body: JSON.stringify({ client_name: "Claude", redirect_uris: ["https://claude.ai/api/mcp/auth_callback"], token_endpoint_auth_method: "none", grant_types: ["authorization_code", "refresh_token"], response_types: ["code"] }),
})).json();
check(reg.client_id, "client registered");

// 2. Consent page, approve, go to Microsoft
const verifier = "v".repeat(50);
const challenge = Buffer.from(await crypto.subtle.digest("SHA-256", new TextEncoder().encode(verifier))).toString("base64url");
const q = new URLSearchParams({ response_type: "code", client_id: reg.client_id, redirect_uri: "https://claude.ai/api/mcp/auth_callback", state: "s1", code_challenge: challenge, code_challenge_method: "S256", resource: `${B}/mcp` });
const consent = await (await req(`/authorize?${q}`)).text();
const handle1 = consent.match(/name="handle" value="([^"]+)"/)![1];
const toMs = await req("/authorize", { method: "POST", headers: { origin: B, "content-type": "application/x-www-form-urlencoded" }, body: new URLSearchParams({ handle: handle1, decision: "approve" }) });
const msUrl = new URL(toMs.headers.get("location")!);
check(msUrl.hostname === "login.microsoftonline.com", "sent to Microsoft sign-in");

// 3. Microsoft sends you back
const back = await req(`/callback?code=fake-code&state=${msUrl.searchParams.get("state")}`);
const toClaude = new URL(back.headers.get("location")!);
check(toClaude.origin === "https://claude.ai" && toClaude.searchParams.get("code"), "sent back to Claude with a code");

// 4. Claude swaps the code for a token
const tok = await (await req("/token", { method: "POST", headers: { "content-type": "application/x-www-form-urlencoded" }, body: new URLSearchParams({ grant_type: "authorization_code", code: toClaude.searchParams.get("code")!, redirect_uri: "https://claude.ai/api/mcp/auth_callback", client_id: reg.client_id, code_verifier: verifier, resource: `${B}/mcp` }) })).json();
check(tok.access_token && tok.refresh_token, "Claude got a token");

// 5. Use tools
const rpc = async (token: string, id: number, method: string, params: object) => {
  const r = await req("/mcp", { method: "POST", headers: { authorization: `Bearer ${token}`, "content-type": "application/json", accept: "application/json, text/event-stream", "mcp-protocol-version": "2025-06-18" }, body: JSON.stringify({ jsonrpc: "2.0", id, method, params }) });
  const t = await r.text();
  return JSON.parse(t.match(/\{.*\}/s)![0]);
};
await rpc(tok.access_token, 1, "initialize", { protocolVersion: "2025-06-18", capabilities: {}, clientInfo: { name: "t", version: "1" } });
const list = await rpc(tok.access_token, 2, "tools/list", {});
check(list.result.tools.length === 15, `${list.result.tools.length} tools listed`);
const who = await rpc(tok.access_token, 3, "tools/call", { name: "whoami", arguments: {} });
check(who.result.content[0].text.includes("sam@outlook.com"), "whoami shows the signed-in account");
const mail = await rpc(tok.access_token, 4, "tools/call", { name: "search_emails", arguments: {} });
check(mail.result.content[0].text.includes('"subject": "Hi"'), "search_emails works");
const ev = await rpc(tok.access_token, 5, "tools/call", { name: "create_event", arguments: { subject: "Study", start: "2026-10-30T15:00:00", end: "2026-10-30T16:00:00" } });
check(ev.result.content[0].text.includes("America/New_York"), "create_event works with New York time");
check(graphCalls.at(-1)!.includes("ms-access-1"), "used the Microsoft token");

// 6. Refresh: Claude refreshes, we refresh Microsoft too
const tok2 = await (await req("/token", { method: "POST", headers: { "content-type": "application/x-www-form-urlencoded" }, body: new URLSearchParams({ grant_type: "refresh_token", refresh_token: tok.refresh_token, client_id: reg.client_id }) })).json();
check(tok2.access_token, "refresh worked");
await rpc(tok2.access_token, 6, "tools/call", { name: "search_emails", arguments: {} });
check(graphCalls.at(-1)!.includes("ms-access-2"), "new Microsoft token used after refresh");

// 7. Wrong token is rejected
const bad = await req("/mcp", { method: "POST", headers: { authorization: "Bearer nope", "content-type": "application/json" }, body: "{}" });
check(bad.status === 401, "bad token rejected");

console.log("\nAll checks passed.");
