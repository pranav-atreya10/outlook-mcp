// Self-test: loads the server with a fake account and checks the tools work.
// It never talks to Microsoft for real. Run with: npm test
import { createMcpHandler } from "@modelcontextprotocol/server";
import { buildServer } from "../src/tools";

const EXPECTED = [
  "whoami", "search_emails", "read_email", "send_email", "draft_email", "reply_email", "forward_email",
  "update_email", "list_mail_folders", "list_events", "get_event", "create_event", "update_event",
  "delete_event", "stop_repeating_event", "respond_to_event",
];

let failures = 0;
const check = (ok: boolean, name: string, detail = "") => {
  console.log(`${ok ? "ok  " : "FAIL"} ${name}${!ok && detail ? `\n     ${detail}` : ""}`);
  if (!ok) failures++;
};

// Pretend Microsoft is unreachable/rejects us, so nothing real is ever called.
globalThis.fetch = (async () => new Response(JSON.stringify({ error: { message: "fake" } }), { status: 401 })) as typeof fetch;

const h = createMcpHandler(() =>
  buildServer({ email: "x@outlook.com", name: "X", accessToken: "fake", refreshToken: "r", expiresAt: 0 }, "America/New_York"),
);
const post = (body: unknown) =>
  h.fetch(
    new Request("http://localhost/mcp", {
      method: "POST",
      headers: { "content-type": "application/json", accept: "application/json, text/event-stream", "mcp-protocol-version": "2025-06-18" },
      body: JSON.stringify(body),
    }),
  );
const rpc = async (id: number, method: string, params: unknown) => {
  const text = await (await post({ jsonrpc: "2.0", id, method, params })).text();
  return JSON.parse(text.match(/\{.*\}/s)![0]);
};

const init = await post({ jsonrpc: "2.0", id: 1, method: "initialize", params: { protocolVersion: "2025-06-18", capabilities: {}, clientInfo: { name: "t", version: "1" } } });
check(init.status === 200, "server starts", `status ${init.status}`);

const tools: { name: string; description?: string }[] = (await rpc(2, "tools/list", {})).result.tools;
const names = tools.map((t) => t.name);
check(names.length === EXPECTED.length, `has ${EXPECTED.length} tools`, `got ${names.length}: ${names.join(", ")}`);
for (const n of EXPECTED) check(names.includes(n), `tool ${n} exists`);
check(tools.every((t) => (t.description ?? "").length > 10), "every tool has a description");

const who = await rpc(3, "tools/call", { name: "whoami", arguments: {} });
check(JSON.stringify(who).includes("x@outlook.com"), "whoami returns the signed-in email");

const search = await rpc(4, "tools/call", { name: "search_emails", arguments: {} });
check(search.result?.isError === true, "a Microsoft error comes back as a friendly tool error, not a crash");
check(JSON.stringify(search).includes("reconnect"), "expired sign-in tells you to reconnect");

console.log(failures ? `\n${failures} check(s) failed.` : "\nAll tool tests passed.");
process.exit(failures ? 1 : 0);
