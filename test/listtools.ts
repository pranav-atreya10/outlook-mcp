// Quick self-test: loads the tools with fake sign-in info and lists them.
// Run with: npx tsx test/listtools.ts
import { createMcpHandler } from "@modelcontextprotocol/server";
import { buildServer } from "../src/tools";
const h = createMcpHandler(() => buildServer({ email: "x@outlook.com", name: "X", accessToken: "fake", refreshToken: "r", expiresAt: 0 }, "America/New_York"));
const post = (body: any) => h.fetch(new Request("http://localhost/mcp", { method: "POST", headers: { "content-type": "application/json", accept: "application/json, text/event-stream", "mcp-protocol-version": "2025-06-18" }, body: JSON.stringify(body) }));
const r1 = await post({ jsonrpc: "2.0", id: 1, method: "initialize", params: { protocolVersion: "2025-06-18", capabilities: {}, clientInfo: { name: "t", version: "1" } } });
console.log("initialize:", r1.status);
const t = await (await post({ jsonrpc: "2.0", id: 2, method: "tools/list", params: {} })).text();
const tools = JSON.parse(t.match(/\{.*\}/s)![0]).result.tools;
console.log(tools.length, "tools:", tools.map((x: any) => x.name).join(", "));
console.log("create_event schema:", JSON.stringify(tools.find((x: any) => x.name === "create_event").inputSchema).slice(0, 250));
console.log("whoami:", (await (await post({ jsonrpc: "2.0", id: 3, method: "tools/call", params: { name: "whoami", arguments: {} } })).text()).slice(0, 250));
console.log("search (fake token, expect friendly error):", (await (await post({ jsonrpc: "2.0", id: 4, method: "tools/call", params: { name: "search_emails", arguments: {} } })).text()).slice(0, 250));
