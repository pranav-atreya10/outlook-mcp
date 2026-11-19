#!/usr/bin/env node
// One-command setup: asks a few questions, then puts your own Outlook connector online.
// Run with: npm run setup
//
// Options:
//   --kv-id=<id>   reuse an existing Cloudflare KV namespace instead of making one
//
// Uses only Node's built-in modules, so it works right after `npm install`.

import { spawn } from "node:child_process";
import { existsSync, readFileSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

const ROOT = join(dirname(fileURLToPath(import.meta.url)), "..");
const CONFIG = join(ROOT, "wrangler.jsonc");
const TEMPLATE = join(ROOT, "wrangler.example.jsonc");
const KV_TITLE = "outlook-mcp-oauth";
const ENTRA_URL = "https://entra.microsoft.com/#view/Microsoft_AAD_RegisteredApps/CreateApplicationBlade";

// Tests can point this at a fake wrangler.
const WRANGLER = process.env.OUTLOOK_MCP_WRANGLER ? [process.env.OUTLOOK_MCP_WRANGLER] : ["npx", "wrangler"];
const args = Object.fromEntries(
  process.argv.slice(2).map((a) => {
    const [k, ...v] = a.replace(/^--/, "").split("=");
    return [k, v.join("=") || true];
  }),
);

// ---------- small helpers ----------

const bold = (s) => (process.stdout.isTTY ? `\x1b[1m${s}\x1b[0m` : s);
const green = (s) => (process.stdout.isTTY ? `\x1b[32m${s}\x1b[0m` : s);
const step = (n, s) => console.log(`\n${bold(`Step ${n}:`)} ${s}`);
const fail = (msg) => {
  console.error(`\n❌ ${msg}`);
  process.exit(1);
};

// Read answers from stdin, whether it's a terminal or piped (piped = tests).
// In "hidden" mode the terminal stops echoing and we print * for each character instead.
const stdin = process.stdin;
let buffer = "";
let hiddenMode = false;
let waiter = null;
let ended = false;
stdin.setEncoding("utf8");
stdin.on("data", (chunk) => {
  for (const ch of chunk) {
    if (ch === "\u0003") process.exit(130); // Ctrl+C in raw mode
    if (ch === "\r" || ch === "\n") {
      if (ch === "\n" && buffer === "" && lastWasCR) continue; // \r\n
      lastWasCR = ch === "\r";
      deliver(buffer);
      buffer = "";
      continue;
    }
    lastWasCR = false;
    if (hiddenMode && (ch === "\u007f" || ch === "\b")) {
      if (buffer) {
        buffer = buffer.slice(0, -1);
        process.stdout.write("\b \b");
      }
      continue;
    }
    buffer += ch;
    if (hiddenMode) process.stdout.write("*");
  }
});
stdin.pause(); // only read while we're asking something
stdin.on("end", () => {
  ended = true;
  if (waiter) waiter(buffer || null);
});
let lastWasCR = false;
const queued = [];
function deliver(line) {
  if (waiter) {
    const w = waiter;
    waiter = null;
    w(line);
  } else queued.push(line);
}
const nextLine = () =>
  queued.length ? Promise.resolve(queued.shift()) : ended ? Promise.resolve(null) : new Promise((r) => (waiter = r));

async function ask(question, { fallback = "", hidden = false } = {}) {
  process.stdout.write(`${question}${fallback ? ` (${fallback})` : ""}: `);
  const useRaw = hidden && stdin.isTTY;
  if (useRaw) {
    stdin.setRawMode(true);
    hiddenMode = true;
  }
  stdin.resume();
  const answer = await nextLine();
  if (useRaw) {
    stdin.setRawMode(false);
    hiddenMode = false;
  }
  stdin.pause();
  if (useRaw || !stdin.isTTY) process.stdout.write("\n");
  if (answer === null) fail("Setup was cancelled.");
  return answer.trim() || fallback;
}

// Run wrangler. `live` shows its output as it goes (and lets it ask you things).
function wrangler(cmdArgs, { input, live = false } = {}) {
  return new Promise((resolve) => {
    const [cmd, ...pre] = WRANGLER;
    const child = spawn(cmd, [...pre, ...cmdArgs], {
      cwd: ROOT,
      shell: process.platform === "win32",
      stdio: [input !== undefined ? "pipe" : live ? "inherit" : "ignore", "pipe", "pipe"],
    });
    let out = "";
    const collect = (stream) => (d) => {
      out += d;
      if (live) stream.write(d);
    };
    child.stdout.on("data", collect(process.stdout));
    child.stderr.on("data", collect(process.stderr));
    if (input !== undefined) child.stdin.end(input);
    child.on("error", (e) => resolve({ code: 1, out: String(e) }));
    child.on("close", (code) => resolve({ code, out }));
  });
}

const parseJson = (text) => {
  const start = text.search(/[[{]/);
  try {
    return start >= 0 ? JSON.parse(text.slice(start)) : undefined;
  } catch {
    return undefined;
  }
};

const EMAIL = /^[^\s@,]+@[^\s@,]+\.[^\s@,]+$/;
const GUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

// ---------- the setup ----------

async function main() {
  console.log(bold("\n📬 Outlook connector for Claude — setup\n"));
  console.log("This puts your own copy online on Cloudflare (free). It takes about 10 minutes.");

  const major = Number(process.versions.node.split(".")[0]);
  if (major < 18) fail(`You need Node.js 18 or newer (you have ${process.versions.node}). Get it at https://nodejs.org`);
  if (!existsSync(join(ROOT, "node_modules"))) fail("Run `npm install` first, then `npm run setup` again.");

  // 1. Cloudflare sign-in
  step(1, "Signing in to Cloudflare");
  let who = await wrangler(["whoami"]);
  if (who.code !== 0 || /not authenticated|not logged in/i.test(who.out)) {
    console.log("Your browser will open so you can sign in (or make a free account).");
    await wrangler(["login"], { live: true });
    who = await wrangler(["whoami"]);
    if (who.code !== 0 || /not authenticated|not logged in/i.test(who.out)) fail("Couldn't sign in to Cloudflare. Run `npx wrangler login`, then try again.");
  }
  console.log(green("✓ Signed in to Cloudflare"));

  // 2. Questions
  step(2, "A few questions");
  let email;
  for (;;) {
    email = (await ask("Your Outlook email (the one you'll sign in with)")).toLowerCase();
    const list = email.split(",").map((e) => e.trim());
    if (list.every((e) => EMAIL.test(e))) {
      email = list.join(",");
      break;
    }
    console.log("That doesn't look like an email. Try again (for example you@outlook.com).");
  }
  if (!/@(outlook|hotmail|live|msn)\./.test(email)) {
    console.log("Heads up: this works with personal Microsoft accounts. Work or school emails won't be able to sign in.");
  }
  const guessedZone = Intl.DateTimeFormat().resolvedOptions().timeZone || "America/New_York";
  let timeZone;
  for (;;) {
    timeZone = await ask("Your time zone", { fallback: guessedZone });
    try {
      Intl.DateTimeFormat("en-US", { timeZone });
      break;
    } catch {
      console.log('Unknown time zone. Use a name like "America/New_York" or "Europe/London".');
    }
  }
  const domain = (await ask("Your own domain, like outlook-mcp.example.com (press Enter to use a free workers.dev address)"))
    .replace(/^https?:\/\//, "")
    .replace(/\/.*$/, "")
    .toLowerCase();

  // 3. Storage
  step(3, "Making storage for sign-ins");
  let kvId = typeof args["kv-id"] === "string" ? args["kv-id"] : "";
  if (!kvId) {
    const list = parseJson((await wrangler(["kv", "namespace", "list"])).out);
    kvId = Array.isArray(list) ? (list.find((n) => n.title === KV_TITLE)?.id ?? "") : "";
    if (kvId) console.log("Found storage from an earlier setup. Reusing it.");
  }
  if (!kvId) {
    const made = await wrangler(["kv", "namespace", "create", KV_TITLE]);
    kvId = made.out.match(/\b([0-9a-f]{32})\b/)?.[1] ?? "";
    if (!kvId) fail(`Couldn't make storage. Cloudflare said:\n${made.out}`);
  }
  console.log(green("✓ Storage ready"));

  // 4. Settings file
  if (existsSync(CONFIG)) writeFileSync(`${CONFIG}.backup`, readFileSync(CONFIG));
  let config = readFileSync(TEMPLATE, "utf8")
    .replace("<KV_NAMESPACE_ID>", kvId)
    .replace("<TIME_ZONE>", timeZone)
    .replace("<YOUR_EMAIL>", email);
  if (domain) {
    config = config.replace('"workers_dev": true,', `"workers_dev": false,\n  "routes": [{ "pattern": "${domain}", "custom_domain": true }],`);
  }
  writeFileSync(CONFIG, config);

  // 5. Deploy
  step(4, "Putting it online");
  const deployed = await wrangler(["deploy"], { live: true });
  if (deployed.code !== 0) {
    if (domain) fail(`Deploy failed. Make sure ${domain.split(".").slice(-2).join(".")} is on the same Cloudflare account, or run setup again and press Enter for a free address.`);
    fail("Deploy failed. See the message above, fix it, and run `npm run setup` again.");
  }
  const url = domain ? `https://${domain}` : deployed.out.match(/https:\/\/[a-z0-9.-]+\.workers\.dev/i)?.[0];
  if (!url) fail("It deployed, but I couldn't find its address in the output above. Look for a workers.dev link there.");
  console.log(green(`✓ Online at ${url}`));

  // 6. Microsoft app (skipped if it's already set up)
  const secrets = parseJson((await wrangler(["secret", "list"])).out);
  const have = new Set(Array.isArray(secrets) ? secrets.map((s) => s.name) : []);
  if (have.has("MS_CLIENT_ID") && have.has("MS_CLIENT_SECRET") && !args["reset-microsoft"]) {
    console.log(green("✓ Microsoft app already connected (run with --reset-microsoft to change it)"));
  } else {
    step(5, "Letting it sign in with Microsoft");
    console.log(`
Open ${bold(ENTRA_URL)}
and sign in with ${email.split(",")[0]}. (If it asks you to make a tenant, use https://portal.azure.com → App registrations instead.)

  1. Name: ${bold("Claude Outlook connector")}
  2. Supported account types: ${bold("Personal Microsoft accounts only")}
  3. Redirect URI: pick ${bold("Web")} and paste ${bold(`${url}/callback`)}
  4. Click Register, then copy the ${bold("Application (client) ID")}.
`);
    let clientId;
    for (;;) {
      clientId = await ask("Paste the Application (client) ID");
      if (GUID.test(clientId)) break;
      console.log("That should look like 1a2b3c4d-1234-5678-9abc-1234567890ab. Try again.");
    }
    console.log(`
  5. Go to ${bold("Certificates & secrets")} → ${bold("New client secret")} → 24 months → Add.
     Copy the ${bold("Value")} column (not the Secret ID). It only shows once.
  6. Go to ${bold("API permissions")} → Add a permission → Microsoft Graph → ${bold("Delegated")}, and add:
     Mail.ReadWrite, Mail.Send, Calendars.ReadWrite, User.Read, offline_access
`);
    let clientSecret;
    for (;;) {
      clientSecret = await ask("Paste the secret Value", { hidden: true });
      if (GUID.test(clientSecret)) {
        console.log("That looks like the Secret ID. Copy the Value column instead.");
      } else if (clientSecret.length >= 20) break;
      else console.log("That looks too short. Copy the whole Value.");
    }
    for (const [name, value] of [
      ["MS_CLIENT_ID", clientId],
      ["MS_CLIENT_SECRET", clientSecret],
    ]) {
      const r = await wrangler(["secret", "put", name], { input: value });
      if (r.code !== 0) fail(`Couldn't save ${name}. Cloudflare said:\n${r.out}`);
    }
    console.log(green("✓ Microsoft app connected"));
  }

  console.log(`
${green(bold("🎉 All done!"))}

Add it to Claude:
  1. Open Claude → ${bold("Settings → Connectors → Add custom connector")}
  2. Name: ${bold("Outlook")}   URL: ${bold(`${url}/mcp`)}
  3. Click Connect, then Continue to Microsoft, and sign in as ${email.split(",")[0]}.

Then try: "What's on my Outlook calendar this week?"
Stuck? See the README, or run \`npx wrangler tail\` to watch for errors.
`);
  process.exit(0);
}

main().catch((e) => fail(e?.stack ?? String(e)));
