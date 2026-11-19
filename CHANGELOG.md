# Changelog

## 1.1.0 — 2026-11-19
- **One-command setup:** `npm run setup` asks for your email and time zone, makes storage, deploys, and walks you through the Microsoft app.
- Your settings now live in a git-ignored `wrangler.jsonc` (made from `wrangler.example.jsonc`), so your email isn't in the repo.
- **Safer by default:** sign-in is refused if `ALLOWED_EMAILS` is empty.
- Sign-in pages can't be framed by other sites (clickjacking protection).
- Retries once when Microsoft says it's busy (429/503), with clearer error messages.
- Real tests (`npm test`) and CI on every pull request.
- MIT license, contributing guide, security policy, issue templates.

## 1.0.0
- 16 tools for Outlook email and calendar, including repeating events.
