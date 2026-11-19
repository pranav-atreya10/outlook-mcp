# Contributing

Thanks for wanting to help! Bug reports, ideas and code changes are all welcome.

## Reporting a bug or asking for a feature
For security problems, see [SECURITY.md](SECURITY.md) instead.

Open an [issue](https://github.com/pranav-atreya10/outlook-mcp/issues). For bugs, say what you asked Claude to do, what happened, and what you expected. Paste any errors from `npx wrangler tail`, but **remove email addresses, message contents and tokens first**.

## Making a change
1. Fork the repo and make a branch: `git checkout -b my-change`
2. Install: `npm install`
3. Make your change. Most tools live in `src/tools.ts`.
4. Check it still works:
   ```sh
   npm run typecheck
   npm test
   ```
5. If you can, try it for real on your own copy (`npm run setup`).
6. Open a pull request that explains what you changed and why.

Please keep each pull request to one change so it's easy to review.

## Adding a new tool
Copy the pattern of an existing tool in `src/tools.ts`:
- Give it a short `snake_case` name and a clear description, since Claude reads the description to decide when to use it.
- Describe each input with `zod`, including `.describe(...)` text for anything that isn't obvious.
- Wrap the handler in `safe(...)` so errors come back as friendly messages.
- If it needs a new Microsoft Graph permission, add it to the scopes in `src/microsoft.ts` and to the permission list in the README.
- Update the tool count and list in the README.

## Ideas for a first contribution
- Add a tool to search for a contact's email address.
- Add a tool to move an email to a different folder.
- Add a tool to download or list email attachments.
- Add a tool to find free time on the calendar (Microsoft Graph has a `findMeetingTimes` endpoint).
- Support work/school accounts (see the comment at the top of `src/microsoft.ts`).

## Be kind
Be respectful and helpful to everyone here. Disagree about code, not people.
