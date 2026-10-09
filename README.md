# Personal Microsoft MCP (Outlook, OneDrive, To Do)

A remote MCP server on Cloudflare Workers that signs in with a **personal** Microsoft account
and exposes Microsoft Graph tools to Claude.

**Tools:** `outlook_email_search`, `outlook_email_read`, `onedrive_search`, `onedrive_list_folder`,
`onedrive_read_file`, `todo_list_lists`, `todo_list_tasks`, `todo_create_task`, `todo_complete_task`

Email and OneDrive are read-only. To Do can create and complete tasks.

> Status: written but not yet deployed or type-checked. If `npm run typecheck` complains about
> the `agents` or `@cloudflare/workers-oauth-provider` APIs, compare against Cloudflare's
> `remote-mcp-github-oauth` template, since those libraries change quickly.

## 1. Register the app in Microsoft Entra

1. Go to https://entra.microsoft.com (sign in with the Microsoft account; a free tenant is created if needed)
   and open **App registrations > New registration**.
2. **Supported account types:** "Personal Microsoft accounts only" (or "any organizational directory and personal Microsoft accounts").
3. **Redirect URI:** platform **Web**, `https://personal-m365-mcp.<your-subdomain>.workers.dev/callback`
   (add `http://localhost:8788/callback` too if you want to use `wrangler dev`).
4. **API permissions > Add > Microsoft Graph > Delegated:** `User.Read`, `Mail.Read`, `Files.Read`,
   `Tasks.ReadWrite`, `offline_access`.
5. **Certificates & secrets > New client secret.** Copy the secret **Value** right away.
6. Copy the **Application (client) ID** from the Overview page.

## 2. Deploy

```bash
npm install
npx wrangler kv namespace create OAUTH_KV     # paste the printed id into wrangler.jsonc
npx wrangler secret put MICROSOFT_CLIENT_ID
npx wrangler secret put MICROSOFT_CLIENT_SECRET
npx wrangler secret put ALLOWED_EMAIL          # optional: your Microsoft email, locks the server to you
npx wrangler deploy
```

If `npm install` reports a peer-dependency conflict, add `--legacy-peer-deps`.

## 3. Connect it to Claude

In Claude's connector settings, add a **custom connector** with the URL
`https://personal-m365-mcp.<your-subdomain>.workers.dev/mcp`, then sign in with your Microsoft account.

## Notes

- Uses the `consumers` Microsoft endpoint, so work/school accounts won't sign in. Change `MS_AUTH`
  in `src/index.ts` to `/common/` if you want both.
- Teams chat and SharePoint are not available for personal accounts, so they aren't included.
- `onedrive_read_file` only reads text-based files (txt, md, csv, json, code). Word/PDF parsing isn't included.
- Client secrets expire (max 24 months). When it does, create a new one and re-run `wrangler secret put MICROSOFT_CLIENT_SECRET`.
