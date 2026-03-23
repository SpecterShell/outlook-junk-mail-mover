# Outlook Junk Mail Mover

Languages: English | [简体中文](README.zh-CN.md) | [繁體中文](README.zh-TW.md)

This script polls the Outlook `junkemail` folder through Microsoft Graph and moves matching messages back into the `inbox`.

- Authenticates with Microsoft Graph using device-code login.
- Looks at recent messages in `junkemail`.
- Matches them by sender address, sender domain, subject keywords, or body keywords.
- Moves matches to `inbox`.
- Repeats on a timer, or runs once for use with cron/systemd.

## Setup

1. Register an app in Microsoft Entra / Azure and note its client ID.
2. In the app registration:
   - Open the app Overview page and copy the Application (client) ID.
   - If you want to sign in with Outlook.com/Hotmail/Live, make sure Supported account types includes personal Microsoft accounts.
   - In Authentication, enable public client flows.
   - In API permissions, add Microsoft Graph delegated permission `Mail.ReadWrite`.
3. Copy `.env.example` to `.env` and fill in your settings. 
   Instead of editing `.env` manually, you can run the configuration wizard:

   ```bash
   uv run outlook_junk_mover.py --configure
   ```

## Example config

```dotenv
OUTLOOK_CLIENT_ID=your-app-client-id
OUTLOOK_TENANT_ID=consumers
OUTLOOK_ALLOWED_SENDERS=john.doe@example.com,noreply@example.com
OUTLOOK_SUBJECT_KEYWORDS=verification code,login code
OUTLOOK_POLL_SECONDS=120
```

For personal Outlook.com/Hotmail/Live mailboxes, set `OUTLOOK_TENANT_ID=consumers` and make sure the app registration supports personal Microsoft accounts.

For work or school mailboxes, `OUTLOOK_TENANT_ID` should usually be your actual tenant GUID or verified domain, not `common`. `common` is convenient in browser-based flows, but device-code flow often needs a tenant-specific authority for workforce tenants.

## Running

Run once:

```bash
python3 outlook_junk_mover.py --once
```

Run continuously:

```bash
python3 outlook_junk_mover.py
```

On the first run, Microsoft will show a device-code login prompt. After you finish sign-in, the script stores a token cache at `.tokens/msal_cache.json`.

## Notes

- Exact sender allowlists are safer than domain allowlists.
- `OUTLOOK_MOVE_ALL=true` will move all recent mail from Junk to Inbox. That is usually a bad idea unless you know what you are doing.
- If your tenant blocks device-code flow with Conditional Access, this script will not be able to authenticate until that policy is changed or you switch to a different auth flow.

## References

- Microsoft Graph message move API: https://learn.microsoft.com/en-us/graph/api/message-move?view=graph-rest-1.0
- Microsoft Graph permissions reference (`Mail.ReadWrite`): https://learn.microsoft.com/en-us/graph/permissions-reference
- MSAL Python token acquisition: https://learn.microsoft.com/en-us/entra/msal/python/getting-started/acquiring-tokens
- Public client / device-code flow guidance: https://learn.microsoft.com/en-us/entra/identity-platform/msal-client-applications
