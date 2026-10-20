# Actionable Message Sender

A dependency-free web app that sends **Outlook Actionable Messages** — Adaptive Cards or legacy
MessageCards authored in the [Actionable Messages Designer](https://amdesigner.azurewebsites.net/) —
through Microsoft Graph **as any mailbox you choose**, using Entra ID **application credentials**
(client ID, tenant ID and client secret; OAuth2 client-credentials flow). There is no interactive sign-in.

The client secret is read from `.env` by the Node server, which calls Graph. It is never sent to the
browser.

## Features

- Send as any mailbox via `POST /users/{from}/sendMail`, or save a draft in that mailbox.
- Paste, upload, or load a sample JSON template; live validation and auto-detection of
  Adaptive Card vs. legacy MessageCard.
- `{{placeholder}}` substitution driven by a simple `name=value` variable list.
- Originator ID injection for cards registered in the Actionable Email Developer Dashboard.
- Builds the Graph-ready HTML body and lets you preview/copy the exact payload.
- Optional `ALLOWED_SENDERS` allow-list to limit which mailboxes the app can send as.

## Prerequisites

- Node.js 18+ (no build step, no dependencies).
- Exchange Online mailboxes for the users you send as.

## 1. Register the Entra ID application

1. **Entra admin center → App registrations → New registration**. No redirect URI is needed.
2. **API permissions → Add a permission → Microsoft Graph → Application permissions**:
   - `Mail.Send`
   - `Mail.ReadWrite` — optional, only for **Save as draft**
3. **Grant admin consent** (required for application permissions).
4. **Certificates & secrets → New client secret**. Copy the secret **Value** immediately.
5. Copy the **Application (client) ID** and **Directory (tenant) ID**.

> **Scope it down.** `Mail.Send` as an application permission lets the app send as *every* mailbox in
> the tenant. Restrict it with
> [RBAC for Applications in Exchange Online](https://learn.microsoft.com/exchange/permissions-exo/application-rbac)
> (or a legacy Application Access Policy), and set `ALLOWED_SENDERS` as a second guard.

## 2. Configure and run

```powershell
cd actionable-messages-spa
Copy-Item .env.example .env   # then edit .env
npm start                     # or: node server.js 3000
```

`.env` values:

| Variable | Required | Purpose |
| --- | --- | --- |
| `TENANT_ID` | yes | Directory (tenant) ID |
| `CLIENT_ID` | yes | Application (client) ID |
| `CLIENT_SECRET` | yes | Client secret value |
| `DEFAULT_SENDER` | no | Pre-fills the **From** field |
| `ALLOWED_SENDERS` | no | Comma-separated mailboxes the app may send as |
| `PORT`, `HOST` | no | Defaults `3000` and `127.0.0.1` |

Open <http://localhost:3000/>. The server listens on loopback only and rejects requests with a foreign
`Host` or `Origin`. Do not expose it to a network without adding authentication, because anyone who can
reach it can send mail as the permitted mailboxes.

## 3. Send a card

1. Design a card at <https://amdesigner.azurewebsites.net/> and copy its JSON (**Card Payload Editor**).
2. Paste it into **1. Card template**, or pick one of the bundled samples in `samples/`.
3. Optionally fill in **Template variables**, one `name=value` per line, to replace `{{name}}` tokens.
4. Enter **From** (the mailbox to send as), recipients, subject and a fallback body.
5. Click **Send to sender** first — a self-addressed card always renders in Outlook, which makes it a
   reliable smoke test.
6. Click **Preview HTML** at any time to inspect the exact `message.body.content` sent to Graph.

## Originator ID and card rendering

Outlook renders actionable cards only when it trusts the sender:

| Scenario | Requirement |
| --- | --- |
| You send a card **to yourself** | Works immediately, no registration needed. |
| Cards with only `Action.OpenUrl` | Still require provider registration to render for other recipients. |
| Cards with `Action.Http` / `ActionCard` / `HttpPOST` | Require a provider registered in the [Actionable Email Developer Dashboard](https://aka.ms/publishoam), and its **originator ID** entered in the app. |

The originator value is injected into the card as the top-level `originator` property before sending.
If an `originator` already exists in your JSON, the value from the UI overrides it.

## How the message is built

Actionable Messages are ordinary HTML emails carrying the card in a `<script>` block inside `<head>`:

```html
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<script type="application/adaptivecard+json">
{ "type": "AdaptiveCard", ... }
</script>
</head>
<body>
<p>Fallback text for clients that cannot render cards.</p>
</body>
</html>
```

Legacy MessageCards use `<script type="application/ld+json">` instead. The card JSON is serialized
with `<`, `>` and `&` escaped as `\uXXXX` so that card content can never break out of the script tag.

That HTML is posted by the server to `POST /users/{from}/sendMail` (or `POST /users/{from}/messages`
for drafts) on `https://graph.microsoft.com/v1.0` with an app-only token, so the mail is sent from the
chosen mailbox.

## Project layout

```
actionable-messages-spa/
├─ index.html                  UI shell
├─ server.js                   static file server + client-credentials mail API (/api/*)
├─ .env.example                tenant / client / secret template (copy to .env)
├─ css/styles.css
├─ js/cardBuilder.js           template validation, variables, HTML generation
├─ js/graph.js                 client for the local /api endpoints
├─ js/app.js                   UI wiring
└─ samples/                    ready-to-use designer templates
```

## Troubleshooting

The activity log decodes common `AADSTS` failures into a one-line fix plus a deep link to your app's
**Authentication** blade.

| Symptom | Fix |
| --- | --- |
| `AADSTS9002326` cross-origin token redemption | The redirect URI is registered under the **Web** platform. See [Fixing AADSTS9002326](#fixing-aadsts9002326) below. |
| `AADSTS7000218` client secret required | Same root cause as above — the app is being treated as a confidential client. Move the URI to the SPA platform. |
| `AADSTS50011` redirect mismatch | The redirect URI in Settings must exactly match the one registered in Entra ID, trailing slash included. |
| `AADSTS65001` consent required | Add delegated `User.Read` + `Mail.Send` and grant consent. |
| `AADSTS700016` app not found | Wrong client ID, or the app lives in a different tenant. |
| Popup closes immediately | Allow popups for `localhost`; the app falls back to redirect sign-in. |
| `ErrorAccessDenied` on send | `Mail.Send` delegated permission is missing or not consented. |
| `ErrorAccessDenied` on **Save as draft** | Drafts need delegated `Mail.ReadWrite`; `Mail.Send` is not enough. |
| Mail arrives but shows only the fallback text | The card needs a registered originator, or the recipient's Outlook does not support cards. |

### Fixing AADSTS9002326

> `Cross-origin token redemption is permitted only for the 'Single-Page Application' client-type.`

MSAL.js redeems the authorization code from the browser using PKCE. Entra ID only allows that when the
redirect URI sits under the **Single-page application** platform. A URI registered under **Web** fails
at the token step — which is why sign-in appears to succeed and *then* errors.

Portal fix:

1. Open **Entra admin center → App registrations → your app → Authentication**.
2. Under the **Web** platform, delete `http://localhost:3000/` (use **Remove platform** if it is the
   only URI there).
3. Click **Add a platform → Single-page application**, enter `http://localhost:3000/`, then **Configure**.
4. **Save**, then reload the app and sign in again.

Manifest fix (same thing, faster) — in **Manifest**, move the URI between these arrays:

```jsonc
"spa": { "redirectUris": ["http://localhost:3000", "http://localhost:3000/"] },
"web": { "redirectUris": [] }
```

`web.redirectUris` must be **empty**. Leaving an entry there is what keeps producing `9002326` even
after you add the SPA platform — the two can coexist, and the Web entry wins.

Changes can take up to a minute to propagate. If sign-in still fails, clear the session: close the tab,
or run `sessionStorage.clear()` in DevTools.

### Verifying the registration without guessing

`AADSTS50011` and `AADSTS9002326` both come from redirect-URI configuration, but they fail at different
stages — which tells you precisely what is wrong:

| Stage that fails | Error | Meaning |
| --- | --- | --- |
| `/authorize` | `AADSTS50011` | That exact URI string is registered **nowhere**. |
| `/token` | `AADSTS9002326` | The URI *is* registered, but under **Web**, not SPA. |
| `/authorize` | `AADSTS9002325` | Registered as SPA (PKCE enforced) — confirms the SPA platform. |

To test a candidate URI without touching the app, open this URL in a browser (substitute your tenant
ID, client ID, and the URL-encoded candidate) and read the outcome:

```text
https://login.microsoftonline.com/<tenant-id>/oauth2/v2.0/authorize?client_id=<client-id>&response_type=code&response_mode=fragment&prompt=none&scope=openid&redirect_uri=<url-encoded-candidate>
```

Landing back on your redirect URI means `/authorize` accepted it; an `AADSTS50011` error page means it
is not registered. Use `response_mode=fragment`, because that is what MSAL.js sends.

## Security notes

- The app is a public client: no secret is stored or required.
- Tokens live in `sessionStorage` and disappear when the tab closes.
- Scopes are requested incrementally: sign-in asks only for `User.Read` + `Mail.Send`; `Mail.ReadWrite`
  is requested on demand the first time you save a draft.
- Client ID / tenant ID / draft content stay in this browser's local storage only.
