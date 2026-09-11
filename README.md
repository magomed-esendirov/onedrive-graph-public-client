# OneDrive via Microsoft Graph: connecting a public client

> Russian version: [README.ru.md](README.ru.md)

How to connect your OneDrive to your own app or script via the Microsoft Graph API — **without a `client_secret`** — and the mistakes we made along the way before it worked.

## TL;DR — the working configuration

- Entra ID → App registrations → New registration
- **Supported account types:** `Accounts in any organizational directory (Any Microsoft Entra ID directory - Multitenant) and personal Microsoft accounts` — if you sign in with a personal account (`@outlook.com`, `@hotmail.com`, `@live.com`)
- **Platform:** `Mobile and desktop applications` (Native client, public client — no secret, none needed)
- **Redirect URI:** `https://login.microsoftonline.com/common/oauth2/nativeclient` (not actually used by the device code flow, but the platform must be added)
- **Authentication → Allow public client flows: Yes**
- **API permissions (Delegated):** `User.Read`, `Files.ReadWrite`
- **Authority/endpoint:** `https://login.microsoftonline.com/common` (multitenant + personal) or `/consumers` (personal only) — **not** a tenant-specific URL
- **Scopes:** `offline_access User.Read Files.ReadWrite`

## Step-by-step app registration

1. Go to https://entra.microsoft.com → **Identity** → **Applications** → **App registrations** → **New registration**.
2. **Name** — anything, e.g. `My OneDrive App`.
3. **Supported account types** — select `Accounts in any organizational directory (Any Microsoft Entra ID directory - Multitenant) and personal Microsoft accounts (e.g. Skype, Xbox)`.
4. You can skip the Redirect URI for now — we'll add the platform next.
5. **Register** → on the Overview page copy the **Application (client) ID** — that's your `client_id`.
6. On the left, **Authentication** → **Add a platform** → **Mobile and desktop applications** → check `https://login.microsoftonline.com/common/oauth2/nativeclient` → **Configure**.
7. On the same page, at the bottom: **Allow public client flows → Yes** → **Save**.
8. On the left, **API permissions** → **Add a permission** → **Microsoft Graph** → **Delegated permissions** → check `User.Read` and `Files.ReadWrite` → **Add permissions**.
9. For delegated permissions with a personal account, no admin consent is needed — you can start using it right away.

## Why not Web or SPA: the errors we hit

We went through three platform types before OAuth worked. Each one fails with a characteristic error, which makes diagnosis easy.

### Error 1: Web platform → `AADSTS70002`

If you add the **Web** platform under **Authentication**, the app is treated as a confidential client. When redeeming the authorization code for a token, Microsoft demands a secret:

```
AADSTS70002: The request body must contain the following parameter:
'client_assertion' or 'client_secret'.
```

For a personal script or agent you don't want to store a secret — and you don't need to. Fix: remove the Web platform and make it a public client (no secret required at all).

### Error 2: SPA platform → `AADSTS90023`

We removed Web and set **Single-page application** — the secret requirement went away, but the token endpoint started responding with:

```
AADSTS90023: Cross-origin token redemption is permitted only for the
'Single-Page Application' client-type.
```

Translation: the SPA type only allows code-for-token redemption **from a browser** (the request must come with an `Origin` header, cross-origin). Code running outside a browser (a server, script, or agent) gets rejected. No setting fixes this — only changing the platform type.

### The fix: Native (Mobile and desktop applications)

The **Mobile and desktop applications** platform is a public client with no secret that is allowed to redeem codes outside the browser (authorization code + PKCE, device code flow). After switching from SPA to Native, OAuth worked on the first try — nothing else needed changing.

### Error 3: `AADSTS50020` — wrong tenant / wrong account types

```
AADSTS50020: User account '...' from identity provider 'live.com' does not
exist in tenant '...'
```

Two possible causes:

1. The registration is set to **single tenant** while you sign in with a personal account. Fix — select **Multitenant + personal Microsoft accounts** (step 3 above).
2. Your code uses a tenant-specific endpoint `https://login.microsoftonline.com/<tenant-id>/...`. For this scenario use `https://login.microsoftonline.com/common` (multitenant + personal) or `https://login.microsoftonline.com/consumers` (personal only).

## Getting a token: device code flow

The simplest approach for a script — no web server or redirect URI needed:

```python
import msal

app = msal.PublicClientApplication(
    client_id="<your application (client) id>",
    authority="https://login.microsoftonline.com/common",
)
flow = app.initiate_device_flow(scopes=["User.Read", "Files.ReadWrite", "offline_access"])
print(flow["message"])  # open the URL from the message, enter the code
result = app.acquire_token_by_device_flow(flow)
access_token = result["access_token"]
```

Tokens live ~1 hour; `offline_access` in the scopes gives you a refresh token, and `acquire_token_silent` renews the session without user interaction.

## Key Graph calls

Base: `https://graph.microsoft.com/v1.0`, header `Authorization: Bearer <access_token>`.

| What | Request |
|---|---|
| Verify the token | `GET /me` |
| OneDrive root | `GET /me/drive/root/children` |
| Search | `GET /me/drive/root/search(q='contract')` |
| Download a file | `GET /me/drive/root:/Documents/file.txt:/content` |
| Upload a file (≤ 4 MB) | `PUT /me/drive/root:/Documents/file.txt:/content` |
| Delete | `DELETE /me/drive/root:/Documents/file.txt` |

Paths like `/me/drive/root:/Documents/Report.pdf` need URL-encoding (except `/`).

### The download gotcha: 302 to another host

`GET ...:/content` does **not** return the file directly — it responds with **302** to a host like `*.my.microsoftpersonalcontent.com`. That's a pre-authenticated URL: the token is already embedded in it.

**Do not send your `Authorization` header there.** It's not needed, and its presence breaks the download (in our case the request just hung until timeout). The correct approach:

1. Request `...:/content` with automatic redirects disabled.
2. Take the `Location` from the response.
3. Download the `Location` **without** the `Authorization` header.

With `requests` this almost works out of the box (it strips `Authorization` on cross-host redirects), but handling the 302 manually is more reliable — see the `download()` function in `example/graph_client.py`.

### Search caveat

`search(q='...')` queries OneDrive's search index, not the live filesystem. A freshly uploaded file shows up in results **after a few minutes** — that's normal, not a bug in your code.

## Example

`example/graph_client.py` — a minimal working client using `msal` + `requests`:

- sign-in via device code flow (with token cache),
- `me`, listing files, search,
- upload, download (with manual 302 handling), delete.

```bash
pip install -r requirements.txt
export ONEDRIVE_CLIENT_ID="<your application (client) id>"
python example/graph_client.py
```

## Further reading

- `AADSTS90023` explained (cross-origin token redemption): https://learn.microsoft.com/en-us/answers/questions/588174/microsoft-devicecode-authentication-returning-inva
- `AADSTS50020` explained (account not found in tenant): https://learn.microsoft.com/en-us/answers/questions/2203176/issue-with-microsoft-oauth-for-outlook-aadsts50020
