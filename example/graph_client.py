#!/usr/bin/env python3
"""Minimal OneDrive client via Microsoft Graph.

Public client (Native / Mobile and desktop applications) — no
client_secret. Sign-in via device code flow.

Requires: pip install -r requirements.txt
Setup:    export ONEDRIVE_CLIENT_ID="<your application (client) id>"
"""
import os
import sys
import urllib.parse

import msal
import requests

CLIENT_ID = os.environ.get("ONEDRIVE_CLIENT_ID", "")
# multitenant + personal accounts. Personal only -> ".../consumers".
AUTHORITY = "https://login.microsoftonline.com/common"
SCOPES = ["User.Read", "Files.ReadWrite", "offline_access"]
GRAPH = "https://graph.microsoft.com/v1.0"


def get_token():
    """Device code flow with silent renewal from the cache."""
    if not CLIENT_ID or CLIENT_ID.startswith("<"):
        raise SystemExit("Set ONEDRIVE_CLIENT_ID (the Application (client) ID from Entra).")
    app = msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY)
    for account in app.get_accounts():
        result = app.acquire_token_silent(SCOPES, account=account)
        if result and "access_token" in result:
            return result["access_token"]
    flow = app.initiate_device_flow(scopes=SCOPES)
    if "user_code" not in flow:
        raise SystemExit(f"Could not start device flow: {flow}")
    print(flow["message"])  # open the URL, enter the code
    result = app.acquire_token_by_device_flow(flow)
    if "access_token" not in result:
        raise SystemExit(f"Token error: {result.get('error_description', result)}")
    return result["access_token"]


def auth_headers(token):
    return {"Authorization": f"Bearer {token}"}


def item_url(path):
    """/me/drive/root:/Documents/file.txt with proper encoding."""
    p = "/" + path.strip().lstrip("/")
    return f"{GRAPH}/me/drive/root:{urllib.parse.quote(p, safe='/')}"


def me(token):
    r = requests.get(f"{GRAPH}/me", headers=auth_headers(token),
                     params={"$select": "displayName,mail"}, timeout=30)
    r.raise_for_status()
    info = r.json()
    print("Signed in as:", info.get("displayName"), info.get("mail"))


def list_dir(token, path="/", top=20):
    base = f"{GRAPH}/me/drive/root" if path in ("", "/") else item_url(path)
    r = requests.get(base + "/children", headers=auth_headers(token),
                     params={"$select": "name,size,folder,file", "$top": top,
                             "$orderby": "name"}, timeout=30)
    r.raise_for_status()
    for it in r.json().get("value", []):
        kind = "DIR " if "folder" in it else "FILE"
        print(f"{kind} {it['name']}")


def search(token, query, top=20):
    q = urllib.parse.quote(query, safe="")
    r = requests.get(f"{GRAPH}/me/drive/root/search(q='{q}')",
                     headers=auth_headers(token),
                     params={"$select": "name,parentReference,size", "$top": top},
                     timeout=30)
    r.raise_for_status()
    items = r.json().get("value", [])
    for it in items:
        parent = (it.get("parentReference") or {}).get("path", "")
        print(f"{it['name']}  [{parent}]")
    if not items:
        print("(nothing found — new files take a few minutes to get indexed)")


def upload(token, local_path, remote_path):
    """Simple upload, files up to 4 MB."""
    with open(local_path, "rb") as f:
        data = f.read()
    if len(data) > 4 * 1024 * 1024:
        raise SystemExit("Simple upload supports files up to 4 MB only.")
    r = requests.put(item_url(remote_path) + ":/content",
                     headers={**auth_headers(token),
                              "Content-Type": "application/octet-stream"},
                     data=data, timeout=120)
    r.raise_for_status()
    print(f"uploaded -> {r.json().get('name')} ({r.json().get('size')} bytes)")


def download(token, remote_path, local_path):
    """Download with manual 302 handling.

    Graph answers 302 to a pre-authenticated URL (the token is already
    embedded in it). Do NOT send the Authorization header there —
    download the Location without any headers.
    """
    r = requests.get(item_url(remote_path) + ":/content",
                     headers=auth_headers(token),
                     allow_redirects=False, timeout=60)
    if r.status_code in (301, 302, 303, 307, 308):
        location = r.headers.get("Location")
        if not location:
            raise SystemExit("302 response has no Location header.")
        dl = requests.get(location, timeout=180)  # no Authorization!
        dl.raise_for_status()
        data = dl.content
    else:
        r.raise_for_status()
        data = r.content
    with open(local_path, "wb") as f:
        f.write(data)
    print(f"downloaded {len(data)} bytes -> {local_path}")


def delete(token, remote_path):
    r = requests.delete(item_url(remote_path), headers=auth_headers(token), timeout=30)
    r.raise_for_status()
    print(f"deleted {remote_path}")


def main():
    token = get_token()
    me(token)

    print("\n-- OneDrive root --")
    list_dir(token)

    # Round trip: upload -> download -> compare -> delete
    probe_remote = "/Documents/__graph_example_probe.txt"
    with open("/tmp/__probe.txt", "w") as f:
        f.write("probe: Microsoft Graph public client works\n")
    print("\n-- round trip --")
    upload(token, "/tmp/__probe.txt", probe_remote)
    download(token, probe_remote, "/tmp/__probe_back.txt")
    same = open("/tmp/__probe.txt", "rb").read() == open("/tmp/__probe_back.txt", "rb").read()
    print("bytes match:", same)
    delete(token, probe_remote)


if __name__ == "__main__":
    main()
