import requests
import plaud_transcribe_notion as m

acc = m.find_account("account2")
print("=== account2 config ===")
print("domain:", acc["domain"])
print("workspace:", acc["workspace"])
print("token_prefix:", (acc["token"][:15] + "...") if acc["token"] else None)
print("refresh_token_len:", len(acc.get("refresh_token") or ""))
print("refresh_token_has_bearer_prefix:", (acc.get("refresh_token") or "").lower().startswith("bearer "))

print("\n=== calling refresh_workspace_token() directly ===")
refresh_token = acc.get("refresh_token")
url = f"{acc['domain']}/user-app/auth/workspace/refresh/{acc['workspace']}"
bearer_prefix = "Bearer" if not refresh_token.lower().startswith("bearer ") else ""
headers = {"Authorization": f"{bearer_prefix} {refresh_token}".strip()}
print("POST", url)
try:
    resp = requests.post(url, json={}, headers=headers, timeout=30)
    print("status_code:", resp.status_code)
    print("body:", resp.text[:1000])
except requests.RequestException as e:
    print("request failed:", e)
