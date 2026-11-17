import plaud_transcribe_notion as m

acc = m.find_account("account2")

print("=== before refresh ===")
print("token_prefix:", acc["token"][:20], "...")
print("_bearer_toggled:", acc.get("_bearer_toggled"))

ok = m.refresh_workspace_token(acc)
print("\n=== refresh_workspace_token() result ===")
print("returned:", ok)
print("token_prefix after:", acc["token"][:20], "...")
print("_bearer_toggled after:", acc.get("_bearer_toggled"))

print("\n=== now calling get_plaud_files_for_account() ===")
files = m.get_plaud_files_for_account(acc)
print("files found:", len(files))
