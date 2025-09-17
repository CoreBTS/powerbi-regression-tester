import os
import msal
import requests
import jwt
from sys import path

import clr
clr.AddReference("System")
import System
from System.Net import ServicePointManager, SecurityProtocolType

# Force TLS 1.2
ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12


# Your Azure AD App Registration details
# CLIENT_ID = '54640219-af44-42bb-adcd-d3722aa55e04'  # Application (client) ID
# # CLIENT_ID = '7f67af8a-fedc-4b08-8b4e-37c4d127b6cf'
# TENANT_ID = 'e39cce29-5716-43ba-b27d-1bdd8fd67901'  # Or your tenant ID
# AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
# AUTHORITY = f"https://login.microsoftonline.com/common"
# # SCOPES = ["https://analysis.windows.net/powerbi/api/.default", "https://analysis.windows.net/powerbi/api/user_impersonation"]
# SCOPES = ["https://analysis.windows.net/powerbi/api/.default"]
# # SCOPES = ["https://analysis.windows.net/powerbi/api/user_impersonation"]

# CACHE_FILE = "token_cache.bin"
# force_reauthentication = True  # Set to True to force re-authentication

# # Create persistent cache
# cache = msal.SerializableTokenCache()
# if os.path.exists(CACHE_FILE):
#     cache.deserialize(open(CACHE_FILE, "r").read())

# app = msal.PublicClientApplication(
#     client_id=CLIENT_ID,
#     # authority="https://login.microsoftonline.com/common",
#     authority=AUTHORITY,
#     token_cache=cache
# )

# # Create the public client app
# app = msal.PublicClientApplication(
#     client_id=CLIENT_ID,
#     authority=AUTHORITY
# )

# Launch browser-based interactive login
# result = app.acquire_token_interactive(scopes=SCOPES, prompt='select_account')


# # Attempt to acquire token silently (cached)
# accounts = app.get_accounts()
# if accounts and not force_reauthentication:
#     result = app.acquire_token_silent(SCOPES, account=accounts[0])
# else:
#     result = None

# # If silent fails, do interactive login with prompt
# if not result:
#     result = app.acquire_token_interactive(
#         scopes=SCOPES,
#         prompt="select_account"  # Always allow user to switch
#     )

# # Save cache back to file
# with open(CACHE_FILE, "w") as f:
#     f.write(cache.serialize())

# # claims = result.get("id_token_claims", {})

# # print("Logged in as:", claims.get("name") or claims.get("preferred_username") or claims.get("upn") or "Unknown user")

# access_token = result["access_token"]
# decoded = jwt.decode(access_token, options={"verify_signature": False})

# print("User in access token:", decoded.get("preferred_username") or decoded.get("upn") or decoded.get("name"))


# Show user info
# print("Logged in as:", result.get("name") or result.get("preferred_username"))




# if "access_token" in result:
#     print("Access Token:")
#     print(result["access_token"])
# else:
#     print("Error acquiring token:")
#     print(result.get("error"))
#     print(result.get("error_description"))


# conn_str = f"""
# Provider=MSOLAP;
# Data Source=pbiazure://api.powerbi.com/;
# Initial Catalog=sobe_wowvirtualserver-37b188f0-d623-4d60-b032-8a1ef55be1fb;
# Integrated Security=ClaimsToken;
# Persist Security Info=True;
# Identity Provider="https://login.microsoftonline.com/common, https://analysis.windows.net/powerbi/api, {CLIENT_ID}";
# Password={result['access_token']}
# """.strip()
# print(conn_str)

# conn_str = (
#     f'Provider=MSOLAP;Data Source=pbiazure://api.powerbi.com/;Initial Catalog=sobe_wowvirtualserver-37b188f0-d623-4d60-aaaa-8a1ef55bcccc;'
#     f'Integrated Security=ClaimsToken;Persist Security Info=True;Identity Provider="https://login.microsoftonline.com/common, https://analysis.windows.net/powerbi/api, your-client-id-here";'
#     f'Password={access_token};'
# )

# a = fr'Test \\1 \2 Done'
# a = fr'Test \1 Done'
# print(a)

# conn_str = (
# f"Provider=MSOLAP;"
# f"Data Source=pbiazure://api.powerbi.com/;"
# f"Initial Catalog=sobe_wowvirtualserver-37b188f0-d623-4d60-b032-8a1ef55be1fb;"
# f"Integrated Security=ClaimsToken;"
# f"Persist Security Info=True;"
# fr'Identity Provider="https://login.microsoftonline.com/common, https://analysis.windows.net/powerbi/api, {CLIENT_ID}";'
# f'Password={result["access_token"]}'
# )

# print(conn_str)

# # # ==== CONFIGURATION ====
# # CLIENT_ID = 'your-client-id-here'  # From Azure app registration
# # TENANT_ID = 'common'               # Or use your specific tenant ID
# # AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
# SCOPE = ["https://analysis.windows.net/powerbi/api/.default"]
# API_BASE = "https://api.powerbi.com/v1.0/myorg"

# # # ==== INTERACTIVE LOGIN ====
# # app = msal.PublicClientApplication(client_id=CLIENT_ID, authority=AUTHORITY)

# # result = app.acquire_token_interactive(scopes=SCOPE)

# # if "access_token" not in result:
# #     print("Failed to acquire token.")
# #     print(result.get("error_description"))
# #     exit(1)

# access_token = result['access_token']
# headers = {
#     'Authorization': f'Bearer {access_token}'
# }

# ==== LIST WORKSPACES ====
# print("🔍 Getting workspaces...")
# res = requests.get(f"{API_BASE}/groups", headers=headers)
# res.raise_for_status()

# groups = res.json().get("value", [])
# print(f"\n📂 Found {len(groups)} workspaces:\n")

# for group in groups:
#     group_id = group["id"]
#     group_name = group["name"]
#     print(f"➡️  Workspace: {group_name} ({group_id})")

#     # ==== LIST DATASETS IN EACH WORKSPACE ====
#     ds_res = requests.get(f"{API_BASE}/groups/{group_id}/datasets", headers=headers)
#     ds_res.raise_for_status()
#     datasets = ds_res.json().get("value", [])

#     if datasets:
#         for ds in datasets:
#             print(f"    📊 Dataset: {ds['name']} ({ds['id']})")
#     else:
#         print("    ⚠️  No datasets in this workspace.")
#     print()


dax = "EVALUATE Apps"

# from sys import path
# adomd_path = r'C:\Program Files\Microsoft.NET\ADOMD.NET\160'

# if not os.path.isdir(adomd_path):
#     print("Folder does not exist.")

# # Check if the ADOMD.NET path is already in the system path
# if adomd_path not in path:
#     path.append(adomd_path)

from pyadomd import Pyadomd
# print(conn_str)

access_token = "eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiIsIng1dCI6Il9qTndqZVNudlRUSzhYRWRyNVFVUGtCUkxMbyIsImtpZCI6Il9qTndqZVNudlRUSzhYRWRyNVFVUGtCUkxMbyJ9.eyJhdWQiOiJodHRwczovL2FuYWx5c2lzLndpbmRvd3MubmV0L3Bvd2VyYmkvYXBpIiwiaXNzIjoiaHR0cHM6Ly9zdHMud2luZG93cy5uZXQvZTM5Y2NlMjktNTcxNi00M2JhLWIyN2QtMWJkZDhmZDY3OTAxLyIsImlhdCI6MTc1MTQ3MDA2MiwibmJmIjoxNzUxNDcwMDYyLCJleHAiOjE3NTE0NzUxMzcsImFjY3QiOjAsImFjciI6IjEiLCJhaW8iOiJBVVFBdS84WkFBQUF2V3FJWU5BR2dZdXFrbGRrOU95OVpxZUVFdlB4R2Y2R2ZFNlZzSUpRNUdJVkdyUHJTNTNVQjRwU2JaVVRYeU1Ja3pJamhCMjR3eThIL2ZaRGRtemtZUT09IiwiYW1yIjpbInB3ZCJdLCJhcHBpZCI6IjdmNjdhZjhhLWZlZGMtNGIwOC04YjRlLTM3YzRkMTI3YjZjZiIsImFwcGlkYWNyIjoiMCIsImlkdHlwIjoidXNlciIsImlwYWRkciI6IjEzNS4xMzEuMTMzLjE0IiwibmFtZSI6ImZ0dTI2Iiwib2lkIjoiYzhiZTY1YmMtM2RkOS00NDNmLWFiNDYtYjZlMWQ1OGYxYTIzIiwicHVpZCI6IjEwMDMyMDA0Q0VCOUVBM0QiLCJyaCI6IjEuQVJnQUtjNmM0eFpYdWtPeWZSdmRqOVo1QVFrQUFBQUFBQUFBd0FBQUFBQUFBQUJxQVFnWUFBLiIsInNjcCI6InVzZXJfaW1wZXJzb25hdGlvbiIsInNpZCI6IjAwNmFmZWI5LWQzZjYtNjNjNC01M2YwLWY3ODRmOTcyM2JhZiIsInN1YiI6Ikc1eVJnaS1BTm5FR1dBZmtUdlJvM1FaUTJFa203SDRTOTN3TmtzdXZYNEEiLCJ0aWQiOiJlMzljY2UyOS01NzE2LTQzYmEtYjI3ZC0xYmRkOGZkNjc5MDEiLCJ1bmlxdWVfbmFtZSI6ImZ0dTI2QFNreWxpbmVEYXRhQW5hbHl0aWNzLm9ubWljcm9zb2Z0LmNvbSIsInVwbiI6ImZ0dTI2QFNreWxpbmVEYXRhQW5hbHl0aWNzLm9ubWljcm9zb2Z0LmNvbSIsInV0aSI6IjZweGl3Y2l6OWtleUlEUnVISGdRQUEiLCJ2ZXIiOiIxLjAiLCJ3aWRzIjpbImI3OWZiZjRkLTNlZjktNDY4OS04MTQzLTc2YjE5NGU4NTUwOSJdLCJ4bXNfZnRkIjoiVW16d1ZPRHBFVXd6Z1dxd0N1R08wZFRzMm5TZy03S05DZ0dKX3dPU3Vjc0JkWE56YjNWMGFDMWtjMjF6IiwieG1zX2lkcmVsIjoiMSAxNiJ9.flXtekKmGEiZ9C3rnQTgVL4IgYETCr0aCHOhMO8CkebHkTWSqQD_8FsUshkUG1yDgNkoSaru7aaYMCHQGeFVr24Y2To0ndi3ocYnEbCQBAxXIRmbzP5DJ_jYVXyQobxYWBTDP35DR0vW0_030iWjuRtDURau6QZmsQEovKJpousMsUikoFBZsm7ZZB-ltGzYKzWCh56LeNoNwrrYoZiV_WORYYiEhPUvrQmQwZExhRH8x3fS2PuxyeO4prKBIYkUj3uHg66PPT3o9zWoT39IxTaVM8HzIUDduxURSB-QPPWCFks1LDiC4LOd2wqikWZ0UXnKY2Ymies66CrqD2NXSg"
conn_str = f"Provider=MSOLAP;Data Source=pbiazure://api.powerbi.com/;Initial Catalog=sobe_wowvirtualserver-37b188f0-d623-4d60-b032-8a1ef55be1fb;Integrated Security=ClaimsToken;Persist Security Info=True;Identity Provider=\"https://login.microsoftonline.com/common, https://analysis.windows.net/powerbi/api, 54640219-af44-42bb-adcd-d3722aa55e04\";Password={access_token}"
# conn_str = "Provider=MSOLAP;Data Source=pbiazure://api.powerbi.com/;Initial Catalog=sobe_wowvirtualserver-37b188f0-d623-4d60-b032-8a1ef55be1fb;Integrated Security=ClaimsToken;Persist Security Info=True;Identity Provider=\"https://login.microsoftonline.com/common, https://analysis.windows.net/powerbi/api, 54640219-af44-42bb-adcd-d3722aa55e04\";Password=eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiIsIng1dCI6Il9qTndqZVNudlRUSzhYRWRyNVFVUGtCUkxMbyIsImtpZCI6Il9qTndqZVNudlRUSzhYRWRyNVFVUGtCUkxMbyJ9.eyJhdWQiOiJodHRwczovL2FuYWx5c2lzLndpbmRvd3MubmV0L3Bvd2VyYmkvYXBpIiwiaXNzIjoiaHR0cHM6Ly9zdHMud2luZG93cy5uZXQvZTM5Y2NlMjktNTcxNi00M2JhLWIyN2QtMWJkZDhmZDY3OTAxLyIsImlhdCI6MTc1MTM5MzcwOCwibmJmIjoxNzUxMzkzNzA4LCJleHAiOjE3NTEzOTkwNDksImFjY3QiOjAsImFjciI6IjEiLCJhaW8iOiJBVVFBdS84WkFBQUFza2lpSERRNXdzQzNHdzZUY0doZVVPdU5WY2Y1UmlpazZibnYveURJZEg2aU1MamhQZVlpNnhlYVZOUXNFZStwRzJpajc4YjVrbTZzT0ZsMkEvMVZtZz09IiwiYW1yIjpbInB3ZCJdLCJhcHBpZCI6IjdmNjdhZjhhLWZlZGMtNGIwOC04YjRlLTM3YzRkMTI3YjZjZiIsImFwcGlkYWNyIjoiMCIsImlkdHlwIjoidXNlciIsImlwYWRkciI6IjEzNS4xMzEuMTMzLjE0IiwibmFtZSI6ImZ0dTI2Iiwib2lkIjoiYzhiZTY1YmMtM2RkOS00NDNmLWFiNDYtYjZlMWQ1OGYxYTIzIiwicHVpZCI6IjEwMDMyMDA0Q0VCOUVBM0QiLCJyaCI6IjEuQVJnQUtjNmM0eFpYdWtPeWZSdmRqOVo1QVFrQUFBQUFBQUFBd0FBQUFBQUFBQUJxQVFnWUFBLiIsInNjcCI6InVzZXJfaW1wZXJzb25hdGlvbiIsInNpZCI6IjAwNmFmZWI5LWQzZjYtNjNjNC01M2YwLWY3ODRmOTcyM2JhZiIsInN1YiI6Ikc1eVJnaS1BTm5FR1dBZmtUdlJvM1FaUTJFa203SDRTOTN3TmtzdXZYNEEiLCJ0aWQiOiJlMzljY2UyOS01NzE2LTQzYmEtYjI3ZC0xYmRkOGZkNjc5MDEiLCJ1bmlxdWVfbmFtZSI6ImZ0dTI2QFNreWxpbmVEYXRhQW5hbHl0aWNzLm9ubWljcm9zb2Z0LmNvbSIsInVwbiI6ImZ0dTI2QFNreWxpbmVEYXRhQW5hbHl0aWNzLm9ubWljcm9zb2Z0LmNvbSIsInV0aSI6InVlUHVJdVJCNjBPdXd2ajVPOVExQUEiLCJ2ZXIiOiIxLjAiLCJ3aWRzIjpbImI3OWZiZjRkLTNlZjktNDY4OS04MTQzLTc2YjE5NGU4NTUwOSJdLCJ4bXNfZnRkIjoiY2FLRXJSUU82bUhyUjFpWG1mUUxLeUJVN3R5N01GNnZqSlJoUlJhMkxTMEJkWE5sWVhOMExXUnpiWE0iLCJ4bXNfaWRyZWwiOiIxIDI4In0.VNZwX8JnYIqgqEnPK_X_qoCQj8Hea2UGaPwgYfibZvtAoDyMI-CCFe35FsjxXyXOz76vxuN8HtcAoT2gFo2geinOc3JvzbRLrle-v1OU9AQ09OFeW-AuTq35lhtunGpm-XbfeKolPwqLr6Hn7XF5A6XilYvnYWdICyovc0vC8Q-Ia1aNyQG53eyLXq_cK8HoJGW-oZcilX_GVAM3kwyGWhvbmY-12NN6y9Z9FyboKtvGcZ_us2L8RhKl8KBkYglJqfe7vGMRwhSYFVLEnjmRqNWHqKWMkmd3kkHff6DlHeAE2QoOXsk4kdNyGpyuE24gF0KwkH94nTrMfSoRMFGpIQ"
print(conn_str)
with Pyadomd(conn_str) as conn:
    with conn.cursor().execute(dax) as cur:
        columns = [col.name for col in cur.description]
        rows = cur.fetchall()

# Print results
print(columns)
for row in rows:
    print(row)
