import os
import msal
import requests
import jwt
from sys import path

# Your Azure AD App Registration details
CLIENT_ID = '54640219-af44-42bb-adcd-d3722aa55e04'  # Application (client) ID
# CLIENT_ID = '7f67af8a-fedc-4b08-8b4e-37c4d127b6cf'
TENANT_ID = 'e39cce29-5716-43ba-b27d-1bdd8fd67901'  # Or your tenant ID
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
AUTHORITY = f"https://login.microsoftonline.com/common"
# SCOPES = ["https://analysis.windows.net/powerbi/api/.default", "https://analysis.windows.net/powerbi/api/user_impersonation"]
SCOPES = ["https://analysis.windows.net/powerbi/api/.default"]
# SCOPES = ["https://analysis.windows.net/powerbi/api/user_impersonation"]

CACHE_FILE = "token_cache.bin"
force_reauthentication = True  # Set to True to force re-authentication

# Create persistent cache
cache = msal.SerializableTokenCache()
if os.path.exists(CACHE_FILE):
    cache.deserialize(open(CACHE_FILE, "r").read())

app = msal.PublicClientApplication(
    client_id=CLIENT_ID,
    # authority="https://login.microsoftonline.com/common",
    authority=AUTHORITY,
    token_cache=cache
)

# # Create the public client app
# app = msal.PublicClientApplication(
#     client_id=CLIENT_ID,
#     authority=AUTHORITY
# )

# Launch browser-based interactive login
# result = app.acquire_token_interactive(scopes=SCOPES, prompt='select_account')


# Attempt to acquire token silently (cached)
accounts = app.get_accounts()
if accounts and not force_reauthentication:
    result = app.acquire_token_silent(SCOPES, account=accounts[0])
else:
    result = None

# If silent fails, do interactive login with prompt
if not result:
    result = app.acquire_token_interactive(
        scopes=SCOPES,
        prompt="select_account"  # Always allow user to switch
    )

# Save cache back to file
with open(CACHE_FILE, "w") as f:
    f.write(cache.serialize())

# claims = result.get("id_token_claims", {})

# print("Logged in as:", claims.get("name") or claims.get("preferred_username") or claims.get("upn") or "Unknown user")

access_token = result["access_token"]
decoded = jwt.decode(access_token, options={"verify_signature": False})

print("User in access token:", decoded.get("preferred_username") or decoded.get("upn") or decoded.get("name"))


# Show user info
# print("Logged in as:", result.get("name") or result.get("preferred_username"))


SCOPE = ["https://analysis.windows.net/powerbi/api/.default"]
API_BASE = "https://api.powerbi.com/v1.0/myorg"

# # ==== INTERACTIVE LOGIN ====
# app = msal.PublicClientApplication(client_id=CLIENT_ID, authority=AUTHORITY)

# result = app.acquire_token_interactive(scopes=SCOPE)

# if "access_token" not in result:
#     print("Failed to acquire token.")
#     print(result.get("error_description"))
#     exit(1)

access_token = result['access_token']
headers = {
    'Authorization': f'Bearer {access_token}'
}


dax = "EVALUATE Apps"

from pyadomd import Pyadomd
from pyadomd import *

access_token = 'eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiIsIng1dCI6Il9qTndqZVNudlRUSzhYRWRyNVFVUGtCUkxMbyIsImtpZCI6Il9qTndqZVNudlRUSzhYRWRyNVFVUGtCUkxMbyJ9.eyJhdWQiOiJodHRwczovL2FuYWx5c2lzLndpbmRvd3MubmV0L3Bvd2VyYmkvYXBpIiwiaXNzIjoiaHR0cHM6Ly9zdHMud2luZG93cy5uZXQvZTM5Y2NlMjktNTcxNi00M2JhLWIyN2QtMWJkZDhmZDY3OTAxLyIsImlhdCI6MTc1MTQ3MDA2MiwibmJmIjoxNzUxNDcwMDYyLCJleHAiOjE3NTE0NzUxMzcsImFjY3QiOjAsImFjciI6IjEiLCJhaW8iOiJBVVFBdS84WkFBQUF2V3FJWU5BR2dZdXFrbGRrOU95OVpxZUVFdlB4R2Y2R2ZFNlZzSUpRNUdJVkdyUHJTNTNVQjRwU2JaVVRYeU1Ja3pJamhCMjR3eThIL2ZaRGRtemtZUT09IiwiYW1yIjpbInB3ZCJdLCJhcHBpZCI6IjdmNjdhZjhhLWZlZGMtNGIwOC04YjRlLTM3YzRkMTI3YjZjZiIsImFwcGlkYWNyIjoiMCIsImlkdHlwIjoidXNlciIsImlwYWRkciI6IjEzNS4xMzEuMTMzLjE0IiwibmFtZSI6ImZ0dTI2Iiwib2lkIjoiYzhiZTY1YmMtM2RkOS00NDNmLWFiNDYtYjZlMWQ1OGYxYTIzIiwicHVpZCI6IjEwMDMyMDA0Q0VCOUVBM0QiLCJyaCI6IjEuQVJnQUtjNmM0eFpYdWtPeWZSdmRqOVo1QVFrQUFBQUFBQUFBd0FBQUFBQUFBQUJxQVFnWUFBLiIsInNjcCI6InVzZXJfaW1wZXJzb25hdGlvbiIsInNpZCI6IjAwNmFmZWI5LWQzZjYtNjNjNC01M2YwLWY3ODRmOTcyM2JhZiIsInN1YiI6Ikc1eVJnaS1BTm5FR1dBZmtUdlJvM1FaUTJFa203SDRTOTN3TmtzdXZYNEEiLCJ0aWQiOiJlMzljY2UyOS01NzE2LTQzYmEtYjI3ZC0xYmRkOGZkNjc5MDEiLCJ1bmlxdWVfbmFtZSI6ImZ0dTI2QFNreWxpbmVEYXRhQW5hbHl0aWNzLm9ubWljcm9zb2Z0LmNvbSIsInVwbiI6ImZ0dTI2QFNreWxpbmVEYXRhQW5hbHl0aWNzLm9ubWljcm9zb2Z0LmNvbSIsInV0aSI6IjZweGl3Y2l6OWtleUlEUnVISGdRQUEiLCJ2ZXIiOiIxLjAiLCJ3aWRzIjpbImI3OWZiZjRkLTNlZjktNDY4OS04MTQzLTc2YjE5NGU4NTUwOSJdLCJ4bXNfZnRkIjoiVW16d1ZPRHBFVXd6Z1dxd0N1R08wZFRzMm5TZy03S05DZ0dKX3dPU3Vjc0JkWE56YjNWMGFDMWtjMjF6IiwieG1zX2lkcmVsIjoiMSAxNiJ9.flXtekKmGEiZ9C3rnQTgVL4IgYETCr0aCHOhMO8CkebHkTWSqQD_8FsUshkUG1yDgNkoSaru7aaYMCHQGeFVr24Y2To0ndi3ocYnEbCQBAxXIRmbzP5DJ_jYVXyQobxYWBTDP35DR0vW0_030iWjuRtDURau6QZmsQEovKJpousMsUikoFBZsm7ZZB-ltGzYKzWCh56LeNoNwrrYoZiV_WORYYiEhPUvrQmQwZExhRH8x3fS2PuxyeO4prKBIYkUj3uHg66PPT3o9zWoT39IxTaVM8HzIUDduxURSB-QPPWCFks1LDiC4LOd2wqikWZ0UXnKY2Ymies66CrqD2NXSg'
conn_str = f'Provider=MSOLAP;Data Source=pbiazure://api.powerbi.com/;Initial Catalog=sobe_wowvirtualserver-37b188f0-d623-4d60-b032-8a1ef55be1fb;Integrated Security=ClaimsToken;Persist Security Info=True;Identity Provider=\"https://login.microsoftonline.com/common, https://analysis.windows.net/powerbi/api, 54640219-af44-42bb-adcd-d3722aa55e04\";Password={access_token}'

conn_str = 'Provider=MSOLAP;Data Source=pbiazure://api.powerbi.com/;Initial Catalog=sobe_wowvirtualserver-37b188f0-d623-4d60-b032-8a1ef55be1fb;Integrated Security=ClaimsToken;Persist Security Info=True;Identity Provider=\"https://login.microsoftonline.com/common, https://analysis.windows.net/powerbi/api, 7f67af8a-fedc-4b08-8b4e-37c4d127b6cf\";Password=eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiIsIng1dCI6Il9qTndqZVNudlRUSzhYRWRyNVFVUGtCUkxMbyIsImtpZCI6Il9qTndqZVNudlRUSzhYRWRyNVFVUGtCUkxMbyJ9.eyJhdWQiOiJodHRwczovL2FuYWx5c2lzLndpbmRvd3MubmV0L3Bvd2VyYmkvYXBpIiwiaXNzIjoiaHR0cHM6Ly9zdHMud2luZG93cy5uZXQvZTM5Y2NlMjktNTcxNi00M2JhLWIyN2QtMWJkZDhmZDY3OTAxLyIsImlhdCI6MTc1MTQ3MDA2MiwibmJmIjoxNzUxNDcwMDYyLCJleHAiOjE3NTE0NzUxMzcsImFjY3QiOjAsImFjciI6IjEiLCJhaW8iOiJBVVFBdS84WkFBQUF2V3FJWU5BR2dZdXFrbGRrOU95OVpxZUVFdlB4R2Y2R2ZFNlZzSUpRNUdJVkdyUHJTNTNVQjRwU2JaVVRYeU1Ja3pJamhCMjR3eThIL2ZaRGRtemtZUT09IiwiYW1yIjpbInB3ZCJdLCJhcHBpZCI6IjdmNjdhZjhhLWZlZGMtNGIwOC04YjRlLTM3YzRkMTI3YjZjZiIsImFwcGlkYWNyIjoiMCIsImlkdHlwIjoidXNlciIsImlwYWRkciI6IjEzNS4xMzEuMTMzLjE0IiwibmFtZSI6ImZ0dTI2Iiwib2lkIjoiYzhiZTY1YmMtM2RkOS00NDNmLWFiNDYtYjZlMWQ1OGYxYTIzIiwicHVpZCI6IjEwMDMyMDA0Q0VCOUVBM0QiLCJyaCI6IjEuQVJnQUtjNmM0eFpYdWtPeWZSdmRqOVo1QVFrQUFBQUFBQUFBd0FBQUFBQUFBQUJxQVFnWUFBLiIsInNjcCI6InVzZXJfaW1wZXJzb25hdGlvbiIsInNpZCI6IjAwNmFmZWI5LWQzZjYtNjNjNC01M2YwLWY3ODRmOTcyM2JhZiIsInN1YiI6Ikc1eVJnaS1BTm5FR1dBZmtUdlJvM1FaUTJFa203SDRTOTN3TmtzdXZYNEEiLCJ0aWQiOiJlMzljY2UyOS01NzE2LTQzYmEtYjI3ZC0xYmRkOGZkNjc5MDEiLCJ1bmlxdWVfbmFtZSI6ImZ0dTI2QFNreWxpbmVEYXRhQW5hbHl0aWNzLm9ubWljcm9zb2Z0LmNvbSIsInVwbiI6ImZ0dTI2QFNreWxpbmVEYXRhQW5hbHl0aWNzLm9ubWljcm9zb2Z0LmNvbSIsInV0aSI6IjZweGl3Y2l6OWtleUlEUnVISGdRQUEiLCJ2ZXIiOiIxLjAiLCJ3aWRzIjpbImI3OWZiZjRkLTNlZjktNDY4OS04MTQzLTc2YjE5NGU4NTUwOSJdLCJ4bXNfZnRkIjoiVW16d1ZPRHBFVXd6Z1dxd0N1R08wZFRzMm5TZy03S05DZ0dKX3dPU3Vjc0JkWE56YjNWMGFDMWtjMjF6IiwieG1zX2lkcmVsIjoiMSAxNiJ9.flXtekKmGEiZ9C3rnQTgVL4IgYETCr0aCHOhMO8CkebHkTWSqQD_8FsUshkUG1yDgNkoSaru7aaYMCHQGeFVr24Y2To0ndi3ocYnEbCQBAxXIRmbzP5DJ_jYVXyQobxYWBTDP35DR0vW0_030iWjuRtDURau6QZmsQEovKJpousMsUikoFBZsm7ZZB-ltGzYKzWCh56LeNoNwrrYoZiV_WORYYiEhPUvrQmQwZExhRH8x3fS2PuxyeO4prKBIYkUj3uHg66PPT3o9zWoT39IxTaVM8HzIUDduxURSB-QPPWCFks1LDiC4LOd2wqikWZ0UXnKY2Ymies66CrqD2NXSg'
print(conn_str)


clr.AddReference('Microsoft.AnalysisServices.AdomdClient')
from Microsoft.AnalysisServices.AdomdClient import AdomdConnection, AdomdCommand # type: ignore


conn = AdomdConnection(conn_str)
conn.ConnectionString = conn_str

with Pyadomd(conn_str) as conn:
    with conn.cursor().execute(dax) as cur:
        columns = [col.name for col in cur.description]
        rows = cur.fetchall()
