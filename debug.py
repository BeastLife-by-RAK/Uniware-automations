import requests
import os
from dotenv import load_dotenv

load_dotenv()

TENANT_URL = os.getenv("UNIWARE_TENANT_URL")
USERNAME   = os.getenv("UNIWARE_USERNAME")
PASSWORD   = os.getenv("UNIWARE_PASSWORD")
CLIENT_ID  = os.getenv("UNIWARE_CLIENT_ID", "my-trusted-client")

# Step 1: Get token
print("--- Fetching token ---")
token_url = f"{TENANT_URL}/oauth/token"
params = {
    "grant_type": "password",
    "client_id":  CLIENT_ID,
    "username":   USERNAME,
    "password":   PASSWORD,
}
r = requests.get(token_url, params=params, timeout=30)
print(f"Status: {r.status_code}")
print(f"Response: {r.text}")

# Step 2: Hit inventory with whatever token we got
if r.status_code == 200:
    token = r.json().get("access_token")
    print(f"\nToken: {token}")

    print("\n--- Hitting inventory endpoint ---")
    inv_url = f"{TENANT_URL}/services/rest/v1/inventory/inventorySnapshot/get"
    headers = {
        "Content-Type":  "application/json",
        "Authorization": f"Bearer {token}",
    }
    r2 = requests.post(inv_url, json={}, headers=headers, timeout=60)
    print(f"Status: {r2.status_code}")
    print(f"Response: {r2.text[:500]}")