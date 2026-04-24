import requests
import json
from pathlib import Path

# Define the endpoint and payload
AUTH_URL = "https://siemens-bt-015.eu.auth0.com/oauth/token"
AUDIENCE = "https://horizon.siemens.com"


def load_auth_secrets():
    secrets_path = Path(__file__).with_name("secrets.json")
    with secrets_path.open("r", encoding="utf-8") as f:
        secrets = json.load(f)
    client = secrets.get("client", {})
    return client.get("client_id"), client.get("client_secret")


CLIENT_ID, CLIENT_SECRET = load_auth_secrets()

def get_access_token():
    payload = {
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "audience": AUDIENCE,
        "grant_type": "client_credentials"
    }
    
    headers = {"Content-Type": "application/json"}

    response = requests.post(AUTH_URL, json=payload, headers=headers)
    
    if response.status_code == 200:
        return response.json().get("access_token")
    else:
        print(f"Error: {response.status_code}, {response.text}")
        return None

# Example usage:
token = get_access_token()
if token:
    print("Access Token:", token)