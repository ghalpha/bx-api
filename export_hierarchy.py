import requests
import pandas as pd
import json
from pathlib import Path

# API Endpoints
AUTH_URL = "https://siemens-bt-015.eu.auth0.com/oauth/token"
BASE_API_URL = "https://api.bpcloud.siemens.com/operations"
AUDIENCE = "https://horizon.siemens.com"

def load_auth_secrets():
    secrets_path = Path(__file__).with_name("secrets.json")
    with secrets_path.open("r", encoding="utf-8") as f:
        secrets = json.load(f)
    client = secrets.get("client", {})
    return client.get("client_id"), client.get("client_secret"), client.get("partition_id")


CLIENT_ID, CLIENT_SECRET, PARTITION_ID = load_auth_secrets()

def get_access_token():
    """Retrieve an OAuth2 Bearer token."""
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
        raise Exception(f"Error getting token: {response.status_code}, {response.text}")

def fetch_locations(token):
    """Fetch all locations."""
    url = f"{BASE_API_URL}/partitions/{PARTITION_ID}/locations?filter[type]=Building&fields[Building]=label"
    headers = {"Authorization": f"Bearer {token}"}
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        return response.json().get("data", [])
    else:
        raise Exception(f"Error fetching locations: {response.status_code}, {response.text}")

def fetch_devices_for_location(token, location_id):
    """Fetch all devices for a given location."""
    url = f"{BASE_API_URL}/partitions/{PARTITION_ID}/devices?filter[location.id]={location_id}"
    headers = {"Authorization": f"Bearer {token}"}
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        return response.json().get("data", [])
    else:
        raise Exception(f"Error fetching devices for location {location_id}: {response.status_code}, {response.text}")

def fetch_points_for_device(token, device_id):
    """Fetch all points for a given device."""
    url = f"{BASE_API_URL}/partitions/{PARTITION_ID}/devices/{device_id}/points"
    headers = {"Authorization": f"Bearer {token}"}
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        return response.json().get("data", [])
    else:
        raise Exception(f"Error fetching points for device {device_id}: {response.status_code}, {response.text}")

def main():
    token = get_access_token()

    # Data storage
    locations_data = []
    devices_data = []
    points_data = []

    # Fetch locations
    locations = fetch_locations(token)
    for loc in locations:
        loc_id = loc.get("id")
        loc_name = loc.get("attributes", {}).get("label", "N/A")  # Use 'label' instead of 'name'
        locations_data.append({"Location ID": loc_id, "Location Name": loc_name})

        # Fetch devices for each location
        devices = fetch_devices_for_location(token, loc_id)
        for dev in devices:
            dev_id = dev.get("id")
    
            # ✅ Extract the correct device name from profile.name
            dev_profile = dev.get("attributes", {}).get("profile", {})  # Extract profile object
            dev_name = dev_profile.get("name", "N/A")  # Extract device name from profile

            devices_data.append({
            "Device ID": dev_id, 
            "Device Name": dev_name,  # ✅ Correctly assigned from profile.name
            "Location ID": loc_id, 
            "Location Name": loc_name
        })


            # Fetch points for each device
            points = fetch_points_for_device(token, dev_id)
            for pt in points:
                pt_id = pt.get("id")
                pt_name = pt.get("attributes", {}).get("name", "N/A")
                
                # ✅ Now including Device Name & Location Name
                points_data.append({
                    "Point ID": pt_id, 
                    "Point Name": pt_name, 
                    "Device ID": dev_id, 
                    "Device Name": dev_name, 
                    "Location Name": loc_name
                })

    # Create DataFrames
    df_locations = pd.DataFrame(locations_data)
    df_devices = pd.DataFrame(devices_data)
    df_points = pd.DataFrame(points_data)

    # Write to Excel with multiple sheets
    with pd.ExcelWriter("Building_Operations_Data.xlsx", engine="openpyxl") as writer:
        df_locations.to_excel(writer, sheet_name="Locations", index=False)
        df_devices.to_excel(writer, sheet_name="Devices", index=False)
        df_points.to_excel(writer, sheet_name="Points", index=False)  # ✅ Now includes Device & Location Names

    print("Data has been successfully written to 'Building_Operations_Data.xlsx'.")

if __name__ == "__main__":
    main()
