import requests
import os
import pandas as pd
import json
import csv
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
    """Fetch all locations with their names and add debug output."""
    url = f"{BASE_API_URL}/partitions/{PARTITION_ID}/locations?filter[type]=Building&fields[Building]=label"
    headers = {"Authorization": f"Bearer {token}"}

    response = requests.get(url, headers=headers)

    if response.status_code == 200:
        locations = response.json().get("data", [])
        locations_data = []
        
        for loc in locations:
            loc_id = loc.get("id")
            loc_label = loc.get("attributes", {}).get("label", "N/A")  # Extract label as location name

            # Debugging output
            print(f"DEBUG: Found Location - ID: {loc_id}, Name: {loc_label}")

            if loc_id is None:
                print(f"ERROR: Location ID is None for location: {loc}")

            locations_data.append({
                "Location ID": loc_id,
                "Location Name": loc_label
            })

        return locations_data
    else:
        raise Exception(f"Error fetching locations: {response.status_code}, {response.text}")

def fetch_devices_for_location(token, location_id):
    """Fetch all devices for a given location, extracting device names and descriptions."""
    base_url = f"{BASE_API_URL}/partitions/{PARTITION_ID}/devices"
    headers = {"Authorization": f"Bearer {token}"}
    
    devices_data = []
    next_url = f"{base_url}?filter[hasLocation.data.id]={location_id}&page[limit]=50&include=hasFeatures.DeviceInfo"

    print(f"DEBUG: Fetching devices for Location ID: {location_id}")

    while next_url:
        response = requests.get(next_url, headers=headers)

        if response.status_code == 200:
            result = response.json()
            devices = result.get("data", [])
            included = result.get("included", [])

            # ✅ Build a mapping of DeviceInfo ID → Device Name & Description
            device_info_map = {}
            for info in included:
                if info.get("type") == "DeviceInfo":
                    device_id = info.get("relationships", {}).get("hasDevice", {}).get("data", {}).get("id")
                    device_name = info.get("attributes", {}).get("name", "N/A")
                    device_description = info.get("attributes", {}).get("description", "No Description")  # Extract description

                    if device_id:
                        device_info_map[device_id] = {
                            "name": device_name,
                            "description": device_description
                        }

            if not devices:
                print(f"DEBUG: No more devices found for Location ID: {location_id}")
                break  # No more devices, stop loop

            for dev in devices:
                dev_id = dev.get("id")

                # ✅ Extract Device Name and Description from `included` section
                device_data = device_info_map.get(dev_id, {"name": "N/A", "description": "No Description"})
                dev_name = device_data["name"]
                dev_description = device_data["description"]

                print(f"DEBUG: Found Device - ID: {dev_id}, Name: {dev_name}, Location ID: {location_id}")

                devices_data.append({
                    "Device ID": dev_id,
                    "Device Name": dev_name,
                    "Device Description": dev_description,  # ✅ Now adding description
                    "Location ID": location_id
                })

            # Handle pagination
            next_link = result.get("links", {}).get("next")

            if next_link:
                if next_link.startswith("/"):
                    next_url = f"{BASE_API_URL}{next_link}"  # Append base URL if necessary
                else:
                    next_url = next_link  # Use full URL if provided

                print(f"DEBUG: Fetching next page of devices for Location ID: {location_id}")
            else:
                print(f"DEBUG: No more pages for Location ID {location_id}")
                next_url = None  # Stop pagination
        elif response.status_code == 404:
            print(f"⚠️ WARNING: Received 404 on pagination request for Location ID {location_id}. Stopping pagination.")
            break  # Stop requesting pages if we get 404
        else:
            print(f"ERROR: Failed to fetch devices for Location ID {location_id}. Response: {response.status_code}, {response.text}")
            break  # Stop on other errors

    return devices_data

def fetch_points_for_device(token, device_id, device_name):
    """Fetch all points for a given device and associate them with the device name."""
    
    if device_id is None:
        print(f"ERROR: Trying to fetch points for a device with None ID! Device Name: {device_name}")
        return []
    
    base_url = f"{BASE_API_URL}/partitions/{PARTITION_ID}/devices/{device_id}/points"
    headers = {"Authorization": f"Bearer {token}"}

    points_data = []
    next_url = f"{base_url}?page[limit]=50"  # Increase page size for efficiency

    print(f"DEBUG: Fetching points for Device ID: {device_id}, Device Name: {device_name}")

    while next_url:
        response = requests.get(next_url, headers=headers)

        try:
            # ✅ Ensure response is valid JSON
            result = response.json()
        except json.JSONDecodeError as e:
            print(f"❌ ERROR: Failed to decode JSON for Device ID {device_id}. Possible invalid API response.")
            print(f"❌ Raw Response: {response.text[:500]}")  # Print first 500 chars to avoid huge logs
            break  # Stop pagination if response is invalid

        if response.status_code == 200:
            points = result.get("data", [])

            if not points:
                print(f"DEBUG: No more points found for Device ID: {device_id}")
                break  # No more points, stop loop

            for pt in points:
                pt_id = pt.get("id")
                pt_name = pt.get("attributes", {}).get("name", "N/A")
                pt_description = pt.get("attributes", {}).get("description", "")
                pt_activated = pt.get("attributes", {}).get("isActive", "")

                print(f"DEBUG: Found Point - ID: {pt_id}, Name: {pt_name}, Device Name: {device_name}")

                points_data.append({
                    "Point ID": pt_id,
                    "Point Name": pt_name,
                    "Point Description": pt_description,
                    "Point Activated": pt_activated,
                    "Device ID": device_id,
                    "Device Name": device_name
                })

            # Handle pagination
            next_link = result.get("links", {}).get("next")

            if next_link:
                if next_link.startswith("/"):
                    next_url = f"{BASE_API_URL}{next_link}"  # Append base URL if necessary
                else:
                    next_url = next_link  # Use full URL if provided

                print(f"DEBUG: Fetching next page of points for Device ID: {device_id}")
            else:
                print(f"DEBUG: No more pages for Device ID {device_id}")
                next_url = None  # Stop pagination
        elif response.status_code == 404:
            print(f"⚠️ WARNING: Received 404 on pagination request for Device ID {device_id}. Stopping pagination.")
            break  # Stop requesting pages if we get 404
        else:
            print(f"ERROR: Failed to fetch points for Device ID {device_id}. Response: {response.status_code}, {response.text}")
            break  # Stop on other errors

    return points_data

def main():
    token = get_access_token()

    # Define the output directory and CSV file path
    output_dir = "output"
    output_file = os.path.join(output_dir, "Building_Operations_Data.csv")

    # Ensure the output directory exists
    os.makedirs(output_dir, exist_ok=True)

    # Define column order
    column_order = [
        "Location ID", "Location Name", 
        "Device ID", "Device Name", "Device Description", 
        "Point ID", "Point Name", "Point Description", "Point Activated"
    ]

    # Open CSV file in append mode so data is written row-by-row
    with open(output_file, mode="w", newline='', encoding="utf-8") as csvfile:
        writer = csv.DictWriter(csvfile, fieldnames=column_order)

        # Write header row first
        writer.writeheader()

        # Fetch locations
        locations = fetch_locations(token)
        for loc in locations:
            loc_id = loc.get("Location ID")
            loc_name = loc.get("Location Name")

            if loc_id is None:
                print(f"ERROR: Trying to fetch devices for a location with None ID! Location Name: {loc_name}")
                continue  # Skip this location if no valid ID

            print(f"📍 Processing Location: {loc_name} (ID: {loc_id})")

            # Fetch devices for each location
            fetched_devices = fetch_devices_for_location(token, loc_id)

            if fetched_devices:
                for dev in fetched_devices:
                    dev_id = dev.get("Device ID")
                    dev_name = dev.get("Device Name")
                    dev_description = dev.get("Device Description")

                    if dev_id is None:
                        print(f"ERROR: Trying to fetch points for a device with None ID! Device Name: {dev_name}")
                        continue  # Skip this device if no valid ID

                    print(f"🔧 Fetching points for Device ID: {dev_id}, Name: {dev_name}")

                    # Fetch points for each device
                    fetched_points = fetch_points_for_device(token, dev_id, dev_name)

                    if fetched_points:
                        for pt in fetched_points:
                            row = {
                                "Location ID": loc_id,
                                "Location Name": loc_name,
                                "Device ID": dev_id,
                                "Device Name": dev_name,
                                "Device Description": dev_description,
                                "Point ID": pt.get("Point ID"),
                                "Point Name": pt.get("Point Name"),
                                "Point Description": pt.get("Point Description"), 
                                "Point Activated": pt.get("Point Activated")
                            }
                            writer.writerow(row)  # ✅ Writes row immediately to CSV
                    else:
                        # ✅ Write a row even if no points exist for this device
                        row = {
                            "Location ID": loc_id,
                            "Location Name": loc_name,
                            "Device ID": dev_id,
                            "Device Name": dev_name,
                            "Device Description": dev_description,
                            "Point ID": "",
                            "Point Name": "",
                            "Point Description": "",
                            "Point Activated": ""
                        }
                        writer.writerow(row)

    print(f"✅ CSV file saved at: {os.path.abspath(output_file)}")

if __name__ == "__main__":
    main()
