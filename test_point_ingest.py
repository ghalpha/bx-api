import argparse
import json
from datetime import datetime, timezone
from pathlib import Path

import requests

AUTH_URL = "https://siemens-bt-015.eu.auth0.com/oauth/token"
AUDIENCE = "https://horizon.siemens.com"
INGEST_BASE_URL = "https://api.bpcloud.siemens.com/ingest"


def load_client_secrets():
    secrets_path = Path(__file__).with_name("secrets.json")
    if not secrets_path.exists():
        raise FileNotFoundError(
            f"Missing {secrets_path.name}. Create it from your client details first."
        )

    with secrets_path.open("r", encoding="utf-8") as f:
        secrets = json.load(f)

    client = secrets.get("client", {})
    client_id = client.get("client_id")
    client_secret = client.get("client_secret")
    partition_id = client.get("partition_id")

    missing = [
        key
        for key, value in {
            "client.client_id": client_id,
            "client.client_secret": client_secret,
            "client.partition_id": partition_id,
        }.items()
        if not value
    ]

    if missing:
        raise ValueError(f"Missing required secrets: {', '.join(missing)}")

    return client_id, client_secret, partition_id


def get_access_token(client_id, client_secret):
    payload = {
        "client_id": client_id,
        "client_secret": client_secret,
        "audience": AUDIENCE,
        "grant_type": "client_credentials",
    }
    headers = {"Content-Type": "application/json"}

    response = requests.post(AUTH_URL, json=payload, headers=headers, timeout=30)
    response.raise_for_status()
    token = response.json().get("access_token")

    if not token:
        raise RuntimeError("Auth response did not contain access_token")

    return token


def build_payload(value, timestamp, quality_of_value=None):
    attributes = {
        "timestamp": timestamp,
        "value": str(value),
    }

    if quality_of_value is not None:
        attributes["qualityOfValue"] = quality_of_value

    return {
        "data": [
            {
                "type": "PointValue",
                "attributes": attributes,
            }
        ]
    }


def ingest_point_value(token, partition_id, point_id, payload):
    url = f"{INGEST_BASE_URL}/partitions/{partition_id}/points/{point_id}/values"
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/vnd.api+json",
        "Accept": "application/vnd.api+json",
    }

    response = requests.post(url, json=payload, headers=headers, timeout=30)

    if response.status_code == 204:
        print("Success: point value ingested (HTTP 204).")
        return

    print(f"Ingest failed: HTTP {response.status_code}")
    try:
        print(json.dumps(response.json(), indent=2))
    except ValueError:
        print(response.text)

    response.raise_for_status()


def parse_args():
    parser = argparse.ArgumentParser(
        description="Test Siemens Building X Point Value Ingest API for one point."
    )
    parser.add_argument("--point-id", required=True, help="Target point ID (UUID).")
    parser.add_argument(
        "--value",
        required=True,
        help="Point value to ingest. Sent as a string per API schema.",
    )
    parser.add_argument(
        "--timestamp",
        default=datetime.now(timezone.utc).isoformat().replace("+00:00", "Z"),
        help="Timestamp in ISO 8601 date-time format (default: current UTC time).",
    )
    parser.add_argument(
        "--quality",
        type=int,
        choices=[0, 1, 2, 3],
        default=None,
        help="Optional qualityOfValue (0=Good, 1..3=Bad variants).",
    )
    return parser.parse_args()


def main():
    args = parse_args()
    client_id, client_secret, partition_id = load_client_secrets()
    token = get_access_token(client_id, client_secret)
    payload = build_payload(args.value, args.timestamp, args.quality)
    ingest_point_value(token, partition_id, args.point_id, payload)


if __name__ == "__main__":
    main()
