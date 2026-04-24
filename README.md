# Building X Operations API Scripts

Utilities for authenticating and testing Siemens Building X APIs.

## Included Scripts

| Script | Purpose |
|---|---|
| `get_token.py` | Fetches and prints an OAuth2 access token. Useful for quickly verifying credentials or grabbing a token for manual API calls. |
| `get_info.py` | Walks the operations hierarchy (locations → devices → points) and exports three CSV files to the `output/` folder. Includes device descriptions and handles paginated API responses. |
| `export_hierarchy.py` | Earlier version of the hierarchy export. Walks locations → devices → points and writes the results to a single Excel workbook (`Building_Operations_Data.xlsx`) with one sheet per resource type. |
| `test_point_ingest.py` | Posts a single point value to the Point Value Ingest API. Accepts a point ID, numeric value, optional timestamp, and optional quality flag via CLI arguments. |

## Prerequisites

- Python 3.12+
- A Siemens Building X client with API access

## Setup

1. Create and activate a virtual environment.
2. Install dependencies:

   ```bash
   pip install -r requirements.txt
   ```

3. Create local secrets file from the example:

   PowerShell:

   ```powershell
   Copy-Item secrets.example.json secrets.json
   ```

   Bash:

   ```bash
   cp secrets.example.json secrets.json
   ```

4. Edit `secrets.json` with your client details:

   - `client.client_id`
   - `client.client_secret`
   - `client.partition_id`

## Test Point Ingest

Post one value using the current UTC timestamp by default:

```bash
python test_point_ingest.py --point-id <POINT_ID> --value 29.92
```

Optional arguments:

- `--timestamp 2026-04-24T18:30:00Z`
- `--quality 0`

## Security

- `secrets.json` is intentionally gitignored and should never be committed.
- Share only `secrets.example.json` in source control.

## Notes

- The ingest endpoint expects content type `application/vnd.api+json`.
- A `403` response usually means the client is authenticated but not authorized for ingest on the target partition/point.
