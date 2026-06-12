import os
import urllib3
import requests
from db_manager import init_db, insert_netapp_volumes

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

_BYTES_TO_GB = 1024 ** 3

def _bytes_to_gb(value: int) -> float:
    return round(value / _BYTES_TO_GB, 2)

def extract_netapp_volumes(cluster_ip: str, api_user: str, api_password: str) -> bool:
    """Extract NetApp volumes and save to database. Returns True on success, False on error."""
    url = f"https://{cluster_ip}/api/storage/volumes"
    params = {
        "fields": "name,space.size,space.used,space.available",
        "return_records": "true",
    }

    try:
        response = requests.get(
            url,
            auth=(api_user, api_password),
            params=params,
            verify=False,
            timeout=15,
        )
        response.raise_for_status()

        records = response.json().get("records", [])

        if not records:
            print("No NetApp volumes found")
            return True

        rows = []
        for volume in records:
            space = volume.get("space", {})
            rows.append([
                volume.get("name"),
                _bytes_to_gb(space.get("size", 0)),
                _bytes_to_gb(space.get("used", 0)),
                _bytes_to_gb(space.get("available", 0)),
            ])

        init_db()
        insert_netapp_volumes(rows)
        print(f"Successfully inserted {len(rows)} NetApp volumes into database")
        return True

    except requests.HTTPError as e:
        print(f"NetApp API error ({e.response.status_code}): {e.response.text}")
        return False
    except Exception as e:
        print(f"Unexpected error during NetApp extraction: {e}")
        return False


def extract_netapp_data() -> bool:
    """Extract NetApp data using credentials from environment variables."""
    try:
        # Get cluster IP from environment - this should be set separately
        cluster_ip = os.getenv("NETAPP_CLUSTER_IP")
        api_user = os.getenv("NETAPP_USER")
        api_password = os.getenv("NETAPP_PASS")

        if not cluster_ip or not api_user or not api_password:
            print("Error: NETAPP_CLUSTER_IP, NETAPP_USER, or NETAPP_PASS environment variables not set")
            return False

        return extract_netapp_volumes(cluster_ip, api_user, api_password)

    except Exception as e:
        print(f"NetApp data extraction failed: {e}")
        return False