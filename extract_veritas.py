import os
import requests
import winrm
import urllib3
from datetime import datetime, timedelta
from db_manager import init_db, insert_netbackup_jobs, insert_netbackup_policies, insert_netbackup_tapes

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

_NETBACKUP_PORT = 1556
_WINRM_PORT = 5985
_PAGE_LIMIT = 200
_DAYS_LOOKBACK = 30

_SCHEDULE_FILTER = "(scheduleName eq 'Full_Mensal' or scheduleName eq 'Full_Mensal_Synthetic')"


def _netbackup_headers(auth_token: str) -> dict:
    return {
        "Authorization": auth_token,
        "Accept": "application/vnd.netbackup+json;version=12.0",
        "Content-Type": "application/vnd.netbackup+json",
    }


def _base_url(master_server: str) -> str:
    return f"https://{master_server}:{_NETBACKUP_PORT}/netbackup"


def _paginate(session: requests.Session, url: str, params: dict) -> list:
    collected = []
    page_params = dict(params)

    while True:
        response = session.get(url, params=page_params, verify=False)
        response.raise_for_status()

        payload = response.json()
        page_data = payload.get("data", [])

        if not page_data:
            break

        collected.extend(page_data)

        cursor = payload.get("meta", {}).get("page", {}).get("cursor")
        if cursor and len(page_data) == page_params.get("page[limit]", _PAGE_LIMIT):
            page_params["page[after]"] = cursor
        else:
            break

    return collected


def extract_netbackup_jobs(master_server: str, auth_token: str) -> bool:
    """Extract NetBackup jobs and save to database. Returns True on success, False on error."""
    cutoff = (datetime.now() - timedelta(days=_DAYS_LOOKBACK)).strftime("%Y-%m-%dT%H:%M:%S.000Z")

    odata_filter = (
        f"jobType eq 'BACKUP' and state eq 'DONE' and status eq 0 and "
        f"{_SCHEDULE_FILTER} and startTime gt {cutoff}"
    )

    params = {"filter": odata_filter, "page[limit]": _PAGE_LIMIT}

    try:
        with requests.Session() as session:
            session.headers.update(_netbackup_headers(auth_token))
            records = _paginate(session, f"{_base_url(master_server)}/admin/jobs", params)

        rows = [
            {
                "job_id":       attr.get("jobId"),
                "job_policy":   attr.get("policyName"),
                "job_schedule": attr.get("scheduleName"),
                "start_time":   attr.get("startTime"),
                "end_time":     attr.get("endTime"),
                "elapsed_time": attr.get("elapsedTime"),
                "kilobytes":    attr.get("kilobytesTransferred"),
                "kb_sec":       attr.get("transferRate"),
            }
            for job in records
            for attr in [job.get("attributes", {})]
        ]

        init_db()
        insert_netbackup_jobs(rows)
        print(f"Successfully inserted {len(rows)} NetBackup jobs into database")
        return True

    except Exception as e:
        print(f"NetBackup jobs extraction error: {e}")
        return False


def extract_netbackup_policies(master_server: str, auth_token: str) -> bool:
    """Extract NetBackup policies and save to database. Returns True on success, False on error."""
    params = {"page[limit]": _PAGE_LIMIT}

    try:
        with requests.Session() as session:
            session.headers.update(_netbackup_headers(auth_token))

            all_policies = _paginate(session, f"{_base_url(master_server)}/config/policies", params)

            if not all_policies:
                print("No NetBackup policies found")
                return True

            rows = []
            for item in all_policies:
                policy_name = item.get("id")
                detail_url = f"{_base_url(master_server)}/config/policies/{policy_name}"

                detail_resp = session.get(detail_url, verify=False)
                if not detail_resp.ok:
                    continue

                attributes = detail_resp.json().get("data", {}).get("attributes", {})
                policy_block = attributes.get("policy") or attributes

                is_active = str(policy_block.get("active", True)).lower() in ("true", "1")
                if not is_active:
                    continue

                selections_raw = policy_block.get("backupSelections", {})
                if isinstance(selections_raw, dict):
                    selections = selections_raw.get("selections", [])
                elif isinstance(selections_raw, list):
                    selections = selections_raw
                else:
                    selections = []

                rows.append({
                    "policy_name":       policy_name,
                    "backup_selections": "; ".join(str(s) for s in selections) if selections else "[]",
                })

        init_db()
        insert_netbackup_policies(rows)
        print(f"Successfully inserted {len(rows)} NetBackup policies into database")
        return True

    except Exception as e:
        print(f"NetBackup policies extraction error: {e}")
        return False


def extract_tape_report(master_server: str, winrm_user: str, winrm_password: str) -> bool:
    """Extract tape report and save to database. Returns True on success, False on error."""
    _VMQUERY_PATH  = r"D:\Program Files\Veritas\Volmgr\bin\vmquery.exe"
    _BPMEDIA_PATH  = r"D:\Program Files\Veritas\NetBackup\bin\admincmd\bpmedialist.exe"

    try:
        session = winrm.Session(
            f"http://{master_server}:{_WINRM_PORT}/wsman",
            auth=(winrm_user, winrm_password),
            transport="ntlm",
        )

        bp_result = session.run_ps(f'& "{_BPMEDIA_PATH}" -mlist -l -d hcart2')

        kb_by_media: dict[str, int] = {}
        if bp_result.status_code == 0:
            for line in bp_result.std_out.decode("cp1252", errors="ignore").splitlines():
                parts = line.split()
                if len(parts) > 8:
                    try:
                        kb_by_media[parts[0]] = int(parts[8])
                    except ValueError:
                        kb_by_media[parts[0]] = 0

        vm_result = session.run_ps(f'& "{_VMQUERY_PATH}" -a -bx')
        if vm_result.status_code != 0:
            print(f"vmquery execution failed (exit code {vm_result.status_code})")
            return False

        rows = []
        for line in vm_result.std_out.decode("cp1252", errors="ignore").splitlines():
            parts = line.split()

            if not parts or parts[0].lower() == "media":
                continue

            if len(parts) >= 14 and parts[1].upper() == "HCART2":
                media_id   = parts[0]
                media_type = parts[1]
                robot      = parts[2]
                pool       = parts[-1]
                kilobytes  = kb_by_media.get(media_id, 0)

                rows.append([media_id, media_type, robot, pool, kilobytes])

        init_db()
        insert_netbackup_tapes(rows)
        print(f"Successfully inserted {len(rows)} NetBackup tapes into database")
        return True

    except Exception as e:
        print(f"Tape report extraction error: {e}")
        return False


def extract_netbackup_data() -> bool:
    """Extract all NetBackup data using credentials from environment variables."""
    try:
        master_server = os.getenv("NETBACKUP_MASTER_SERVER")
        auth_token = os.getenv("NETBACKUP_API_TOKEN")
        winrm_user = os.getenv("NETBACKUP_USER")
        winrm_password = os.getenv("NETBACKUP_PASS")

        if not master_server:
            print("Error: NETBACKUP_MASTER_SERVER environment variable not set")
            return False

        if not auth_token:
            print("Error: NETBACKUP_API_TOKEN environment variable not set")
            return False

        if not winrm_user or not winrm_password:
            print("Error: NETBACKUP_USER or NETBACKUP_PASS environment variables not set")
            return False

        success = True

        if not extract_netbackup_jobs(master_server, auth_token):
            success = False

        if not extract_netbackup_policies(master_server, auth_token):
            success = False

        if not extract_tape_report(master_server, winrm_user, winrm_password):
            success = False

        return success

    except Exception as e:
        print(f"NetBackup data extraction failed: {e}")
        return False