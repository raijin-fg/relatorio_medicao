import os
import time
from pyzabbix import ZabbixAPI
from db_manager import init_db, insert_zabbix_availability

_TARGET_GROUPS = ["CPTM - Switches", "CPTM - Roteadores", "CPTM - Servidores"]
_PING_KEY = "icmpping"
_PING_KEY_FALLBACK = "agent.ping"
_LOOKBACK_DAYS = 30


def _get_time_range() -> tuple[int, int]:
    now = int(time.time())
    return now - (_LOOKBACK_DAYS * 24 * 60 * 60), now


def _availability_row(group_name: str, host_name: str, item_display: str, availability) -> dict:
    return {
        "group_name": group_name,
        "host": host_name,
        "item": item_display,
        "availability_pct": availability,
    }


def _compute_availability(trends: list) -> float:
    if not trends:
        return 0.0
    values = [float(t["value_avg"]) for t in trends]
    return round((sum(values) / len(values)) * 100, 2)


def _get_ping_item_with_fallback(zapi: ZabbixAPI, hostid: str) -> tuple[dict, str] | tuple[None, None]:
    """Try to get ping item: first 'icmpping', then 'agent.ping'. Returns (item, key_used) or (None, None)."""
    for key in [_PING_KEY, _PING_KEY_FALLBACK]:
        items = zapi.item.get(
            hostids=hostid,
            filter={"key_": key},
            output=["itemid", "name"],
        )
        if items:
            return items[0], key
    return None, None


def _process_group(zapi: ZabbixAPI, group: dict, time_from: int, time_till: int) -> list:
    """Process a single group and return list of availability rows."""
    group_id = group["groupid"]
    group_name = group["name"]

    hosts = zapi.host.get(
        groupids=group_id,
        output=["hostid", "name"],
        monitored=True,
        filter={"status": 0}
    )

    if not hosts:
        return []

    rows = []
    for host in hosts:
        item, key_used = _get_ping_item_with_fallback(zapi, host["hostid"])

        if not item:
            rows.append(_availability_row(
                group_name,
                host["name"],
                f"Não encontrado ({_PING_KEY} / {_PING_KEY_FALLBACK})",
                "Sem monitoramento de Ping",
            ))
            continue

        trends = zapi.trend.get(
            itemids=item["itemid"],
            time_from=time_from,
            time_till=time_till,
            output=["value_avg"],
        )

        rows.append(_availability_row(
            group_name,
            host["name"],
            f"{item['name']} ({key_used})",
            _compute_availability(trends),
        ))

    return rows


def generate_availability_report(zapi: ZabbixAPI) -> bool:
    """Generate availability report and save to database. Returns True on success, False on error."""
    try:
        groups = zapi.hostgroup.get(filter={"name": _TARGET_GROUPS}, output=["groupid", "name"])

        if not groups:
            print("No target groups found on the Zabbix server.")
            return False

        time_from, time_till = _get_time_range()

        all_rows = []
        for group in groups:
            rows = _process_group(zapi, group, time_from, time_till)
            all_rows.extend(rows)

        init_db()
        insert_zabbix_availability(all_rows)
        print(f"Successfully inserted {len(all_rows)} Zabbix availability records into database")
        return True

    except Exception as e:
        print(f"Error generating availability report: {e}")
        return False


def extract_zabbix_data() -> bool:
    """Extract Zabbix data using credentials from environment variables."""
    try:
        zabbix_url = os.getenv("ZABBIX_URL")
        zabbix_token = os.getenv("ZABBIX_TOKEN")

        if not zabbix_url or not zabbix_token:
            print("Error: ZABBIX_URL or ZABBIX_TOKEN environment variables not set")
            return False

        zapi = ZabbixAPI(zabbix_url)
        zapi.login(api_token=zabbix_token)

        result = generate_availability_report(zapi)
        zapi.session.close()
        
        return result

    except Exception as e:
        print(f"Zabbix extraction failed: {e}")
        return False