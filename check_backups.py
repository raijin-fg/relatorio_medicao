"""
check_backups.py

Verifica, para um schedule de full escolhido (Full_Mensal ou Full_Semanal),
se todas as políticas ativas do NetBackup tiveram um job de backup rodado
e finalizado com sucesso (status 0) dentro do período analisado.

Uso:
    python check_backups.py --schedule semanal
    python check_backups.py --schedule mensal
    python check_backups.py --schedule semanal --days 10   # sobrescreve o lookback padrão

Variáveis de ambiente necessárias:
    NETBACKUP_MASTER_SERVER
    NETBACKUP_API_TOKEN
"""

import os
import sys
import argparse
from datetime import datetime, timedelta
from dotenv import load_dotenv

import requests
import urllib3

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

load_dotenv()

_NETBACKUP_PORT = 1556
_PAGE_LIMIT = 200

# Mapeamento do seletor de linha de comando -> nomes reais dos schedules no NetBackup
# (inclui a variante _Synthetic de cada um)
_SCHEDULE_MAP = {
    "mensal": ["Full_Mensal", "Full_Mensal_Synthetic"],
    "semanal": ["Full_Semanal", "Full_Semanal_Synthetic"],
}

# Lookback padrão (em dias) para cada tipo de schedule, caso --days não seja informado
_DEFAULT_DAYS = {
    "mensal": 30,
    "semanal": 7,
}


# ---------------------------------------------------------------------------
# Helpers reaproveitados da extração original (duplicados aqui para o script
# ser self-contained; se preferir, importe do módulo original em vez disso)
# ---------------------------------------------------------------------------

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


# ---------------------------------------------------------------------------
# Coleta de dados
# ---------------------------------------------------------------------------

def _is_policy_active(policy_name: str, attributes: dict) -> bool:
    """
    Verifica se uma política está ativa, procurando o campo 'active' em alguns
    lugares possíveis da resposta da API (a estrutura pode variar conforme a
    versão do NetBackup). Se não encontrar o campo em nenhum lugar conhecido,
    avisa no console e assume ativa (comportamento antigo) para não esconder
    políticas por engano.
    """
    policy_block = attributes.get("policy") or attributes

    candidates = [
        attributes.get("active"),
        policy_block.get("active"),
        policy_block.get("policyAttributes", {}).get("active") if isinstance(policy_block.get("policyAttributes"), dict) else None,
        attributes.get("policyAttributes", {}).get("active") if isinstance(attributes.get("policyAttributes"), dict) else None,
    ]

    for value in candidates:
        if value is not None:
            return str(value).lower() in ("true", "1")

    print(f"  [AVISO] Campo 'active' não encontrado para a política '{policy_name}'. "
          f"Assumindo ativa. Chaves disponíveis em attributes: {list(attributes.keys())}")
    return True


def get_active_policies(master_server: str, auth_token: str) -> list[str]:
    """Retorna a lista de nomes de políticas ativas."""
    params = {"page[limit]": _PAGE_LIMIT}
    active_policies = []

    with requests.Session() as session:
        session.headers.update(_netbackup_headers(auth_token))

        all_policies = _paginate(session, f"{_base_url(master_server)}/config/policies", params)

        for item in all_policies:
            policy_name = item.get("id")
            detail_url = f"{_base_url(master_server)}/config/policies/{policy_name}"

            detail_resp = session.get(detail_url, verify=False)
            if not detail_resp.ok:
                continue

            attributes = detail_resp.json().get("data", {}).get("attributes", {})

            if _is_policy_active(policy_name, attributes):
                active_policies.append(policy_name)

    return active_policies


def get_jobs_by_schedule(master_server: str, auth_token: str, schedule_names: list[str], days: int) -> list[dict]:
    """Retorna todos os jobs de backup (qualquer status) para os schedules e período informados."""
    cutoff = (datetime.now() - timedelta(days=days)).strftime("%Y-%m-%dT%H:%M:%S.000Z")

    schedule_clause = " or ".join(f"scheduleName eq '{name}'" for name in schedule_names)
    odata_filter = (
        f"jobType eq 'BACKUP' and ({schedule_clause}) and startTime gt {cutoff}"
    )
    params = {"filter": odata_filter, "page[limit]": _PAGE_LIMIT}

    with requests.Session() as session:
        session.headers.update(_netbackup_headers(auth_token))
        records = _paginate(session, f"{_base_url(master_server)}/admin/jobs", params)

    jobs = []
    for job in records:
        attr = job.get("attributes", {})
        jobs.append({
            "job_id":     attr.get("jobId"),
            "policy":     attr.get("policyName"),
            "schedule":   attr.get("scheduleName"),
            "state":      attr.get("state"),
            "status":     attr.get("status"),
            "start_time": attr.get("startTime"),
            "end_time":   attr.get("endTime"),
        })

    return jobs


# ---------------------------------------------------------------------------
# Comparação e relatório
# ---------------------------------------------------------------------------

def build_report(active_policies: list[str], jobs: list[dict]) -> dict:
    """
    Classifica cada política ativa em:
      - "ok"          -> tem job com state == DONE e status == 0
      - "falhou"      -> tem job(s), mas nenhum com state DONE e status 0
      - "nao_rodou"   -> nenhum job encontrado no período para essa política
    """
    jobs_by_policy: dict[str, list[dict]] = {}
    for job in jobs:
        jobs_by_policy.setdefault(job["policy"], []).append(job)

    report = {"ok": [], "falhou": [], "nao_rodou": []}

    for policy in active_policies:
        policy_jobs = jobs_by_policy.get(policy, [])

        if not policy_jobs:
            report["nao_rodou"].append(policy)
            continue

        success = any(
            str(j["state"]).upper() == "DONE" and j["status"] == 0
            for j in policy_jobs
        )

        if success:
            report["ok"].append(policy)
        else:
            report["falhou"].append({"policy": policy, "jobs": policy_jobs})

    return report


def print_report(report: dict, schedule_name: str, days: int) -> None:
    print("=" * 70)
    print(f"RELATÓRIO DE BACKUPS - schedule: {schedule_name} | últimos {days} dia(s)")
    print("=" * 70)

    print(f"\nPolíticas OK: {len(report['ok'])}")
    for p in report["ok"]:
        print(f"  [OK] {p}")

    print(f"\nPolíticas que NÃO RODARAM: {len(report['nao_rodou'])}")
    for p in report["nao_rodou"]:
        print(f"  [NÃO RODOU] {p}")

    print(f"\nPolíticas com job(s) que RODARAM mas NÃO FINALIZARAM OK: {len(report['falhou'])}")
    for entry in report["falhou"]:
        print(f"  [FALHOU] {entry['policy']}")
        for j in entry["jobs"]:
            print(
                f"      job_id={j['job_id']} state={j['state']} status={j['status']} "
                f"start={j['start_time']} end={j['end_time']}"
            )

    total_problemas = len(report["nao_rodou"]) + len(report["falhou"])
    print("\n" + "=" * 70)
    if total_problemas == 0:
        print("Resultado: todas as políticas ativas rodaram e finalizaram com sucesso.")
    else:
        print(f"Resultado: {total_problemas} política(s) com problema. Verifique acima.")
    print("=" * 70)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(description="Verifica status dos backups full por política.")
    parser.add_argument(
        "--schedule",
        required=True,
        choices=list(_SCHEDULE_MAP.keys()),
        help="Qual schedule de full analisar: 'mensal' (Full_Mensal) ou 'semanal' (Full_Semanal).",
    )
    parser.add_argument(
        "--days",
        type=int,
        default=None,
        help="Quantos dias para trás olhar (default: 35 para mensal, 7 para semanal).",
    )
    args = parser.parse_args()

    schedule_names = _SCHEDULE_MAP[args.schedule]
    days = args.days if args.days is not None else _DEFAULT_DAYS[args.schedule]

    master_server = os.getenv("NETBACKUP_MASTER_SERVER")
    auth_token = os.getenv("NETBACKUP_API_TOKEN")

    if not master_server:
        print("Error: NETBACKUP_MASTER_SERVER environment variable not set")
        sys.exit(1)

    if not auth_token:
        print("Error: NETBACKUP_API_TOKEN environment variable not set")
        sys.exit(1)

    try:
        active_policies = get_active_policies(master_server, auth_token)
        jobs = get_jobs_by_schedule(master_server, auth_token, schedule_names, days)
        report = build_report(active_policies, jobs)
        print_report(report, ", ".join(schedule_names), days)

        # Exit code != 0 se houver algum problema, útil para integração com cron/monitoramento
        if report["nao_rodou"] or report["falhou"]:
            sys.exit(2)

    except Exception as e:
        print(f"Erro ao checar backups: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
