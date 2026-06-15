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


def _compute_availability(events: list, time_from: int, time_till: int) -> float:
    """
    Calcula a disponibilidade (SLA) percentual de um host baseando-se no histórico
    de eventos (incidentes) críticos de ICMP/Acessibilidade dentro do período.
    
    Mesmo que a disponibilidade seja baixíssima ou zero, o valor real será retornado.
    """
    # 1. Definir o tempo total da amostragem (em segundos)
    tempo_total_periodo_segundos = time_till - time_from
    
    if tempo_total_periodo_segundos <= 0:
        return 100.0

    total_downtime_segundos = 0.0

    # 2. Varrer os incidentes para consolidar o tempo offline
    for ev in events:
        event_name = ev.get("name", "")
        
        # Filtro de garantia: Garante que estamos avaliando apenas o que afeta conectividade
        if "icmp" not in event_name.lower() and "indispon" not in event_name.lower() and "não acess" not in event_name.lower():
            continue

        # Normalização dos timestamps de início e fim do incidente
        start_time = int(ev["clock"])
        r_eventid = ev.get("r_eventid", "0")
        
        end_time = None
        
        # 1ª Tentativa: Checar se possui evento de recuperação explícito
        if r_eventid and r_eventid != "0":
            # Nota: No loop principal do relatório, certifique-se de que o 'r_clock' 
            # já foi coletado ou faça o fetch dele aqui se necessário.
            # Assumindo a lógica blindada do script de teste:
            pass
        
        # 2ª Tentativa: Checar se o fechamento veio via ações de atualização/sistema
        if not end_time and ev.get("acknowledges"):
            for ack in ev["acknowledges"]:
                if int(ack.get("action", 0)) & 1:  # Action 1 = Sistema fechou
                    end_time = int(ack["clock"])
                    break

        # Delimitação do tempo dentro do escopo do relatório mensal
        # Se o incidente começou antes do mês, limitamos o início ao time_from
        start_calculado = max(start_time, time_from)
        
        if end_time and end_time > start_calculado:
            # Se terminou depois do fim do período, limitamos ao time_till
            end_calculado = min(end_time, time_till)
            duracao_segundos = end_calculado - start_calculado
        else:
            if start_time < time_from:
                duracao_segundos = time_till - time_from  # Contagem total do mês para incidentes persistentes
            else:
            # Fallback seguro para micro-quedas sem registro de fim ou limitação ao limite atual
                duracao_segundos = min(time_till - start_calculado, 300) # Máximo 5 min por flape sem fim

        total_downtime_segundos += duracao_segundos

    # 3. Trava matemática lógica (O downtime máximo não pode passar do tempo do próprio mês)
    if total_downtime_segundos > tempo_total_periodo_segundos:
        total_downtime_segundos = tempo_total_periodo_segundos

    # 4. Cálculo final do SLA real (Trabalha em qualquer faixa de 0.00% a 100.00%)
    tempo_online_segundos = tempo_total_periodo_segundos - total_downtime_segundos
    availability_pct = (tempo_online_segundos / tempo_total_periodo_segundos) * 100

    return round(availability_pct, 2)


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
    """Process a single group and return list of availability rows based on real events and persistent blocks."""
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
        # Mantemos o fallback para garantir que o host possui o item de ping ativo
        item, key_used = _get_ping_item_with_fallback(zapi, host["hostid"])

        if not item:
            rows.append(_availability_row(
                group_name,
                host["name"],
                f"Não encontrado ({_PING_KEY} / {_PING_KEY_FALLBACK})",
                "Sem monitoramento de Ping",
            ))
            continue

        # ---------------------------------------------------------------------
        # 1. PEGAR EVENTOS QUE INICIARAM DENTRO DO MÊS ATUAL
        # ---------------------------------------------------------------------
        events = zapi.event.get(
            hostids=host["hostid"],
            time_from=time_from,
            time_till=time_till,
            source=0,               # Eventos originados de Triggers
            value=1,                # Apenas eventos que INICIARAM um problema
            severities=[3, 4, 5],   # Média, Alta, Desastre
            output=["eventid", "clock", "name", "r_eventid"],
            select_acknowledges=["clock", "action"],
            sortfield="clock",
            sortorder="ASC"
        )

        # ---------------------------------------------------------------------
        # 2. CONTINGÊNCIA PARA PEGAR INCIDENTES ANTIGOS QUE CONTINUAM ABERTOS (OBRAS, ETC)
        # ---------------------------------------------------------------------
        # Busca na tabela de problemas ativos se há algo para este host nas severidades críticas
        active_problems = zapi.problem.get(
            hostids=host["hostid"],
            severities=[3, 4, 5],   # Mesmas severidades do SLA
            source=0,               # Triggers
            output=["eventid", "clock", "name"]
        )

        for prob in active_problems:
            prob_name = prob.get("name", "")
            # Filtro fino de string complementar por garantia
            if "icmp" not in prob_name.lower() and "indispon" not in prob_name.lower() and "não acess" not in prob_name.lower():
                continue
                
            prob_clock = int(prob["clock"])
            
            # Se o problema começou ANTES do início do nosso mês atual...
            if prob_clock < time_from:
                # Criamos um "evento fantasma" estruturado para a nossa função de cálculo entender.
                # Como ele começou antes e ainda está ativo (tabela de problems), 
                # dizemos que o clock de início dele é o próprio clock original da abertura.
                # O 'r_eventid' fica "0" (não resolvido), forçando o downtime a contar o mês todo.
                persistent_event = {
                    "eventid": prob["eventid"],
                    "clock": str(prob_clock),
                    "name": prob_name,
                    "r_eventid": "0",
                    "acknowledges": []
                }
                
                # Evita duplicar se por algum acaso do destino ele já estivesse na lista
                if not any(e["eventid"] == prob["eventid"] for e in events):
                    events.append(persistent_event)

        # ---------------------------------------------------------------------
        # 3. CÁLCULO E INSERÇÃO NO COCKPIT DA PLANILHA / BANCO
        # ---------------------------------------------------------------------
        availability_pct = _compute_availability(events, time_from, time_till)

        rows.append(_availability_row(
            group_name,
            host["name"],
            f"{item['name']} ({key_used})",
            availability_pct,
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