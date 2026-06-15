"""
CPTM - Gerador de Relatório Mensal de TI
Lê dados do banco SQLite e gera relatório Excel em output/
"""

import sqlite3
import os
import re
from datetime import datetime, timezone
from pathlib import Path
from isodate import parse_duration
import isodate

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side, GradientFill
)
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.series import DataPoint
from openpyxl.formatting.rule import ColorScaleRule, CellIsRule, FormulaRule
from openpyxl.styles.numbers import FORMAT_PERCENTAGE_00

# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------
DB_PATH = Path(__file__).parent / "data" / "prd_data.db"
OUTPUT_DIR = Path(__file__).parent / "output"
NUM_DRIVERS = 6
MAX_NETAPP_HOURS = 10.0   # duração alvo máxima por política NetApp (horas)

# Cores
C_HEADER_BG   = "1F3864"   # azul escuro
C_HEADER_FT   = "FFFFFF"
C_SECTION_BG  = "D6E4F0"   # azul claro
C_SECTION_FT  = "1F3864"
C_ALT_ROW     = "F2F7FC"
C_GREEN_BG    = "C6EFCE"
C_GREEN_FT    = "276221"
C_YELLOW_BG   = "FFEB9C"
C_YELLOW_FT   = "9C6500"
C_RED_BG      = "FFC7CE"
C_RED_FT      = "9C0006"
C_ORANGE_BG   = "FFD966"
C_ORANGE_FT   = "7D4E00"
C_GRAY_BG     = "D9D9D9"
C_WHITE       = "FFFFFF"
C_ACCENT      = "2E75B6"

FONT_NAME = "Arial"

# ---------------------------------------------------------------------------
# Helpers de estilo
# ---------------------------------------------------------------------------

def _font(bold=False, size=10, color="000000", italic=False):
    return Font(name=FONT_NAME, bold=bold, size=size, color=color, italic=italic)

def _fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)

def _border(style="thin"):
    s = Side(style=style)
    return Border(left=s, right=s, top=s, bottom=s)

def _align(h="left", v="center", wrap=False):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

def _header_row(ws, row, values, widths=None, bg=C_HEADER_BG, ft=C_HEADER_FT, size=10):
    for col, val in enumerate(values, 1):
        cell = ws.cell(row=row, column=col, value=val)
        cell.font = _font(bold=True, size=size, color=ft)
        cell.fill = _fill(bg)
        cell.alignment = _align("center")
        cell.border = _border()
    if widths:
        for col, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(col)].width = w

def _data_row(ws, row, values, alt=False):
    bg = C_ALT_ROW if alt else C_WHITE
    for col, val in enumerate(values, 1):
        cell = ws.cell(row=row, column=col, value=val)
        cell.font = _font()
        cell.fill = _fill(bg)
        cell.alignment = _align()
        cell.border = _border()

def _section_title(ws, row, text, ncols, bg=C_SECTION_BG, ft=C_SECTION_FT):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=ncols)
    cell = ws.cell(row=row, column=1, value=text)
    cell.font = _font(bold=True, size=11, color=ft)
    cell.fill = _fill(bg)
    cell.alignment = _align("left")
    cell.border = _border()

def _set_col_widths(ws, widths):
    for col, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(col)].width = w

def _iso_to_hours(iso_str):
    """Converte string ISO 8601 duration (PT13H47M2S) para horas float."""
    try:
        return isodate.parse_duration(iso_str).total_seconds() / 3600
    except Exception:
        return 0.0

def _iso_to_hhmm(iso_str):
    """Converte ISO duration para string HH:MM legível."""
    try:
        total_sec = int(isodate.parse_duration(iso_str).total_seconds())
        h, rem = divmod(total_sec, 3600)
        m = rem // 60
        return f"{h:02d}:{m:02d}"
    except Exception:
        return "00:00"

def _parse_dt(iso_str):
    """Parse ISO 8601 datetime para datetime ciente de UTC."""
    if not iso_str:
        return None
    return datetime.fromisoformat(iso_str.replace("Z", "+00:00"))

def _kb_to_tb(kb):
    return kb / (1024 ** 3)

def _kb_to_gb(kb):
    return kb / (1024 ** 2)

def _sla_color(pct):
    if pct is None:
        return C_GRAY_BG, "000000"
    if pct >= 99.9:
        return C_GREEN_BG, C_GREEN_FT
    if pct >= 99.0:
        return C_YELLOW_BG, C_YELLOW_FT
    return C_RED_BG, C_RED_FT

# ---------------------------------------------------------------------------
# Leitura do banco
# ---------------------------------------------------------------------------

def _load(db_path):
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    try:
        cursor = conn.cursor()


        avail   = conn.execute("SELECT * FROM zabbix_availability ORDER BY group_name, host").fetchall()
        jobs    = conn.execute("SELECT * FROM netbackup_jobs ORDER BY start_time").fetchall()
        pols    = conn.execute("SELECT * FROM netbackup_policies").fetchall()
        tapes   = conn.execute("SELECT * FROM netbackup_tapes ORDER BY snapshot_date").fetchall()
        volumes = conn.execute("""
            SELECT * FROM netapp_volumes
            ORDER BY snapshot_date, volume_name
        """).fetchall()

        return avail, jobs, pols, tapes, volumes
    finally:
        conn.close()

# ---------------------------------------------------------------------------
# Aba 1 — Capa / Resumo Executivo
# ---------------------------------------------------------------------------

def _sheet_capa(wb, avail, jobs, volumes, report_month):
    ws = wb.create_sheet("Capa")
    ws.sheet_view.showGridLines = False

    # Título principal
    ws.merge_cells("A1:H2")
    c = ws["A1"]
    c.value = "RELATÓRIO MENSAL DE TI — CPTM"
    c.font = _font(bold=True, size=18, color=C_HEADER_FT)
    c.fill = _fill(C_HEADER_BG)
    c.alignment = _align("center")

    ws.merge_cells("A3:H3")
    c = ws["A3"]
    c.value = f"Competência: {report_month}"
    c.font = _font(bold=True, size=12, color=C_SECTION_FT)
    c.fill = _fill(C_SECTION_BG)
    c.alignment = _align("center")

    ws.row_dimensions[1].height = 28
    ws.row_dimensions[2].height = 28

    # ---- KPIs de disponibilidade por grupo ----
    row = 5
    _section_title(ws, row, "DISPONIBILIDADE — RESUMO POR GRUPO", 8)

    row += 1
    _header_row(ws, row, ["Grupo", "Total Hosts", "≥ 99,9%", "≥ 99,0%", "< 99,0%", "Offline (0%)", "Uptime Médio", "SLA Status"])

    groups = {}
    for r in avail:
        g = r["group_name"]
        pct = r["availability_pct"]
        if isinstance(pct, str):
            pct = None
        if g not in groups:
            groups[g] = []
        groups[g].append(pct)

    row += 1
    for g, vals in sorted(groups.items()):
        numeric = [v for v in vals if v is not None]
        total   = len(vals)
        above999 = sum(1 for v in numeric if v >= 99.9)
        above99  = sum(1 for v in numeric if 99.0 <= v < 99.9)
        below99  = sum(1 for v in numeric if v < 99.0)
        offline  = sum(1 for v in numeric if v == 0.0)
        avg      = sum(numeric) / len(numeric) if numeric else None

        bg, ft = _sla_color(avg)
        sla_txt = "✔ OK" if avg and avg >= 99.9 else ("⚠ ATENÇÃO" if avg and avg >= 99.0 else "✖ CRÍTICO")

        for col, val in enumerate([g, total, above999, above99, below99, offline,
                                    f"{avg:.4f}%" if avg is not None else "N/A", sla_txt], 1):
            cell = ws.cell(row=row, column=col, value=val)
            cell.font = _font(bold=(col == 8), color=ft if col in (7, 8) else "000000")
            cell.fill = _fill(bg if col in (7, 8) else C_WHITE)
            cell.alignment = _align("center")
            cell.border = _border()
        row += 1

    # ---- KPIs de backup ----
    row += 1
    _section_title(ws, row, "BACKUP — RESUMO DO MÊS", 8)

    row += 1
    _header_row(ws, row, ["Métrica", "Valor"])

    total_jobs  = len(jobs)
    total_tb    = sum(_kb_to_tb(j["kilobytes"]) for j in jobs)
    netapp_jobs = [j for j in jobs if j["job_policy"].startswith("NETAPP")]
    vmware_jobs = [j for j in jobs if j["job_policy"].startswith("VMWARES")]
    oracle_jobs = [j for j in jobs if "ORACLE" in j["job_policy"]]
    exch_jobs   = [j for j in jobs if "EXCHANGE" in j["job_policy"]]

    metrics = [
        ("Total de jobs executados no mês", total_jobs),
        ("Volume total transferido (TB)", f"{total_tb:.2f}"),
        ("Volume NetApp (TB)", f"{sum(_kb_to_tb(j['kilobytes']) for j in netapp_jobs):.2f}"),
        ("Volume VMware (TB)", f"{sum(_kb_to_tb(j['kilobytes']) for j in vmware_jobs):.2f}"),
        ("Volume Oracle (TB)", f"{sum(_kb_to_tb(j['kilobytes']) for j in oracle_jobs):.2f}"),
        ("Volume Exchange (TB)", f"{sum(_kb_to_tb(j['kilobytes']) for j in exch_jobs):.2f}"),
        ("Job mais longo", max((_iso_to_hhmm(j["elapsed_time"]) for j in jobs), default="—")),
        ("Política mais longa",  max(jobs, key=lambda j: _iso_to_hours(j["elapsed_time"]))["job_policy"] if jobs else "—"),
    ]

    row += 1
    for i, (label, val) in enumerate(metrics):
        bg = C_ALT_ROW if i % 2 else C_WHITE
        for col, v in enumerate([label, val], 1):
            cell = ws.cell(row=row, column=col, value=v)
            cell.font = _font(bold=(col == 1))
            cell.fill = _fill(bg)
            cell.alignment = _align()
            cell.border = _border()
        row += 1

    # ---- NetApp storage ----
    row += 1
    _section_title(ws, row, "ARMAZENAMENTO NETAPP — SNAPSHOT MAIS RECENTE", 8)
    row += 1
    _header_row(ws, row, ["Métrica", "Valor"])

    latest = max((v["snapshot_date"] for v in volumes), default=None)
    latest_vols = [v for v in volumes if v["snapshot_date"] == latest] if latest else []
    total_alloc = sum(v["total_gb"] for v in latest_vols)
    total_used  = sum(v["used_gb"]  for v in latest_vols)
    total_free  = sum(v["free_gb"]  for v in latest_vols)
    pct_used    = (total_used / total_alloc * 100) if total_alloc else 0

    stor_metrics = [
        ("Data do snapshot", latest or "—"),
        ("Capacidade total alocada (TB)", f"{total_alloc/1024:.2f}"),
        ("Utilizado (TB)",  f"{total_used/1024:.2f}"),
        ("Disponível (TB)", f"{total_free/1024:.2f}"),
        ("% Utilizado", f"{pct_used:.1f}%"),
        ("Volumes monitorados", len(latest_vols)),
    ]
    row += 1
    for i, (label, val) in enumerate(stor_metrics):
        bg = C_ALT_ROW if i % 2 else C_WHITE
        for col, v in enumerate([label, val], 1):
            cell = ws.cell(row=row, column=col, value=v)
            cell.font = _font(bold=(col == 1))
            cell.fill = _fill(bg)
            cell.alignment = _align()
            cell.border = _border()
        row += 1

    _set_col_widths(ws, [38, 14, 10, 10, 10, 12, 14, 14])

# ---------------------------------------------------------------------------
# Aba 2 — Disponibilidade (icmpping)
# ---------------------------------------------------------------------------

def _sheet_disponibilidade(wb, avail):
    ws = wb.create_sheet("Disponibilidade")
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A3"

    _header_row(ws, 1,
        ["Grupo", "Host", "Disponibilidade (%)", "SLA", "Observação"],
        widths=[28, 55, 22, 14, 30])

    # Ordenar: grupo asc, disponibilidade asc (piores primeiro dentro do grupo)
    rows_sorted = sorted(avail, key=lambda r: (
        r["group_name"],
        r["availability_pct"] if isinstance(r["availability_pct"], float) else -1
    ))

    data_start = 2
    for i, r in enumerate(rows_sorted):
        row = data_start + i
        pct = r["availability_pct"]
        pct_num = pct if isinstance(pct, float) else None
        pct_str = f"{pct_num:.4f}%" if pct_num is not None else str(pct)

        bg, ft = _sla_color(pct_num)

        if pct_num is None:
            sla = "Sem dados"
            obs = "Item não retornou valor numérico"
        elif pct_num == 0.0:
            sla = "OFFLINE"
            obs = "Host indisponível durante todo o período"
        elif pct_num >= 99.9:
            sla = "✔ OK"
            obs = ""
        elif pct_num >= 99.0:
            sla = "⚠ ATENÇÃO"
            obs = f"Downtime ≈ {(100 - pct_num) / 100 * 43200:.1f} min no mês"
        else:
            sla = "✖ CRÍTICO"
            obs = f"Downtime ≈ {(100 - pct_num) / 100 * 43200:.0f} min no mês"

        alt = i % 2 == 1
        row_bg = C_ALT_ROW if alt else C_WHITE

        vals = [r["group_name"], r["host"], pct_str, sla, obs]
        for col, val in enumerate(vals, 1):
            cell = ws.cell(row=row, column=col, value=val)
            if col in (3, 4):
                cell.font = _font(bold=(col == 4), color=ft)
                cell.fill = _fill(bg)
            else:
                cell.font = _font()
                cell.fill = _fill(row_bg)
            cell.alignment = _align("center" if col in (3, 4) else "left")
            cell.border = _border()

    # Resumo por grupo no final
    data_end = data_start + len(rows_sorted)
    row = data_end + 1
    _section_title(ws, row, "RESUMO POR GRUPO", 5)
    row += 1
    _header_row(ws, row, ["Grupo", "Hosts", "Uptime Médio", "Piores hosts (< 99%)", ""])

    groups = {}
    for r in avail:
        g = r["group_name"]
        pct = r["availability_pct"]
        if g not in groups:
            groups[g] = []
        if isinstance(pct, float):
            groups[g].append((r["host"], pct))

    row += 1
    for g, items in sorted(groups.items()):
        vals_only = [v for _, v in items]
        avg = sum(vals_only) / len(vals_only) if vals_only else None
        worst = [h for h, v in sorted(items, key=lambda x: x[1])[:3] if v < 99.0]
        worst_str = "; ".join(worst) if worst else "—"
        bg, ft = _sla_color(avg)
        for col, val in enumerate([g, len(items), f"{avg:.4f}%" if avg else "N/A", worst_str, ""], 1):
            cell = ws.cell(row=row, column=col, value=val)
            cell.font = _font(bold=(col in (1, 3)), color=ft if col == 3 else "000000")
            cell.fill = _fill(bg if col == 3 else C_WHITE)
            cell.alignment = _align("center" if col in (2, 3) else "left")
            cell.border = _border()
        row += 1

# ---------------------------------------------------------------------------
# Aba 3 — Backup Geral
# ---------------------------------------------------------------------------

def _sheet_backup_geral(wb, jobs):
    ws = wb.create_sheet("Backup — Visão Geral")
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A3"

    headers = ["Política", "Schedule", "Início", "Término", "Duração",
               "Volume (GB)", "Volume (TB)", "Velocidade (MB/s)", "Categoria"]
    widths  = [28, 22, 22, 22, 12, 14, 12, 18, 18]
    _header_row(ws, 1, headers, widths)

    def _categoria(policy):
        p = policy.upper()
        if p.startswith("NETAPP"):   return "NetApp"
        if p.startswith("VMWARES"): return "VMware"
        if "ORACLE" in p:           return "Oracle"
        if "EXCHANGE" in p:         return "Exchange"
        if p.startswith("VMWARE"):  return "VMware"
        return "Windows/Outros"

    jobs_sorted = sorted(jobs, key=lambda j: j["start_time"] or "")

    for i, j in enumerate(jobs_sorted):
        row = 2 + i
        alt = i % 2 == 1
        start = _parse_dt(j["start_time"])
        end   = _parse_dt(j["end_time"])
        gb    = _kb_to_gb(j["kilobytes"])
        tb    = _kb_to_tb(j["kilobytes"])
        mbps  = j["kb_sec"] / 1024 if j["kb_sec"] else 0

        vals = [
            j["job_policy"],
            j["job_schedule"],
            start.astimezone().strftime("%d/%m/%Y %H:%M") if start else "",
            end.astimezone().strftime("%d/%m/%Y %H:%M") if end else "",
            _iso_to_hhmm(j["elapsed_time"]),
            round(gb, 2),
            round(tb, 4),
            round(mbps, 1),
            _categoria(j["job_policy"]),
        ]
        _data_row(ws, row, vals, alt)

    # Totais por categoria
    data_end = 1 + len(jobs_sorted)
    row = data_end + 2
    _section_title(ws, row, "TOTAIS POR CATEGORIA", len(headers))
    row += 1
    _header_row(ws, row, ["Categoria", "Qtd Jobs", "Volume Total (TB)", "Duração Total (h)", "Velocidade Média (MB/s)"])

    cats = {}
    for j in jobs:
        cat = _categoria(j["job_policy"])
        cats.setdefault(cat, []).append(j)

    row += 1
    for cat, jlist in sorted(cats.items()):
        total_tb  = sum(_kb_to_tb(j["kilobytes"]) for j in jlist)
        total_h   = sum(_iso_to_hours(j["elapsed_time"]) for j in jlist)
        avg_mbps  = sum((j["kb_sec"] or 0) / 1024 for j in jlist) / len(jlist)
        for col, val in enumerate([cat, len(jlist), round(total_tb, 2),
                                    round(total_h, 1), round(avg_mbps, 1)], 1):
            cell = ws.cell(row=row, column=col, value=val)
            cell.font = _font(bold=(col == 1))
            cell.fill = _fill(C_ALT_ROW if row % 2 else C_WHITE)
            cell.alignment = _align("center" if col > 1 else "left")
            cell.border = _border()
        row += 1

    # Top 10 jobs mais longos
    row += 1
    _section_title(ws, row, "TOP 10 — JOBS MAIS LONGOS", len(headers))
    row += 1
    _header_row(ws, row, ["Política", "Duração", "Volume (TB)", "Velocidade (MB/s)", "Início"])

    top10 = sorted(jobs, key=lambda j: _iso_to_hours(j["elapsed_time"]), reverse=True)[:10]
    row += 1
    for i, j in enumerate(top10):
        alt = i % 2 == 1
        start = _parse_dt(j["start_time"])
        for col, val in enumerate([
            j["job_policy"],
            _iso_to_hhmm(j["elapsed_time"]),
            round(_kb_to_tb(j["kilobytes"]), 4),
            round((j["kb_sec"] or 0) / 1024, 1),
            start.astimezone().strftime("%d/%m/%Y %H:%M") if start else "",
        ], 1):
            cell = ws.cell(row=row, column=col, value=val)
            cell.font = _font()
            cell.fill = _fill(C_YELLOW_BG if i == 0 else (C_ALT_ROW if alt else C_WHITE))
            cell.alignment = _align("center" if col > 1 else "left")
            cell.border = _border()
        row += 1

# ---------------------------------------------------------------------------
# Aba 4 — Otimização de Janelas (Gantt + análise de drivers)
# ---------------------------------------------------------------------------

def _sheet_janelas(wb, jobs):
    ws = wb.create_sheet("Otimização — Janelas")
    ws.sheet_view.showGridLines = False

    # ---- Tabela principal de análise ----
    _section_title(ws, 1, "ANÁLISE DE JANELAS DE BACKUP — UTILIZAÇÃO DE DRIVERS", 9)

    headers = ["Política", "Início", "Término", "Duração (h)",
               "Volume (TB)", "Concorrentes no pico", "Conflito de drivers?",
               "Sugestão de início", "Observação"]
    widths  = [28, 20, 20, 14, 12, 22, 20, 20, 38]
    _header_row(ws, 2, headers, widths)

    # Monta lista de jobs com datetimes
    job_list = []
    for j in jobs:
        start = _parse_dt(j["start_time"])
        end   = _parse_dt(j["end_time"])
        if not start or not end:
            continue
        job_list.append({
            "policy":   j["job_policy"],
            "schedule": j["job_schedule"],
            "start":    start,
            "end":      end,
            "hours":    _iso_to_hours(j["elapsed_time"]),
            "tb":       _kb_to_tb(j["kilobytes"]),
            "elapsed":  j["elapsed_time"],
        })

    # Para cada job, conta quantos outros rodam simultaneamente
    def _concurrent_count(job, all_jobs):
        return sum(
            1 for j in all_jobs
            if j is not job and j["start"] < job["end"] and j["end"] > job["start"]
        )

    # Agrupa por política (pega o de maior duração se houver múltiplos)
    policy_jobs = {}
    for j in job_list:
        p = j["policy"]
        if p not in policy_jobs or j["hours"] > policy_jobs[p]["hours"]:
            policy_jobs[p] = j

    representative = list(policy_jobs.values())
    representative.sort(key=lambda j: j["start"])

    # Heatmap de uso de drivers por hora (baseado em UTC local)
    hour_counts = [0] * 24
    for j in job_list:
        s = j["start"].replace(tzinfo=None)
        e = j["end"].replace(tzinfo=None)
        cur = s
        while cur < e:
            hour_counts[cur.hour] += 1
            cur = cur.replace(minute=0, second=0) + pd.Timedelta(hours=1)  # type: ignore

    peak_hour = max(range(24), key=lambda h: hour_counts[h])

    row = 3
    for i, j in enumerate(representative):
        conc = _concurrent_count(j, job_list)
        conflict = conc >= NUM_DRIVERS
        conflict_str = f"⚠ Sim ({conc} jobs)" if conflict else f"Não ({conc} jobs)"

        # Sugestão simples: hora de menor ocupação próxima ao início atual
        cur_hour = j["start"].hour
        sorted_hours = sorted(range(24), key=lambda h: hour_counts[h])
        suggested_h = next((h for h in sorted_hours if h != peak_hour), cur_hour)
        suggestion = f"{suggested_h:02d}:00" if conflict else "—"
        obs = (f"Considere mover para {suggestion} — menor concorrência" if conflict
               else "Janela adequada")

        alt = i % 2 == 1
        bg_conflict = C_RED_BG if conflict else (C_ALT_ROW if alt else C_WHITE)

        vals = [
            j["policy"],
            j["start"].astimezone().strftime("%d/%m/%Y %H:%M"),
            j["end"].astimezone().strftime("%d/%m/%Y %H:%M"),
            round(j["hours"], 2),
            round(j["tb"], 3),
            conc,
            conflict_str,
            suggestion,
            obs,
        ]
        for col, val in enumerate(vals, 1):
            cell = ws.cell(row=row, column=col, value=val)
            cell.font = _font(bold=(col == 7 and conflict), color=(C_RED_FT if (col == 7 and conflict) else "000000"))
            cell.fill = _fill(bg_conflict if col in (6, 7) else (C_ALT_ROW if alt else C_WHITE))
            cell.alignment = _align("center" if col in (2, 3, 4, 5, 6, 7, 8) else "left")
            cell.border = _border()
        row += 1

    # ---- Heatmap de ocupação por hora ----
    row += 1
    _section_title(ws, row, "HEATMAP — JOBS ATIVOS POR FAIXA HORÁRIA (horário local)", 9)
    row += 1
    _header_row(ws, row, [f"{h:02d}h" for h in range(24)] + ["Limite drivers"])
    row += 1
    for col, cnt in enumerate(hour_counts, 1):
        cell = ws.cell(row=row, column=col, value=cnt)
        cell.font = _font(bold=(cnt > NUM_DRIVERS), color=C_RED_FT if cnt > NUM_DRIVERS else "000000")
        cell.fill = _fill(C_RED_BG if cnt > NUM_DRIVERS else (C_YELLOW_BG if cnt == NUM_DRIVERS else C_GREEN_BG))
        cell.alignment = _align("center")
        cell.border = _border()
        ws.column_dimensions[get_column_letter(col)].width = 6
    # Coluna "Limite drivers"
    cell = ws.cell(row=row, column=25, value=NUM_DRIVERS)
    cell.font = _font(bold=True)
    cell.fill = _fill(C_SECTION_BG)
    cell.alignment = _align("center")
    cell.border = _border()
    ws.column_dimensions[get_column_letter(25)].width = 16

# ---------------------------------------------------------------------------
# Aba 5 — Otimização NetApp (split de políticas)
# ---------------------------------------------------------------------------

def _parse_netapp_volumes_from_policy(backup_selections):
    """Extrai nomes de volumes de backup_selections."""
    if not backup_selections:
        return []
    volumes = []
    for part in backup_selections.split(";"):
        part = part.strip()
        if not part or part.upper().startswith("SET"):
            continue
        # Remove prefixo /cptm_nas/ e barra final
        name = re.sub(r"^/?cptm_nas/?", "", part, flags=re.IGNORECASE).strip("/")
        if name:
            volumes.append(name.upper())
    return volumes

def _sheet_netapp_split(wb, jobs, pols, volumes):
    ws = wb.create_sheet("Otimização — NetApp Split")
    ws.sheet_view.showGridLines = False

    # Mapa volume -> tamanho (último snapshot)
    latest_date = max((v["snapshot_date"] for v in volumes), default=None)
    vol_size = {}
    for v in volumes:
        if v["snapshot_date"] == latest_date:
            vol_size[v["volume_name"].upper()] = v["used_gb"]

    # Mapa de políticas NetApp e seus volumes
    netapp_pols = {}
    for p in pols:
        if p["policy_name"].upper().startswith("NETAPP") and p["backup_selections"]:
            vols = _parse_netapp_volumes_from_policy(p["backup_selections"])
            if vols:
                netapp_pols[p["policy_name"]] = vols

    # Jobs NetApp para calcular velocidade média real
    netapp_jobs = [j for j in jobs if j["job_policy"].upper().startswith("NETAPP")]
    avg_kbsec = (sum(j["kb_sec"] for j in netapp_jobs if j["kb_sec"]) /
                 max(1, sum(1 for j in netapp_jobs if j["kb_sec"])))
    avg_gbsec = avg_kbsec / (1024 ** 2)  # KB/s -> GB/s

    # ---- Seção 1: Mapa atual de volumes por política ----
    _section_title(ws, 1, "MAPA ATUAL — VOLUMES POR POLÍTICA NETAPP", 8)
    _header_row(ws, 2,
        ["Política", "Volume", "Tamanho Usado (GB)", "% da política",
         "Duração estimada (h)", "Duração real último job (h)", "Status", ""],
        widths=[18, 22, 22, 16, 22, 26, 14, 14])

    row = 3
    policy_summaries = {}
    for pol_name in sorted(netapp_pols):
        vols = netapp_pols[pol_name]
        total_gb_pol = sum(vol_size.get(v, 0) for v in vols)

        # Duração real do último job dessa política
        pol_jobs = [j for j in jobs if j["job_policy"] == pol_name]
        real_h = max((_iso_to_hours(j["elapsed_time"]) for j in pol_jobs), default=None)

        policy_summaries[pol_name] = {
            "volumes": vols,
            "total_gb": total_gb_pol,
            "real_h": real_h,
            "est_h": total_gb_pol / (avg_gbsec * 3600) if avg_gbsec > 0 else 0,
        }

        # Linha de cabeçalho da política
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=8)
        cell = ws.cell(row=row, column=1,
            value=f"  {pol_name}  —  {len(vols)} volumes  |  {total_gb_pol:.1f} GB usado  |  Duração real: {_iso_to_hhmm(pol_jobs[-1]['elapsed_time']) if pol_jobs else '—'}")
        cell.font = _font(bold=True, size=10, color=C_SECTION_FT)
        cell.fill = _fill(C_SECTION_BG)
        cell.border = _border()
        row += 1

        for i, vol in enumerate(vols):
            gb = vol_size.get(vol, None)
            pct = (gb / total_gb_pol * 100) if (gb and total_gb_pol) else None
            est_h = (gb / (avg_gbsec * 3600)) if (gb and avg_gbsec > 0) else None
            alt = i % 2 == 1
            status = "✔ No banco" if gb else "⚠ Não encontrado"
            for col, val in enumerate([
                pol_name, vol,
                round(gb, 1) if gb else "—",
                f"{pct:.1f}%" if pct else "—",
                round(est_h, 2) if est_h else "—",
                round(real_h, 2) if real_h else "—",
                status, "",
            ], 1):
                cell = ws.cell(row=row, column=col, value=val)
                cell.font = _font(color=(C_RED_FT if status.startswith("⚠") and col == 7 else "000000"))
                cell.fill = _fill(C_ALT_ROW if alt else C_WHITE)
                cell.alignment = _align("center" if col > 2 else "left")
                cell.border = _border()
            row += 1

    # ---- Seção 2: Sugestão de split ----
    row += 1
    _section_title(ws, row, f"SUGESTÃO DE SPLIT — META: MÁXIMO {MAX_NETAPP_HOURS}h POR POLÍTICA", 8)
    row += 1
    _header_row(ws, row,
        ["Nova política sugerida", "Volumes", "Total GB", "Duração estimada (h)",
         "Duração atual (h)", "Diferença (h)", "Recomendação", ""])
    row += 1

    # Coleta todos os volumes de todas as políticas NetApp com seus tamanhos
    all_vol_entries = []
    for pol_name, info in policy_summaries.items():
        for vol in info["volumes"]:
            gb = vol_size.get(vol, 0)
            all_vol_entries.append({"vol": vol, "gb": gb, "orig_pol": pol_name})

    # Bin packing greedy: ordena por tamanho desc, aloca em bins até MAX_NETAPP_HOURS
    all_vol_entries.sort(key=lambda x: x["gb"], reverse=True)
    bins = []  # lista de {"volumes": [], "total_gb": 0}

    max_gb_per_bin = MAX_NETAPP_HOURS * avg_gbsec * 3600 if avg_gbsec > 0 else float("inf")

    for entry in all_vol_entries:
        placed = False
        for b in bins:
            if b["total_gb"] + entry["gb"] <= max_gb_per_bin:
                b["volumes"].append(entry["vol"])
                b["total_gb"] += entry["gb"]
                placed = True
                break
        if not placed:
            bins.append({"volumes": [entry["vol"]], "total_gb": entry["gb"]})

    for i, b in enumerate(bins):
        pol_label = f"NETAPP_{chr(65 + i)}"  # NETAPP_A, NETAPP_B, ...
        est_h = b["total_gb"] / (avg_gbsec * 3600) if avg_gbsec > 0 else 0
        vols_str = ", ".join(b["volumes"])

        # Duração atual aproximada (média das políticas originais envolvidas)
        orig_pols = set(next((e["orig_pol"] for e in all_vol_entries if e["vol"] == v), "") for v in b["volumes"])
        current_h = max((policy_summaries[p]["real_h"] or 0 for p in orig_pols if p in policy_summaries), default=0)
        diff = est_h - current_h if current_h > 0 else None
        rec = "✔ Dentro do limite" if est_h <= MAX_NETAPP_HOURS else "⚠ Ainda excede limite — reduzir volumes"

        alt = i % 2 == 1
        for col, val in enumerate([
            pol_label, vols_str, round(b["total_gb"], 1), round(est_h, 2),
            round(current_h, 2) if current_h else "—",
            round(diff, 2) if diff is not None else "—",
            rec, "",
        ], 1):
            cell = ws.cell(row=row, column=col, value=val)
            cell.font = _font(bold=(col == 7),
                              color=(C_GREEN_FT if rec.startswith("✔") else C_RED_FT) if col == 7 else "000000")
            cell.fill = _fill(C_GREEN_BG if rec.startswith("✔") and col == 7 else
                              (C_RED_BG if col == 7 else (C_ALT_ROW if alt else C_WHITE)))
            cell.alignment = _align("left" if col in (1, 2, 7) else "center", wrap=(col == 2))
            cell.border = _border()
        row += 1

    # Nota metodológica
    row += 1
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=8)
    cell = ws.cell(row=row, column=1,
        value=(f"Nota: Duração estimada calculada com base na velocidade média real de "
               f"{avg_kbsec / 1024:.0f} MB/s observada nos jobs NetApp deste mês. "
               f"Meta configurável em MAX_NETAPP_HOURS no início do script."))
    cell.font = _font(italic=True, size=9, color="595959")
    cell.fill = _fill(C_WHITE)
    cell.alignment = _align("left", wrap=True)
    ws.row_dimensions[row].height = 30

    _set_col_widths(ws, [22, 55, 18, 22, 22, 16, 30, 10])

# ---------------------------------------------------------------------------
# Aba 6 — Volumes NetApp (crescimento)
# ---------------------------------------------------------------------------

def _sheet_netapp_volumes(wb, volumes):
    ws = wb.create_sheet("NetApp — Volumes")
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A3"

    dates = sorted(set(v["snapshot_date"] for v in volumes))
    vol_names = sorted(set(v["volume_name"] for v in volumes))

    # Índice: (date, vol_name) -> row
    idx = {(v["snapshot_date"], v["volume_name"]): dict(v) for v in volumes}

    # Headers: Volume | Total | [data1 Used | data2 Used | ...] | Crescimento (GB)
    headers = ["Volume", "Cota (GB)"] + [f"Usado {d}" for d in dates] + ["Cresc. último mês (GB)", "% Cota usada"]
    widths  = [22, 14] + [18] * len(dates) + [24, 18]
    _header_row(ws, 1, headers, widths)

    for i, vol in enumerate(vol_names):
        row = 2 + i
        alt = i % 2 == 1

        total = idx.get((dates[-1], vol), {}).get("total_gb") if dates else None
        used_vals = [idx.get((d, vol), {}).get("used_gb") for d in dates]

        if len(used_vals) >= 2 and used_vals[-1] is not None and used_vals[-2] is not None:
            growth = used_vals[-1] - used_vals[-2]
        else:
            growth = None

        pct_used = (used_vals[-1] / total * 100) if (used_vals and used_vals[-1] and total) else None

        row_data = [vol, round(total, 1) if total else "—"] + \
                   [round(u, 1) if u is not None else "—" for u in used_vals] + \
                   [round(growth, 1) if growth is not None else "—",
                    f"{pct_used:.1f}%" if pct_used else "—"]

        for col, val in enumerate(row_data, 1):
            cell = ws.cell(row=row, column=col, value=val)
            cell.font = _font()
            cell.fill = _fill(C_ALT_ROW if alt else C_WHITE)

            # Colorir crescimento
            if col == len(dates) + 3 and growth is not None:
                if growth > 50:
                    cell.fill = _fill(C_RED_BG)
                    cell.font = _font(color=C_RED_FT)
                elif growth > 10:
                    cell.fill = _fill(C_YELLOW_BG)
                    cell.font = _font(color=C_YELLOW_FT)
                elif growth < 0:
                    cell.fill = _fill(C_GREEN_BG)
                    cell.font = _font(color=C_GREEN_FT)

            # Colorir % utilização
            if col == len(dates) + 4 and pct_used is not None:
                if pct_used >= 90:
                    cell.fill = _fill(C_RED_BG)
                    cell.font = _font(color=C_RED_FT)
                elif pct_used >= 75:
                    cell.fill = _fill(C_YELLOW_BG)
                    cell.font = _font(color=C_YELLOW_FT)

            cell.alignment = _align("center" if col > 1 else "left")
            cell.border = _border()

# ---------------------------------------------------------------------------
# Aba 7 — Fitas
# ---------------------------------------------------------------------------

def _sheet_fitas(wb, tapes):
    ws = wb.create_sheet("Fitas — Tapes")
    ws.sheet_view.showGridLines = False

    _section_title(ws, 1, "UTILIZAÇÃO DE FITAS LTO POR POOL", 6)
    _header_row(ws, 2, ["Pool", "Qtd Fitas", "Volume Total (TB)", "Volume Médio/Fita (TB)", "Data Snapshot", ""],
                widths=[28, 14, 20, 24, 18, 10])

    # Agrupa por (snapshot_date, pool)
    pools = {}
    for t in tapes:
        key = (t["snapshot_date"], t["pool"])
        if key not in pools:
            pools[key] = {"qty": 0, "kb": 0}
        pools[key]["qty"] += 1
        pools[key]["kb"] += (t["kilobytes"] or 0)

    row = 3
    for i, ((date, pool), data) in enumerate(sorted(pools.items())):
        total_tb = data["kb"] / (1024 ** 3)
        avg_tb   = total_tb / data["qty"] if data["qty"] else 0
        alt = i % 2 == 1
        for col, val in enumerate([pool, data["qty"], round(total_tb, 2), round(avg_tb, 3), date, ""], 1):
            cell = ws.cell(row=row, column=col, value=val)
            cell.font = _font()
            cell.fill = _fill(C_ALT_ROW if alt else C_WHITE)
            cell.alignment = _align("center" if col > 1 else "left")
            cell.border = _border()
        row += 1

    # Totais por data
    row += 1
    _section_title(ws, row, "TOTAIS POR DATA DE SNAPSHOT", 6)
    row += 1
    _header_row(ws, row, ["Data", "Total Fitas", "Volume Total (TB)", "", "", ""])
    row += 1

    by_date = {}
    for t in tapes:
        d = t["snapshot_date"]
        by_date.setdefault(d, {"qty": 0, "kb": 0})
        by_date[d]["qty"] += 1
        by_date[d]["kb"] += (t["kilobytes"] or 0)

    for i, (date, data) in enumerate(sorted(by_date.items())):
        total_tb = data["kb"] / (1024 ** 3)
        alt = i % 2 == 1
        for col, val in enumerate([date, data["qty"], round(total_tb, 2), "", "", ""], 1):
            cell = ws.cell(row=row, column=col, value=val)
            cell.font = _font(bold=(col in (1, 2, 3)))
            cell.fill = _fill(C_ALT_ROW if alt else C_WHITE)
            cell.alignment = _align("center" if col > 1 else "left")
            cell.border = _border()
        row += 1

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def generate_report(db_path=DB_PATH, output_dir=OUTPUT_DIR):
    print(f"Lendo banco: {db_path}")
    avail, jobs, pols, tapes, volumes = _load(db_path)

    # Detecta mês de referência a partir dos jobs
    dates_in_jobs = [_parse_dt(j["start_time"]) for j in jobs if j["start_time"]]
    if dates_in_jobs:
        ref_dt = min(dates_in_jobs)
        report_month = ref_dt.strftime("%B/%Y").capitalize()
    else:
        report_month = datetime.now().strftime("%B/%Y").capitalize()

    wb = Workbook()
    wb.remove(wb.active)  # remove aba padrão

    print("Gerando abas...")
    _sheet_capa(wb, avail, jobs, volumes, report_month)
    _sheet_disponibilidade(wb, avail)
    _sheet_backup_geral(wb, jobs)
    _sheet_janelas(wb, jobs)
    _sheet_netapp_split(wb, jobs, pols, volumes)
    _sheet_netapp_volumes(wb, volumes)
    _sheet_fitas(wb, tapes)

    # Salva com timestamp
    output_dir.mkdir(parents=True, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = output_dir / f"relatorio_ti_cptm_{ts}.xlsx"
    wb.save(filename)
    print(f"Relatório gerado: {filename}")
    return filename


if __name__ == "__main__":
    generate_report()