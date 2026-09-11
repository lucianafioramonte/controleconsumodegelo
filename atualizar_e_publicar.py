# -*- coding: utf-8 -*-
"""
Atualiza o painel Torre de Controle - Consumo de Gelo e publica no GitHub.

Faz tudo em um unico script (sem depender de .bat nem do comando "python"
no PATH), pensado para ser compilado como executavel com PyInstaller:

    pyinstaller --onefile --add-data "template.html;." atualizar_e_publicar.py

Isso gera dist\atualizar_e_publicar.exe - coloque esse .exe na pasta
"Controle de Gelo" (junto do repositorio git) e rode ele (manualmente ou
via Task Scheduler) no lugar do .bat.

O QUE ELE FAZ:
  1. Le a planilha RESUMO USINA CONCRETO.xlsx
  2. Recalcula os indicadores de consumo de gelo
  3. Gera o index.html atualizado dentro da pasta "Controle de Gelo"
  4. Da git add / commit / push automaticamente (so se algo mudou)
  5. Registra tudo em log_atualizacao.txt
"""

import openpyxl
import datetime
import json
import collections
import statistics
import sys
import os
import subprocess

# ------------------------------------------------------------------
# CAMINHOS - ajuste aqui se as pastas mudarem
# ------------------------------------------------------------------
SOURCE_XLSX = r"C:\Users\lucianaboas\OneDrive - Construtora Tripoloni\Planejamento\08 - Controles\06. Controle da Usina de Concreto\2026\RESUMO USINA CONCRETO.xlsx"
OUTPUT_DIR = r"C:\Users\lucianaboas\OneDrive - Construtora Tripoloni\Documentos\Controle de Gelo"
OUTPUT_HTML = os.path.join(OUTPUT_DIR, "index.html")

SHEET_NAME = "RESUMO"
FIRST_DATA_ROW = 9
SACO_KG = 20  # peso de cada saco de gelo
PRECO_KG_GELO = 0.70  # R$ por kg de gelo (escama)

GIT_COMMIT_MSG_PREFIX = "Atualizacao automatica do painel"


def resource_path(filename):
    """
    Acha o template.html tanto rodando como script .py quanto rodando
    como .exe compilado pelo PyInstaller (que extrai os dados anexados
    para uma pasta temporaria apontada por sys._MEIPASS).
    """
    if hasattr(sys, "_MEIPASS"):
        base = sys._MEIPASS
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, filename)


def script_dir():
    """Pasta onde o .exe/.py esta rodando (para os comandos git)."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


TEMPLATE_HTML = resource_path("template.html")
LOG_PATH = os.path.join(script_dir(), "log_atualizacao.txt")


def log(msg):
    line = f"{datetime.datetime.now()} - {msg}"
    print(line)
    with open(LOG_PATH, "a", encoding="utf-8") as f:
        f.write(line + "\n")


# ------------------------------------------------------------------
# CLASSIFICACAO
# ------------------------------------------------------------------
def classify_elemento(destino):
    if not destino:
        return "Não informado"
    d = str(destino).upper()
    rules = [
        ("Aduela", ["ADUELA"]),
        ("Viaduto", ["VIADUTO", "VIADULTO"]),
        ("Travessa", ["TRAVESSA", "TRAVASSA"]),
        ("Transversina", ["TRANSVERSINA"]),
        ("Viga", ["VIGA"]),
        ("Laje", ["LAJE", "LAGE"]),
        ("Estaca", ["ESTACA", "ESTAC"]),
        ("Pilar", ["PILAR"]),
        ("Galeria", ["GALERIA", "GALARIA"]),
        ("Dissipador", ["DISSIPADOR"]),
        ("Ala/Cortina", ["ALA", "CORTINA"]),
        ("Guarda-rodas", ["GUARDA"]),
        ("Canaleta", ["CANALETA", "CANELETA"]),
        ("Caixa/Drenagem", ["CAIXA", "DRENAGEM", "BUEIRO"]),
        ("Berço", ["BERÇO", "BERCO"]),
        ("Bloco de Fundação", ["BLOCO"]),
        ("Magro", ["MAGRO"]),
        ("Calçada/Meio-fio", ["CALÇADA", "CALCADA", "MEIO FIO", "MEIO-FIO"]),
        ("Escada", ["ESCADA"]),
        ("Tubulação", ["TUBULA"]),
    ]
    for name, kws in rules:
        for kw in kws:
            if kw in d:
                return name
    return "Outros"


def classify_frente(obs, destino):
    if obs:
        o = str(obs).upper()
        if "TELES PIRES" in o or "PONTE" in o or "TRELI" in o:
            return "Ponte Teles Pires"
        if "PRIMAVERA" in o:
            return "Diamante Primavera"
        if "SORRISO" in o:
            return "Diamante Sorriso"
        if "GALERIA" in o:
            return "Galeria"
        if "CANTEIRO" in o:
            return "Canteiro / Pátio de Pré-moldados"
        if "LAVAGEM" in o or "DESCART" in o:
            return "Outros/Observação"
        return "Outros/Observação"
    return "Sem frente informada"


# ------------------------------------------------------------------
# EXTRACAO
# ------------------------------------------------------------------
def extract_rows(xlsx_path):
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws = wb[SHEET_NAME]
    rows = []
    for i, row in enumerate(
        ws.iter_rows(min_row=FIRST_DATA_ROW, max_row=ws.max_row, values_only=True),
        start=FIRST_DATA_ROW,
    ):
        (date, tick, codigo, hent, hsai, mpa, traco, destino, obs,
         km, vol, tusina, tcampo, gelo_sacos, dmt, volacum) = row
        if date is None or not isinstance(date, datetime.datetime):
            continue
        rows.append({
            "row": i,
            "date": date.strftime("%Y-%m-%d"),
            "hora_saida": hsai.strftime("%H:%M") if hasattr(hsai, "strftime") else None,
            "codigo": codigo,
            "destino": destino,
            "obs": obs,
            "volume": vol if isinstance(vol, (int, float)) else None,
            "temp_usina": tusina if isinstance(tusina, (int, float)) else None,
            "temp_campo": tcampo if isinstance(tcampo, (int, float)) else None,
            "gelo_sacos": gelo_sacos if isinstance(gelo_sacos, (int, float)) else None,
            "elemento": classify_elemento(destino),
            "frente": classify_frente(obs, destino),
        })
    return rows


# ------------------------------------------------------------------
# AGREGACAO
# ------------------------------------------------------------------
def build_aggregates(rows):
    by_date_temp = collections.defaultdict(list)
    for r in rows:
        if r["temp_usina"] is not None:
            by_date_temp[r["date"]].append(r["temp_usina"])
    daily_temp = []
    for d in sorted(by_date_temp.keys()):
        vals = by_date_temp[d]
        daily_temp.append({
            "date": d, "avg": round(statistics.mean(vals), 1),
            "min": min(vals), "max": max(vals), "n": len(vals),
        })

    ice = [r for r in rows if r["gelo_sacos"] is not None]
    for r in ice:
        r["gelo_kg"] = r["gelo_sacos"] * SACO_KG
        r["kg_m3"] = round(r["gelo_kg"] / r["volume"], 1) if r["volume"] else None
        r["custo"] = round(r["gelo_kg"] * PRECO_KG_GELO, 2)
        r["delta_temp"] = (
            round(r["temp_usina"] - r["temp_campo"], 1)
            if (r["temp_usina"] is not None and r["temp_campo"] is not None)
            else None
        )

    by_date_gelo = collections.defaultdict(lambda: {"sacos": 0, "kg": 0, "volume": 0.0, "n": 0})
    for r in ice:
        v = by_date_gelo[r["date"]]
        v["sacos"] += r["gelo_sacos"]
        v["kg"] += r["gelo_kg"]
        v["n"] += 1
        if r["volume"]:
            v["volume"] += r["volume"]
    daily_gelo = []
    for d in sorted(by_date_gelo.keys()):
        v = by_date_gelo[d]
        kg_m3 = round(v["kg"] / v["volume"], 1) if v["volume"] else None
        daily_gelo.append({
            "date": d, "sacos": v["sacos"], "kg": v["kg"],
            "volume": round(v["volume"], 1), "kg_m3": kg_m3, "n": v["n"],
        })

    by_elem = collections.defaultdict(lambda: {"sacos": 0, "kg": 0, "volume": 0.0, "n": 0, "n_vol": 0})
    for r in ice:
        v = by_elem[r["elemento"]]
        v["sacos"] += r["gelo_sacos"]
        v["kg"] += r["gelo_kg"]
        v["n"] += 1
        if r["volume"]:
            v["volume"] += r["volume"]
            v["n_vol"] += 1
    elem_list = []
    for e, v in sorted(by_elem.items(), key=lambda x: -x[1]["kg"]):
        kg_m3_medio = round(v["kg"] / v["volume"], 1) if v["volume"] else None
        vol_medio = round(v["volume"] / v["n_vol"], 1) if v["n_vol"] else None
        elem_list.append({
            "elemento": e, "sacos": v["sacos"], "kg": v["kg"],
            "volume": round(v["volume"], 1), "n": v["n"],
            "kg_m3_medio": kg_m3_medio, "vol_medio_viagem": vol_medio,
        })

    by_frente = collections.defaultdict(lambda: {"sacos": 0, "kg": 0, "n": 0})
    for r in ice:
        v = by_frente[r["frente"]]
        v["sacos"] += r["gelo_sacos"]
        v["kg"] += r["gelo_kg"]
        v["n"] += 1
    frente_list = [
        {"frente": f, "sacos": v["sacos"], "kg": v["kg"], "n": v["n"]}
        for f, v in sorted(by_frente.items(), key=lambda x: -x[1]["kg"])
    ]

    scatter = [
        {"temp": r["temp_usina"], "kg_m3": r["kg_m3"], "date": r["date"], "elemento": r["elemento"]}
        for r in ice if r["temp_usina"] is not None and r["kg_m3"] is not None
    ]

    perda_pairs = [
	{
		"date": r["date"],
		"elemento": r["elemento"],
		"temp_usina": r["temp_usina"],
		"temp_campo": r["temp_campo"],
		"delta": round(r["temp_usina"] - r["temp_campo"], 1)
	}
	for r in rows
	if r["temp_usina"] is not None and r["temp_campo"] is not None
]

    total_sacos = sum(r["gelo_sacos"] for r in ice)
    total_kg = total_sacos * SACO_KG
    total_vol_com_gelo = sum(r["volume"] for r in ice if r["volume"])
    kg_m3_medio = round(total_kg / total_vol_com_gelo, 1) if total_vol_com_gelo else None
    temp_usina_medio_geral = round(
        statistics.mean([r["temp_usina"] for r in rows if r["temp_usina"] is not None]), 1
    ) if any(r["temp_usina"] is not None for r in rows) else None
    temp_usina_medio_periodo_gelo = round(
        statistics.mean([r["temp_usina"] for r in ice if r["temp_usina"] is not None]), 1
    ) if any(r["temp_usina"] is not None for r in ice) else None
    n_temp_campo = sum(1 for r in rows if r["temp_campo"] is not None)
    custo_total = round(total_kg * PRECO_KG_GELO, 2)
    temp_campo_vals_ice = [r["temp_campo"] for r in ice if r["temp_campo"] is not None]
    temp_campo_medio = round(statistics.mean(temp_campo_vals_ice), 1) if temp_campo_vals_ice else None
    n_temp_campo_ice = len(temp_campo_vals_ice)

    kpis = {
        "total_sacos": total_sacos,
        "total_kg": total_kg,
        "total_vol_com_gelo": round(total_vol_com_gelo, 1),
        "kg_m3_medio": kg_m3_medio,
        "temp_usina_medio_geral": temp_usina_medio_geral,
        "temp_usina_medio_periodo_gelo": temp_usina_medio_periodo_gelo,
        "n_temp_campo": n_temp_campo,
        "n_viagens_gelo": len(ice),
        "periodo_temp": [daily_temp[0]["date"], daily_temp[-1]["date"]] if daily_temp else None,
        "periodo_gelo": [daily_gelo[0]["date"], daily_gelo[-1]["date"]] if daily_gelo else None,
        "custo_total": custo_total,
        "preco_kg_gelo": PRECO_KG_GELO,
        "temp_campo_medio": temp_campo_medio,
        "n_temp_campo_ice": n_temp_campo_ice,
    }

    return {
        "daily_temp": daily_temp,
        "daily_gelo": daily_gelo,
        "by_elemento": elem_list,
        "by_frente": frente_list,
        "scatter": scatter,
        "perda_pairs": perda_pairs,
        "kpis": kpis,
        "ice_trips": ice,
    }


# ------------------------------------------------------------------
# GERACAO DO HTML
# ------------------------------------------------------------------
def build_html(aggregates):
    with open(TEMPLATE_HTML, encoding="utf-8") as f:
        template = f.read()
    data_json = json.dumps(aggregates, ensure_ascii=False)
    if "__DATA_JSON__" not in template:
        raise RuntimeError("template.html nao contem o marcador __DATA_JSON__")
    return template.replace("__DATA_JSON__", data_json)


# ------------------------------------------------------------------
# GIT: add / commit / push (so se algo mudou)
# ------------------------------------------------------------------
def git(*args):
    result = subprocess.run(
        ["git"] + list(args),
        cwd=script_dir(),
        capture_output=True,
        text=True,
    )
    output = (result.stdout or "") + (result.stderr or "")
    if output.strip():
        log(f"git {' '.join(args)} -> {output.strip()}")
    return result.returncode, output


def publish_to_git():
	git("add", "index.html")
	msg = f"{GIT_COMMIT_MSG_PREFIX} - {datetime.datetime.now():%d/%m/%Y %H:%M}"
	git("commit", "--allow-empty", "-m", msg)
	rc, out = git("push")
	if rc == 0:
		log("Painel atualizado e publicado no GitHub.")
	else:
		log("ERRO no git push - verifique autenticacao/token. Veja detalhes acima no log.")


# ------------------------------------------------------------------
# MAIN
# ------------------------------------------------------------------
def main():
    log("Iniciando atualizacao")

    if not os.path.exists(SOURCE_XLSX):
        log(f"ERRO: planilha nao encontrada em {SOURCE_XLSX}")
        sys.exit(1)

    try:
        rows = extract_rows(SOURCE_XLSX)
        log(f"{len(rows)} linhas com data lidas da planilha")

        aggregates = build_aggregates(rows)
        log(f"{aggregates['kpis']['n_viagens_gelo']} viagens com registro de gelo")

        html = build_html(aggregates)

        os.makedirs(OUTPUT_DIR, exist_ok=True)
        with open(OUTPUT_HTML, "w", encoding="utf-8") as f:
            f.write(html)
        log(f"Painel gerado em: {OUTPUT_HTML}")

    except Exception as e:
        log(f"ERRO ao gerar o painel: {e}")
        sys.exit(1)

    publish_to_git()
    log("Concluido.\n")


if __name__ == "__main__":
    main()
