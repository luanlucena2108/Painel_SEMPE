"""
SEMPE — Gerador de dados.json
Lê a planilha SEMPE_Base_PowerBI_2025_preenchida.xlsx e gera docs/dados.json
Roda automaticamente pelo GitHub Actions todo dia às 07h (Brasília)
"""
import json, os
from datetime import datetime, timezone, timedelta
from openpyxl import load_workbook

XLSX  = "SEMPE_Base_PowerBI_2025_preenchida.xlsx"
SAIDA = "docs/dados.json"

def num(v):
    try: return float(v) if v not in (None, "", "-") else None
    except: return None

def txt(v):
    return str(v).strip() if v not in (None, "") else ""

def ler_aba(ws):
    rows = []
    for i, row in enumerate(ws.iter_rows(min_row=7, values_only=True)):
        ind = txt(row[0])
        if not ind or ind.startswith("⚠") or ind.upper().startswith("TOTAL"):
            continue
        rows.append({
            "indicador":     ind,
            "mes_ref":       txt(row[1]),
            "meta_anual":    num(row[2]),
            "realizado":     num(row[3]),
            "municipios":    num(row[5]),
            "operacoes":     num(row[6]),
            "valor_rs":      num(row[7]),
            "beneficiarios": num(row[8]),
        })
    return rows

wb   = load_workbook(XLSX, data_only=True)
dpf  = ler_aba(wb["DPF"])
dap  = ler_aba(wb["DAP"])
de   = ler_aba(wb["DE"])

def soma(rows, campo): return sum(r[campo] or 0 for r in rows)

bpp_op   = soma(dpf, "operacoes")
bpp_val  = soma(dpf, "valor_rs")
bpp_mun  = soma(dpf, "municipios")
bpp_ben  = soma(dpf, "beneficiarios")
bpp_meta = next((r["meta_anual"] for r in dpf if r["meta_anual"]), 26250)

dap_ben  = soma(dap, "beneficiarios")
dap_ev   = soma(dap, "operacoes")
dap_mun  = soma(dap, "municipios")
dap_meta = next((r["meta_anual"] for r in dap if r["meta_anual"]), 34000)

de_ev    = soma(de, "operacoes")
de_mun   = soma(de, "municipios")
de_ben   = soma(de, "beneficiarios")
de_meta  = next((r["meta_anual"] for r in de if r["meta_anual"]), 100)

mes_ref  = next((r["mes_ref"] for r in dpf if r["mes_ref"]), "—")
brt      = datetime.now(timezone(timedelta(hours=-3)))

dados = {
    "meta": {
        "periodo":    f"2023 – {mes_ref}",
        "atualizado": brt.strftime("%d/%m/%Y %H:%M"),
        "gerado_em":  brt.isoformat()
    },
    "geral": {
        "op_credito":              int(bpp_op),
        "montante":                round(bpp_val, 2),
        "municipios_redesim":      642,
        "empresas_facilita":       56000,
        "concluintes_artesanais":  int(dap_ben) if dap_ben else 14274,
        "carteiras_artesao":       5897
    },
    "dpf": {
        "op_2023": 14458, "op_2024": 13533, "op_2025": 11623,
        "op_2026": int(bpp_op),
        "val_2023": 232400000, "val_2024": 208000000, "val_2025": 187800000,
        "val_2026": round(bpp_val, 2),
        "municipios":   int(bpp_mun) if bpp_mun else 547,
        "adimplencia":  0.79,
        "ticket_medio": round(bpp_val / bpp_op) if bpp_op else 16155,
        "meta_op":      bpp_meta,
        "pct_meta":     round(bpp_op / bpp_meta, 4) if bpp_meta else None,
        "indicadores": [
            {"nome": "Nº de operações de crédito", "v2023": "14.458", "v2024": "13.533", "v2025": "11.623", "v2026": f"{int(bpp_op):,}".replace(",","."), "ppa": True,  "prio": "Alta"},
            {"nome": "Taxa de adimplência",        "v2023": "—",      "v2024": "81,8%",  "v2025": "79,0%",  "v2026": "—",                                  "ppa": True,  "prio": "Alta"},
            {"nome": "Montante desembolsado",      "v2023": "R$232,4mi","v2024":"R$208mi","v2025":"R$187,8mi","v2026":f"R${bpp_val/1e6:.1f}mi",            "ppa": False, "prio": "—"},
            {"nome": "Municípios conveniados",     "v2023": "554",    "v2024": "557",    "v2025": "547",    "v2026": str(int(bpp_mun) if bpp_mun else 547),"ppa": False, "prio": "Alta"},
        ],
        "linhas": dpf
    },
    "dap": {
        "ativ_sem_alvara": 927, "municipios_redesim": 642, "municipios_operando": 591,
        "viab_automatica": 478, "empresas_6m": 56000, "horas_abertura": "2h25", "deliberacao": 630,
        "beneficiarios": int(dap_ben), "turmas": int(dap_ev), "municipios": int(dap_mun),
        "meta": dap_meta,
        "pct_meta": round(dap_ben / dap_meta, 4) if dap_meta and dap_ben else None,
        "selos": [
            {"nome": "Bronze",   "qtd": 220, "pct": 100, "cor": "#7B4D2E"},
            {"nome": "Inovação", "qtd": 25,  "pct": 11,  "cor": "#CC0000"},
            {"nome": "Prata",    "qtd": 10,  "pct": 5,   "cor": "#9A9A9A"},
            {"nome": "Ouro",     "qtd": 9,   "pct": 4,   "cor": "#9A6B00"}
        ],
        "linhas": dap
    },
    "de": {
        "concluintes": int(dap_ben) if dap_ben else 14274,
        "carteiras_pab": 4276, "carteiras_emp": 1621,
        "inscritos_feiras": 250, "selecionados": 81, "missoes": 8,
        "eventos": int(de_ev), "municipios": int(de_mun), "participantes": int(de_ben),
        "meta": de_meta,
        "pct_meta": round(de_mun / de_meta, 4) if de_meta and de_mun else None,
        "linhas": de
    },
    "jornada": {
        "etapas": [
            {"pill":"1 · Conhecer",   "cor":"#E5EDFA","txt":"#174FA3","indicador":"Municípios c/ diagnóstico atualizado","meta":"100",    "realizado": str(int(de_mun)) if de_mun else "Em andamento"},
            {"pill":"2 · Simplificar","cor":"#E6F4EC","txt":"#1D6B3A","indicador":"Municípios aderentes ao Facilita SP",  "meta":"645",    "realizado": "645 ✔"},
            {"pill":"2 · Simplificar","cor":"#E6F4EC","txt":"#1D6B3A","indicador":"Atividades dispensadas de alvará",     "meta":"—",      "realizado": "927"},
            {"pill":"3 · Qualificar", "cor":"#FDF0E5","txt":"#C05A00","indicador":"Vagas de qualificação / ano",          "meta":"34.000", "realizado": f"{int(dap_ben):,}".replace(",",".") if dap_ben else "Em andamento"},
            {"pill":"3 · Qualificar", "cor":"#FDF0E5","txt":"#C05A00","indicador":"Ela Empreende (Itaú Social)",          "meta":"5.000",  "realizado": "Aberto"},
            {"pill":"3 · Qualificar", "cor":"#FDF0E5","txt":"#C05A00","indicador":"QualificaSP Empreenda",                "meta":"10.000", "realizado": "Em breve"},
            {"pill":"4 · Capitalizar","cor":"#F9EAEA","txt":"#A30000","indicador":"Operações microcrédito BPP / ano",     "meta":"26.250", "realizado": f"{int(bpp_op):,}".replace(",",".") if bpp_op else "11.623 (2025)"},
            {"pill":"4 · Capitalizar","cor":"#F9EAEA","txt":"#A30000","indicador":"Municípios conveniados BPP",           "meta":"—",      "realizado": str(int(bpp_mun) if bpp_mun else 547)},
            {"pill":"4 · Capitalizar","cor":"#F9EAEA","txt":"#A30000","indicador":"Taxa de adimplência BPP",              "meta":"—",      "realizado": "79%"},
        ],
        "score_limeira": [
            {"dim":"D1 Vitalidade Econômica","val":7.8,"pct":78,"cor":"#174FA3"},
            {"dim":"D2 Facilita SP",         "val":6.5,"pct":65,"cor":"#1D6B3A"},
            {"dim":"D3 Crédito / BPP",       "val":5.2,"pct":52,"cor":"#CC0000"},
            {"dim":"D4 Qualificação",        "val":4.2,"pct":42,"cor":"#C05A00"},
            {"dim":"D5 Artesanato",          "val":4.8,"pct":48,"cor":"#6B3FA0"},
            {"dim":"D6 Ecossistema",         "val":5.5,"pct":55,"cor":"#0B7B6B"}
        ]
    }
}

os.makedirs("docs", exist_ok=True)
with open(SAIDA, "w", encoding="utf-8") as f:
    json.dump(dados, f, ensure_ascii=False, indent=2)

print(f"✓ {SAIDA} gerado — {brt.strftime('%d/%m/%Y %H:%M')} BRT")
