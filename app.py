"""
app.py
Interface web Flask para o Planejamento ORT.
Rotas de cache, execução de jobs e servir dashboards.
"""

from flask import Flask, jsonify, request, render_template, send_file
import pandas as pd
import json
from pathlib import Path
import cache_manager as cm
import job_runner as jr
import sys
import os


def resource_path(relative_path: str) -> str:
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath('.'), relative_path)


def get_glpk_path() -> str:
    if hasattr(sys, '_MEIPASS'):
        return resource_path(os.path.join('solvers', 'glpsol.exe'))
    return 'glpsol'


app = Flask(
    __name__,
    template_folder=resource_path('templates'),
)

EXCEL_PATH = resource_path("Simulador - Mix de cartões_3__2_.xlsx")

# ── Ao iniciar: limpar cache se não for persistente ──────────────────────────
cm.limpar_se_nao_persistir()


# ════════════════════════════════════════════════════════════════════════════════
# ROTA PRINCIPAL
# ════════════════════════════════════════════════════════════════════════════════

@app.route("/")
def index():
    return render_template("index.html")


# ════════════════════════════════════════════════════════════════════════════════
# DASHBOARDS (abrem em nova aba)
# ════════════════════════════════════════════════════════════════════════════════

@app.route("/dashboard/mix")
def dashboard_mix():
    return send_file(resource_path("dashboard_mix30dias_ort.html"))


@app.route("/dashboard/sequenciamento")
def dashboard_seq():
    return send_file(resource_path("sequenciamento_v2_ort.html"))


@app.route("/data/dashboard_data_ort.json")
def serve_dashboard_json():
    return send_file(resource_path("dashboard_data_ort.json"), mimetype="application/json")


@app.route("/data/sequenciamento_ort30dias_data.json")
def serve_seq_json():
    return send_file(resource_path("sequenciamento_ort30dias_data.json"), mimetype="application/json")


# ════════════════════════════════════════════════════════════════════════════════
# STATUS DO JOB
# ════════════════════════════════════════════════════════════════════════════════

@app.route("/api/status")
def api_status():
    return jsonify(jr.get_state())


# ════════════════════════════════════════════════════════════════════════════════
# EXECUÇÃO DOS SCRIPTS
# ════════════════════════════════════════════════════════════════════════════════

@app.route("/api/mix/run", methods=["POST"])
def api_run_mix():
    ok, msg = jr.rodar_mix()
    return jsonify({"ok": ok, "mensagem": msg}), (200 if ok else 409)


@app.route("/api/sequenciamento/run", methods=["POST"])
def api_run_seq():
    ok, msg = jr.rodar_sequenciamento()
    return jsonify({"ok": ok, "mensagem": msg}), (200 if ok else 409)


# ════════════════════════════════════════════════════════════════════════════════
# CACHE — SESSÃO (persistência)
# ════════════════════════════════════════════════════════════════════════════════

@app.route("/api/cache/sessao", methods=["GET"])
def get_sessao():
    return jsonify(cm.get_secao("sessao") or {"persistir": False})


@app.route("/api/cache/sessao", methods=["POST"])
def set_sessao():
    cm.set_secao("sessao", request.json)
    return jsonify({"ok": True})


# ════════════════════════════════════════════════════════════════════════════════
# CACHE — CAPACIDADES
# ════════════════════════════════════════════════════════════════════════════════

@app.route("/api/cache/capacidades", methods=["GET"])
def get_capacidades():
    cache = cm.get_secao("capacidades")
    if cache is not None:
        return jsonify(cache)
    try:
        df = pd.read_excel(
            EXCEL_PATH,
            sheet_name="Balanço fábrica ORT",
            skiprows=3,
            usecols="K:P",
            header=0,
        )
        # Colunas confirmadas: Área.1, Unidade, Capacidade MSR, ...
        df = df[["Área.1", "Capacidade MSR"]].dropna(subset=["Área.1"])
        resultado = {
            str(k): float(v)
            for k, v in df.set_index("Área.1")["Capacidade MSR"].items()
            if pd.notna(v)
        }
        return jsonify(resultado)
    except Exception as e:
        import traceback
        print("ERRO get_capacidades:", traceback.format_exc())
        return jsonify({"erro": str(e)}), 500


@app.route("/api/cache/capacidades", methods=["POST"])
def set_capacidades():
    cm.set_secao("capacidades", request.json)
    return jsonify({"ok": True})


@app.route("/api/cache/capacidades/reset", methods=["POST"])
def reset_capacidades():
    cm.reset_secao("capacidades")
    return jsonify({"ok": True})


# ════════════════════════════════════════════════════════════════════════════════
# CACHE — DEMANDA
# ════════════════════════════════════════════════════════════════════════════════

@app.route("/api/cache/demanda", methods=["GET"])
def get_demanda():
    # Distingue cache presente (mesmo vazio) de cache ausente
    cache = cm.get_secao("demanda")
    if cache is not None:
        return jsonify(cache)
    # Cache ausente — pré-popula do Excel
    try:
        df = pd.read_excel(EXCEL_PATH, sheet_name="Demanda", skiprows=2)
        df = df[["Mercado", "Produto", "Quantidade (TO)", "Preço"]].dropna(subset=["Produto"])
        # Converte NaN numéricos para None (JSON-serializável)
        records = df.where(pd.notna(df), None).to_dict(orient="records")
        return jsonify(records)
    except Exception as e:
        import traceback
        print("ERRO get_demanda:", traceback.format_exc())
        return jsonify({"erro": str(e)}), 500


@app.route("/api/cache/demanda", methods=["POST"])
def set_demanda():
    cm.set_secao("demanda", request.json)
    return jsonify({"ok": True})


@app.route("/api/cache/demanda/reset", methods=["POST"])
def reset_demanda():
    cm.reset_secao("demanda")
    return jsonify({"ok": True})


# ════════════════════════════════════════════════════════════════════════════════
# CACHE — CUSTOS
# ════════════════════════════════════════════════════════════════════════════════

@app.route("/api/cache/custos", methods=["GET"])
def get_custos():
    cache = cm.get_secao("custos")
    if cache is not None:
        return jsonify(cache)
    try:
        df = pd.read_excel(EXCEL_PATH, sheet_name="Custos", skiprows=2)
        # Tenta os dois nomes possíveis da coluna (com e sem espaço)
        col_custo = "Custo Variavel " if "Custo Variavel " in df.columns else "Custo Variavel"
        df = df[["Produto", "Máquina", col_custo]].dropna(subset=["Produto"])
        df = df.rename(columns={col_custo: "Custo Variavel"})
        df = df[df["Máquina"].isin([27, 28, 25, 26])]
        records = df.where(pd.notna(df), None).to_dict(orient="records")
        return jsonify(records)
    except Exception as e:
        import traceback
        print("ERRO get_custos:", traceback.format_exc())
        return jsonify({"erro": str(e)}), 500


@app.route("/api/cache/custos", methods=["POST"])
def set_custos():
    cm.set_secao("custos", request.json)
    return jsonify({"ok": True})


@app.route("/api/cache/custos/reset", methods=["POST"])
def reset_custos():
    cm.reset_secao("custos")
    return jsonify({"ok": True})


# ════════════════════════════════════════════════════════════════════════════════
# CACHE — PARÂMETROS GLOBAIS (câmbio, preços celulose)
# ════════════════════════════════════════════════════════════════════════════════

PARAMS_GLOBAIS_KEYS = [
    "Custo Variavel Celulose FC",
    "Custo Variavel Celulose FL",
    "Preço de Cel. Merc. ME FC",
    "Preço de Cel. Merc. ME FF",
    "Cambio",
]


@app.route("/api/cache/parametros", methods=["GET"])
def get_parametros():
    cache = cm.get_secao("parametros")
    if cache is not None:
        return jsonify(cache)
    try:
        df = pd.read_excel(EXCEL_PATH, sheet_name="Parâmetros", skiprows=2)
        df = df[["Parâmetro", "Valor"]].dropna(subset=["Parâmetro"])
        df = df[df["Parâmetro"].isin(PARAMS_GLOBAIS_KEYS)]
        resultado = {
            str(k): float(v) if pd.notna(v) else None
            for k, v in df.set_index("Parâmetro")["Valor"].to_dict().items()
        }
        return jsonify(resultado)
    except Exception as e:
        import traceback
        print("ERRO get_parametros:", traceback.format_exc())
        return jsonify({"erro": str(e)}), 500


@app.route("/api/cache/parametros", methods=["POST"])
def set_parametros():
    cm.set_secao("parametros", request.json)
    return jsonify({"ok": True})


@app.route("/api/cache/parametros/reset", methods=["POST"])
def reset_parametros():
    cm.reset_secao("parametros")
    return jsonify({"ok": True})


# ════════════════════════════════════════════════════════════════════════════════
# CACHE — PARÂMETROS DO MIX
# ════════════════════════════════════════════════════════════════════════════════

MIX_PARAMS_DEFAULT = {
    "DIAS_PERIODO": 7,
    "BASE_DEMANDA_DIAS": 365,
}


@app.route("/api/cache/mix_params", methods=["GET"])
def get_mix_params():
    cache = cm.get_secao("mix_params")
    return jsonify(cache or MIX_PARAMS_DEFAULT)


@app.route("/api/cache/mix_params", methods=["POST"])
def set_mix_params():
    cm.set_secao("mix_params", request.json)
    return jsonify({"ok": True})


@app.route("/api/cache/mix_params/reset", methods=["POST"])
def reset_mix_params():
    cm.reset_secao("mix_params")
    return jsonify({"ok": True})


# ════════════════════════════════════════════════════════════════════════════════
# CACHE — PARÂMETROS DO SEQUENCIAMENTO
# ════════════════════════════════════════════════════════════════════════════════

SEQ_PARAMS_DEFAULT = {
    "DIAS_PERIODO": 7,
    "LOTE_T": 700.0,
    "PERDA_SETUP_T_DEFAULT": 20.0,
    "FATOR_TOLERANCIA_RITMO": 0.5,
    "FATOR_SPREAD": 1.0,
    "FRACAO_MIN_LOTE_PARCIAL": 0.50,
    "DISPERSAO_UNIFORME_ORT": True,
    "FATOR_DISPERSAO_ORT": 0.5,
}


@app.route("/api/cache/seq_params", methods=["GET"])
def get_seq_params():
    cache = cm.get_secao("seq_params")
    return jsonify(cache or SEQ_PARAMS_DEFAULT)


@app.route("/api/cache/seq_params", methods=["POST"])
def set_seq_params():
    dados = request.json
    # DIAS_PERIODO do sequenciamento deve ser igual ao do mix — sincroniza
    mix_params = cm.get_secao("mix_params") or MIX_PARAMS_DEFAULT
    dados["DIAS_PERIODO"] = mix_params.get("DIAS_PERIODO", SEQ_PARAMS_DEFAULT["DIAS_PERIODO"])
    cm.set_secao("seq_params", dados)
    return jsonify({"ok": True})


@app.route("/api/cache/seq_params/reset", methods=["POST"])
def reset_seq_params():
    cm.reset_secao("seq_params")
    return jsonify({"ok": True})


# ════════════════════════════════════════════════════════════════════════════════
# RESET GLOBAL
# ════════════════════════════════════════════════════════════════════════════════

@app.route("/api/cache/reset_tudo", methods=["POST"])
def reset_tudo():
    cm.reset_tudo()
    return jsonify({"ok": True})


# ════════════════════════════════════════════════════════════════════════════════
# ENTRY POINT
# ════════════════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    app.run(host='127.0.0.1', port=8766, debug=False, use_reloader=False)
