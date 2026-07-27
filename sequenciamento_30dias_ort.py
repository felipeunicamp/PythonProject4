import json
import math
import unicodedata
import pyomo.environ as pyo
from pyomo.opt import SolverFactory
import pandas as pd
from collections import defaultdict
import sys
import os


def get_glpk_path() -> str:
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, 'solvers', 'glpsol.exe')
    return 'glpsol'

# ══════════════════════════════════════════════════════════════════════
# BLOCO DE CACHE — aplicado quando chamado via app.py (Flask)
# ══════════════════════════════════════════════════════════════════════
import json as _json
from pathlib import Path as _Path

_CACHE_PATH = _Path("user_cache.json")
_cache = {}
if _CACHE_PATH.exists():
    try:
        with open(_CACHE_PATH, encoding="utf-8") as _f:
            _cache = _json.load(_f)
    except Exception:
        _cache = {}

def _cache_get(secao):
    return _cache.get(secao)
# ══════════════════════════════════════════════════════════════════════

# ═══════════════════════════════════════════════════════
# SEÇÃO 1 — PARÂMETROS CONFIGURÁVEIS
# ═══════════════════════════════════════════════════════

ARQUIVO_EXCEL   = 'Simulador - Mix de cartões_3__2_.xlsx'
ARQUIVO_JSON    = 'dashboard_data_ort.json'

_seq_params = _cache_get("seq_params") or {}
DIAS_PERIODO            = _seq_params.get("DIAS_PERIODO",            7)
LOTE_T                  = _seq_params.get("LOTE_T",                  700.0)
PERDA_SETUP_T_DEFAULT   = _seq_params.get("PERDA_SETUP_T_DEFAULT",   20.0)
PERDA_SETUP_T = {
    27: PERDA_SETUP_T_DEFAULT,
    28: PERDA_SETUP_T_DEFAULT,
    25: PERDA_SETUP_T_DEFAULT,
    26: PERDA_SETUP_T_DEFAULT,
}
JANELA_MAX_DIAS         = DIAS_PERIODO  # um ciclo por horizonte quando DIAS_PERIODO <= 60
FATOR_TOLERANCIA_RITMO  = _seq_params.get("FATOR_TOLERANCIA_RITMO",  0.5)
FATOR_SPREAD            = _seq_params.get("FATOR_SPREAD",            1)
FRACAO_MIN_LOTE_PARCIAL = _seq_params.get("FRACAO_MIN_LOTE_PARCIAL", 0.50)
DISPERSAO_UNIFORME_ORT  = _seq_params.get("DISPERSAO_UNIFORME_ORT",  True)
FATOR_DISPERSAO_ORT     = _seq_params.get("FATOR_DISPERSAO_ORT",     0.5)

MAQUINAS_ORT = [27, 28, 25, 26]
MAQUINA_CENTRO = {27: 'ORT', 28: 'ORT', 25: 'ORT', 26: 'ORT'}

# Máquinas ORT com matrizes de setup dependentes de sequência
MAQUINAS_ORT_COM_SETUP_MATRIZ = [27, 28]
# Máquinas ORT com setup fixo (PERDA_SETUP_T)
MAQUINAS_ORT_SETUP_FIXO = [25, 26]

RESTRICOES_HORARIAS = {'Evaporação ORT', 'Outorga Captação', 'Outorga Emissario'}

# ── Setup dependente ORT ────────────────────────────────
ARQUIVO_SETUP_MP27          = 'setup_mp27.xlsx'
ARQUIVO_SETUP_MP28          = 'setup_mp28.xlsx'
SETUP_ENTRE_FAMILIAS_T      = 20.0   # fallback para transições entre famílias não cadastradas
LIMIAR_PENALIDADE_SWAP_T    = 50.0   # delta máximo de custo de setup aceito num swap ORT (t)

# ═══════════════════════════════════════════════════════
# SEÇÃO 2 — LEITURA DE DADOS DO EXCEL (apenas ORT)
# ═══════════════════════════════════════════════════════

# 2A. Produtividade (apenas máquinas ORT)
data1 = pd.read_excel(ARQUIVO_EXCEL, sheet_name='Produtividade', skiprows=2)
produtividade_df = data1[['Produto', 'Máquina',
                           'Produtividade máxima (t/h)', 'Taxa PERF (Meta)']].copy()
taxa_qual_por_maquina = {27: 1.0, 28: 1.0, 25: 1.0, 26: 1.0}

dict_produtividade_bruta = {}
for _, row in produtividade_df.dropna(subset=['Produto', 'Máquina']).iterrows():
    prod = row['Produto']
    maq  = int(row['Máquina'])
    if maq not in MAQUINAS_ORT:
        continue
    pmax = float(row['Produtividade máxima (t/h)'])
    tperf = float(row['Taxa PERF (Meta)'])
    tq   = taxa_qual_por_maquina.get(maq, 1.0)
    dict_produtividade_bruta[(maq, prod)] = tq * pmax * tperf

# 2B. Máquinas (apenas ORT)
maquinas_raw = pd.read_excel(ARQUIVO_EXCEL, sheet_name='Máquinas', skiprows=2)
maquinas_raw = maquinas_raw[['Máquina', 'Tempo de carga (h)', 'Taxa DISP']].dropna(
    subset=['Máquina'])
dict_tempo_carga = {}
dict_taxa_DISP   = {}
for _, row in maquinas_raw.iterrows():
    m = int(row['Máquina'])
    if m not in MAQUINAS_ORT:
        continue
    dict_tempo_carga[m] = float(row['Tempo de carga (h)'])
    dict_taxa_DISP[m]   = float(row['Taxa DISP'])

# horas disponíveis POR DIA por máquina (base 365 dias)
horas_dia = {
    m: (dict_tempo_carga[m] * dict_taxa_DISP[m]) / 365
    for m in dict_tempo_carga
}

# 2C. Produto por MP (apenas colunas ORT)
data6 = pd.read_excel(ARQUIVO_EXCEL, sheet_name='Produto por MP', skiprows=2)
produto_por_MP = data6[['Produto', 'MP27', 'MP28', 'MC25', 'MC26']].copy()
cols_maq = ['MP27', 'MP28', 'MC25', 'MC26']
longa = produto_por_MP.melt(id_vars=['Produto'], value_vars=cols_maq,
                             var_name='maq_str', value_name='flag')
longa['Máquina'] = longa['maq_str'].str.replace('MP', '').str.replace('MC', '').astype(int)
longa['flag'] = longa['flag'].fillna(0).astype(int)
dict_prod_por_MP = longa.set_index(['Máquina', 'Produto'])['flag'].to_dict()

# 2D. Lista de Materiais
data8 = pd.read_excel(ARQUIVO_EXCEL, sheet_name='Lista de Materiais', skiprows=2)
lista_mat = data8[['Centro', 'Índice', 'Código', 'UMB', 'Material', 'Valor']].copy()
lista_mat = lista_mat[
    (~lista_mat['Código'].astype(str).str.isnumeric()) &
    lista_mat['Índice'].isin(['Esp', 'Específico', 'Qtd']) &
    ~lista_mat['Material'].isin([
        'AGUA-TRATADA', 'ACIDO-SULF', 'VAPOR-MEDIO',
        'ENERG-MEDIA', 'ENERG-TERM', 'VAPOR-LICOR', 'AGUA-DESMI',
        'ENERG-PROP', 'AGUA-CAP1', 'ETE',
        'VAPOR-BIOMASSA', 'VAPOR-OLEO', 'BIOMASSA-F', 'CAV-ENERGIA',
        'DISPER-MP12', 'DISPER-MP13', 'DISPER-MP16', 'DISPER-MP23'
    ])
].copy()
for idx, row in lista_mat.iterrows():
    if row['UMB'] == 'KG':
        lista_mat.at[idx, 'Valor'] = row['Valor'] / 1000


def explodir_lista_materiais(df):
    def _recursive_explode(material, bom_map, factor=1.0, _visited=frozenset()):
        totals = defaultdict(float)
        for sub_codigo, qty_direct in bom_map.get(material, []):
            total_qty = qty_direct * factor
            totals[sub_codigo] += total_qty
            if sub_codigo in bom_map and sub_codigo not in _visited:
                for k, v in _recursive_explode(
                        sub_codigo, bom_map, total_qty, _visited | {material}).items():
                    totals[k] += v
        return dict(totals)

    result = {}
    for centro in df['Centro'].unique():
        df_c = df[df['Centro'] == centro]
        bom_map = defaultdict(list)
        for _, row in df_c.iterrows():
            bom_map[row['Material']].append((row['Código'], row['Valor']))
        for produto_raiz in set(df_c['Material']):
            r = _recursive_explode(produto_raiz, bom_map, 1.0)
            if r:
                result[(centro, produto_raiz)] = r
    return result


dict_consumo_especifico = explodir_lista_materiais(lista_mat)


def get_consumo_esp(produto, maq, material):
    centro   = MAQUINA_CENTRO.get(maq, 'ORT')
    prod_adj = produto[:3] + str(maq).zfill(2) + produto[3:]
    return dict_consumo_especifico.get((centro, prod_adj), {}).get(material, 0.0)


# 2E. Balanço fábrica ORT — Produção celulose e consumo fibras
data13 = pd.read_excel(ARQUIVO_EXCEL, sheet_name='Balanço fábrica ORT',
                       skiprows=3, usecols='B:E')
Prod_Cel_Fibras_ORT = pd.DataFrame(data13).dropna()
Prod_Cel_Fibras_ORT['Consumo_Anual_Fibras'] = (
    Prod_Cel_Fibras_ORT['Consumo (t/dia)'] * Prod_Cel_Fibras_ORT['Dias operação'])
Prod_Cel_Fibras_ORT = Prod_Cel_Fibras_ORT.groupby('Fibra')['Consumo_Anual_Fibras'].sum().reset_index()
dict_consumo_fibras_ORT = Prod_Cel_Fibras_ORT.set_index('Fibra')['Consumo_Anual_Fibras'].to_dict()
lista_fibras_ORT = list(Prod_Cel_Fibras_ORT['Fibra'].unique())

# Adicionar fibras CDR que podem não estar na tabela B:E
fibras_cdr_ORT = ['CKN-FC', 'CKN-FL', 'BCTMP']
for _f in fibras_cdr_ORT:
    if _f not in lista_fibras_ORT:
        lista_fibras_ORT.append(_f)

# 2F. Balanço fábrica ORT — Parâmetros adicionais
data14 = pd.read_excel(ARQUIVO_EXCEL, sheet_name='Balanço fábrica ORT',
                       skiprows=3, usecols='G:I')
dict_param_add_ORT = pd.DataFrame(data14).dropna().set_index('Parâmetro')['Valor'].to_dict()

# 2G. Balanço fábrica ORT — Capacidade das plantas
data15 = pd.read_excel(ARQUIVO_EXCEL, sheet_name='Balanço fábrica ORT',
                       skiprows=3, usecols='K:P')
capac_ort = data15.dropna(subset=['Área.1'])
dict_emissario = capac_ort.set_index('Área.1')['Capacidade MSR'].to_dict()
dict_dias_operacao_ORT = capac_ort.set_index('Área.1')['Dias operação.1'].to_dict()
# Aplica overrides de capacidade do cache
_cap_cache = _cache_get("capacidades")
if _cap_cache:
    dict_emissario.update(_cap_cache)

# 2H. Balanço fábrica ORT — Fibras e digestores
data16 = pd.read_excel(ARQUIVO_EXCEL, sheet_name='Balanço fábrica ORT',
                       skiprows=3, usecols='R:U')
Fibras_Digestores_ORT = pd.DataFrame(data16).dropna()
dict_rendimento_ort     = Fibras_Digestores_ORT.set_index('Fibra.1')['Rendimento (%)'].to_dict()
dict_carga_alcalina_ort = Fibras_Digestores_ORT.set_index('Fibra.1')['Carga alcalina (%)'].to_dict()

# ═══════════════════════════════════════════════════════
# SEÇÃO 2I — MATRIZES DE SETUP ORT (apenas MP27 e MP28)
# ═══════════════════════════════════════════════════════

def _carregar_matriz_aba(arquivo, sheet_name):
    """Lê uma aba de matriz de setup e retorna dict {(p_orig, p_dest): toneladas}."""
    df = pd.read_excel(arquivo, sheet_name=sheet_name, index_col=0)
    resultado = {}
    for p_orig in df.index:
        for p_dest in df.columns:
            val = df.loc[p_orig, p_dest]
            if pd.notna(val):
                resultado[(str(p_orig).strip(), str(p_dest).strip())] = float(val)
    return resultado


# MP27 — três famílias em abas separadas
_mp27_plan1   = _carregar_matriz_aba(ARQUIVO_SETUP_MP27, 'Plan1')
_mp27_plan2   = _carregar_matriz_aba(ARQUIVO_SETUP_MP27, 'Planilha1')
_mp27_plan3   = _carregar_matriz_aba(ARQUIVO_SETUP_MP27, 'Planilha3')

# MP28 — duas famílias em abas separadas
_mp28_branco  = _carregar_matriz_aba(ARQUIVO_SETUP_MP28, 'Set up Branco')
_mp28_kraft   = _carregar_matriz_aba(ARQUIVO_SETUP_MP28, 'Set up Kraft')

# Conjunto de produtos Kraft do MP28
_df_kraft_cols = pd.read_excel(ARQUIVO_SETUP_MP28, sheet_name='Set up Kraft', index_col=0)
PRODUTOS_KRAFT_MP28 = set(str(c).strip() for c in _df_kraft_cols.columns)


def get_setup_t(maquina, p_anterior, p_atual):
    """
    Retorna a perda de setup em toneladas para a transição p_anterior → p_atual.

    MC25 e MC26: sempre retorna PERDA_SETUP_T (setup fixo).
    MP27 e MP28: usa matrizes dependentes de sequência com fallback SETUP_ENTRE_FAMILIAS_T.
    Qualquer outra máquina: retorna PERDA_SETUP_T.
    """
    if p_anterior is None:
        return 0.0
    p_ant = str(p_anterior).strip()
    p_atu = str(p_atual).strip()
    if p_ant == p_atu:
        return 0.0

    # MC25 e MC26: setup fixo
    if maquina in MAQUINAS_ORT_SETUP_FIXO:
       return PERDA_SETUP_T.get(maquina, PERDA_SETUP_T_DEFAULT) if isinstance(PERDA_SETUP_T,
                                                                                   dict) else PERDA_SETUP_T

    chave = (p_ant, p_atu)

    if maquina == 27:
        for matriz in (_mp27_plan1, _mp27_plan2, _mp27_plan3):
            if chave in matriz:
                return matriz[chave]
        return SETUP_ENTRE_FAMILIAS_T

    if maquina == 28:
        if p_atu in PRODUTOS_KRAFT_MP28:
            if chave in _mp28_kraft:
                return _mp28_kraft[chave]
        else:
            if chave in _mp28_branco:
                return _mp28_branco[chave]
        return SETUP_ENTRE_FAMILIAS_T

    # Fallback genérico
    return PERDA_SETUP_T.get(maquina, PERDA_SETUP_T_DEFAULT) if isinstance(PERDA_SETUP_T, dict) else PERDA_SETUP_T

def custo_sequencia_setup(maquina, sequencia_produtos):
    """
    Calcula o custo total de setup (em toneladas) para uma sequência ordenada
    de produtos numa máquina ORT.
    """
    if len(sequencia_produtos) < 2:
        return 0.0
    return sum(
        get_setup_t(maquina, sequencia_produtos[i], sequencia_produtos[i + 1])
        for i in range(len(sequencia_produtos) - 1)
    )


# ═══════════════════════════════════════════════════════
# SEÇÃO 3 — LEITURA DE METAS DO dashboard_data_ort.json
# ═══════════════════════════════════════════════════════

with open(ARQUIVO_JSON, encoding='utf-8') as f:
    json_data = json.load(f)

atendida = json_data.get('atendida_por_maquina_produto', {})
# meta_anual[p][m] = toneladas do período otimizadas pelo mix_ort_30dias.py
# Como o mix trabalha com DIAS_PERIODO dias, escalamos para base anual
_dias_periodo_json = json_data.get('dias_periodo', DIAS_PERIODO)
_fator_anual = 1.0  # metas já estão na escala do período — sem escalonamento

meta_anual = {}
for prod, maq_dict in atendida.items():
    meta_anual[prod] = {}
    for maq_str, val_periodo in maq_dict.items():
        maq = int(maq_str)
        if maq not in MAQUINAS_ORT:
            continue
        if val_periodo > 0:
            # Usar diretamente as toneladas do período (DIAS_PERIODO dias)
            meta_anual[prod][maq] = float(val_periodo)

balanco_ORT_dict = {b['nome']: b for b in json_data.get('balanco_ORT', [])}

# ═══════════════════════════════════════════════════════
# SEÇÃO 4 — LIMITES DIÁRIOS DAS RESTRIÇÕES ORT
# ═══════════════════════════════════════════════════════

limites_ORT = {
    'Caustificação':     dict_emissario.get('Caustificação',     1e9),
    'Outorga Captação':  dict_emissario.get('Outorga Captação',  1e9),
    'Outorga Emissario': dict_emissario.get('Outorga Emissario', 1e9),
    'CKB-FC':            dict_emissario.get('CKB-FC',            1e9),
    'CKB-FL':            dict_emissario.get('CKB-FL',            1e9),
    'CDR':               dict_emissario.get('CDR',               1e9),
    'Evaporação ORT':    dict_emissario.get('Evaporação',        1e9),
}

# ═══════════════════════════════════════════════════════
# SEÇÃO 5 — DIAGNÓSTICO DE UNIDADES (print uma vez no início)
# ═══════════════════════════════════════════════════════

print("\n" + "="*70)
print("DIAGNÓSTICO — UNIDADES DAS RESTRIÇÕES DE BALANÇO ORT")
print(f"  Horizonte: {DIAS_PERIODO} dias | Arquivo JSON: {ARQUIVO_JSON}")
print("="*70)
print("\ndict_emissario (ORT):")
for k, v in dict_emissario.items():
    unid = "/h" if k in RESTRICOES_HORARIAS else "/dia (ou t/dia)"
    print(f"  {k:<20}: {v:>12.4f}  [{unid}]")
print(f"\n  Metas anuais carregadas: {sum(len(v) for v in meta_anual.values())} pares produto-máquina")
print(f"  Fator de escala período→anual: {_fator_anual:.4f} ({_dias_periodo_json} dias → 365 dias)")

# ═══════════════════════════════════════════════════════
# SEÇÃO 6 — COEFICIENTES EXATOS POR LOTE (substitui coef_avg/montar_coefs)
# ═══════════════════════════════════════════════════════

def _consumo_fibra_lote(produto, maquina, fibra, volume_t, unidade='ORT'):
    """
    Consumo da fibra para produzir volume_t do produto na maquina (ORT).
    Usa centro='ORT' para o lookup do BOM.
    """
    prod_adj = produto[:3] + str(maquina).zfill(2) + produto[3:]
    c = dict_consumo_especifico.get(('ORT', prod_adj), {}).get(fibra, 0.0)
    return volume_t * c


def consumo_lote_restricao(produto, maquina, restricao, unidade, vol=None):
    """
    Retorna o consumo da restricao para um lote de `vol` toneladas do produto
    na máquina ORT.

    Unidades de retorno (iguais aos limites em dict_emissario):
      - Outorga Captação, Outorga Emissario: m³/h  (taxa horária instantânea)
      - Evaporação ORT: t_H2O/h  (taxa horária)
      - Caustificação: m³ LB/dia
      - CKB-FC, CKB-FL: t/dia
      - CDR: tss/dia
    """
    if vol is None:
        vol = LOTE_T

    setup_t_maq = PERDA_SETUP_T.get(maquina, PERDA_SETUP_T_DEFAULT) if isinstance(PERDA_SETUP_T, dict) else PERDA_SETUP_T
    produtividade_pm = dict_produtividade_bruta.get((maquina, produto), 0.0)
    horas_lote = (vol + setup_t_maq) / produtividade_pm if produtividade_pm > 0 else 0.0

    if restricao == 'Caustificação':
        licor_branco = 0.0
        dias = dict_dias_operacao_ORT.get('Caustificação', 365.0)
        for fibra in lista_fibras_ORT:
            rendimento     = dict_rendimento_ort.get(fibra, 1.0)
            carga_alcalina = dict_carga_alcalina_ort.get(fibra, 0.0)
            consumo_fibra_anual = _consumo_fibra_lote(produto, maquina, fibra, vol)
            if 'CKB' in fibra:
                concentracao = dict_param_add_ORT.get('Concentração caust 1 (g/l NaOH)', 1.0) / 1000.0
                perda_fb = dict_param_add_ORT.get('Perda de fibra branca (%)', 0.0)
                consumo_dia = (consumo_fibra_anual / (1 - perda_fb) / dias
                               if (1 - perda_fb) > 0 else consumo_fibra_anual / dias)
            else:
                concentracao = dict_param_add_ORT.get('Concentração caust 2 (g/l NaOH)', 1.0) / 1000.0
                consumo_dia = consumo_fibra_anual / dias
            if rendimento > 0 and concentracao > 0:
                licor_branco += 0.9 * consumo_dia * carga_alcalina / rendimento / concentracao
        dias_deslig = dict_dias_operacao_ORT.get('Caustificação', 365.0)
        deslig_euca  = (_consumo_fibra_lote(produto, maquina, 'CKB-FC', vol)
                        * dict_param_add_ORT.get('Deslignificação Euca (m3/tsa)', 0.0)
                        / dias_deslig)
        deslig_pinus = (_consumo_fibra_lote(produto, maquina, 'CKB-FL', vol)
                        * dict_param_add_ORT.get('Deslignificação Pinus (m3/tsa)', 0.0)
                        / dias_deslig)
        return licor_branco + deslig_euca + deslig_pinus

    elif restricao == 'Outorga Captação':
        if horas_lote <= 0:
            return 0.0
        captacao = dict_param_add_ORT.get('Captação de Água (m3/t)', 0.0)
        return captacao * vol / horas_lote

    elif restricao == 'Outorga Emissario':
        if horas_lote <= 0:
            return 0.0
        emissario = dict_param_add_ORT.get('Emissário (m3/t)', 0.0)
        return emissario * vol / horas_lote

    elif restricao in ('CKB-FC', 'CKB-FL'):
        dias = dict_dias_operacao_ORT.get(restricao, 365.0)
        consumo_anual = _consumo_fibra_lote(produto, maquina, restricao, vol)
        return consumo_anual / dias

    elif restricao == 'CDR':
        # TSS gerado pelas fibras do lote (coef SOLIDO-SECO negativo = gerado)
        fibras_cdr = ['CKB-FC', 'CKB-FL', 'CKN-FC', 'CKN-FL', 'BCTMP']
        dias = dict_dias_operacao_ORT.get('CDR', 365.0)
        total_tss = 0.0
        for fibra in fibras_cdr:
            coef_ss = dict_consumo_especifico.get(('ORT', fibra), {}).get('SOLIDO-SECO', 0.0)
            consumo_fibra = _consumo_fibra_lote(produto, maquina, fibra, vol)
            total_tss += -1 * coef_ss * consumo_fibra
        return total_tss / dias if dias > 0 else 0.0

    elif restricao == 'Evaporação ORT':
        # Taxa de evaporação em t_H2O/h para o lote
        if horas_lote <= 0:
            return 0.0
        mapa_conc_digestor = {
            'CKB-FC': 'Concentração Licor Preto gerado no digestor 1 (%)',
            'CKB-FL': 'Concentração Licor Preto gerado no digestor 2 (%)',
            'CKN-FC': 'Concentração Licor Preto gerado no digestor 3 (%)',
            'CKN-FL': 'Concentração Licor Preto gerado no digestor 4 (%)',
            'BCTMP':  'Concentração Licor Preto gerado na BCTMP (%)',
        }
        conc_saida_evap = dict_param_add_ORT.get('Concentração Licor Preto na Saída Evap1', 1.0)
        volume_lp = 0.0
        tss_total  = 0.0
        for fibra, chave_conc in mapa_conc_digestor.items():
            coef_ss = dict_consumo_especifico.get(('ORT', fibra), {}).get('SOLIDO-SECO', 0.0)
            consumo_fibra = _consumo_fibra_lote(produto, maquina, fibra, vol)
            tss_fibra = -1 * coef_ss * consumo_fibra
            conc_digestor = dict_param_add_ORT.get(chave_conc, 1.0)
            if conc_digestor > 0:
                volume_lp += tss_fibra / conc_digestor
            tss_total  += tss_fibra
        evap_total = (volume_lp - tss_total / conc_saida_evap) if conc_saida_evap > 0 else 0.0
        # Converter de base anual para t_H2O/h do lote
        return evap_total / horas_lote if horas_lote > 0 else 0.0

    return 0.0


def _compute_fixed_ort_per_day():
    """
    Retorna a contribuição FIXA (independente do mix de papel) de cada
    restrição ORT, em unidade/dia ou unidade/h conforme o limite base.
    Inclui CDR e Evaporação ORT provenientes da produção fixa de celulose.

    ATENÇÃO — consistência de unidades no horizonte de 30 dias:
    dict_consumo_fibras_ORT já foi escalado para DIAS_PERIODO dias
    (= consumo_anual * DIAS_PERIODO/365) pelo mix_ort_30dias.py.
    Por isso todos os denominadores que no arquivo anual usam
    dict_dias_operacao_ORT (365 dias) aqui usam DIAS_PERIODO,
    garantindo que numerador e denominador estejam na mesma base.
    A exceção é tss_angatuba_dia, que já é uma taxa diária lida
    diretamente dos parâmetros da planilha.
    """
    # prod_cel está em base DIAS_PERIODO (já escalado pelo mix)
    prod_cel = sum(dict_consumo_fibras_ORT.get(f, 0) for f in lista_fibras_ORT)

    captacao_rate  = dict_param_add_ORT.get('Captação de Água (m3/t)', 0.0)
    emissario_rate = dict_param_add_ORT.get('Emissário (m3/t)',        0.0)

    # prod_cel / DIAS_PERIODO = t/dia; / 24 = t/h → m³/h com o rate
    fixed_capt = captacao_rate  * prod_cel / 24.0 / DIAS_PERIODO
    fixed_emis = emissario_rate * prod_cel / 24.0 / DIAS_PERIODO

    # Caustificação fixa — denominador = DIAS_PERIODO (mesma base do numerador)
    licor_branco_fixo = 0.0
    for fibra in lista_fibras_ORT:
        rendimento     = dict_rendimento_ort.get(fibra, 1.0)
        carga_alcalina = dict_carga_alcalina_ort.get(fibra, 0.0)
        consumo_cel_periodo = dict_consumo_fibras_ORT.get(fibra, 0.0)  # base DIAS_PERIODO
        if 'CKB' in fibra:
            concentracao = dict_param_add_ORT.get('Concentração caust 1 (g/l NaOH)', 1.0) / 1000.0
            perda_fb = dict_param_add_ORT.get('Perda de fibra branca (%)', 0.0)
            consumo_dia = (consumo_cel_periodo / (1 - perda_fb) / DIAS_PERIODO
                           if (1 - perda_fb) > 0 else consumo_cel_periodo / DIAS_PERIODO)
        else:
            concentracao = dict_param_add_ORT.get('Concentração caust 2 (g/l NaOH)', 1.0) / 1000.0
            consumo_dia = consumo_cel_periodo / DIAS_PERIODO
        if rendimento > 0 and concentracao > 0:
            licor_branco_fixo += 0.9 * consumo_dia * carga_alcalina / rendimento / concentracao
    licor_ang = dict_param_add_ORT.get('Licor branco Angatuba (m3 LB/d)', 0.0)
    fixed_caus = licor_branco_fixo + licor_ang

    # CKB-FC / CKB-FL: consumo_periodo / DIAS_PERIODO = t/dia
    fixed_ckbfc = dict_consumo_fibras_ORT.get('CKB-FC', 0.0) / DIAS_PERIODO
    fixed_ckbfl = dict_consumo_fibras_ORT.get('CKB-FL', 0.0) / DIAS_PERIODO

    # CDR fixo — TSS de Angatuba (taxa diária, não escalonada) +
    #            TSS da celulose fixa (base DIAS_PERIODO → dividir por DIAS_PERIODO)
    volume_angatuba = dict_param_add_ORT.get('Volume recebido de Licor Preto de Angatuba (m³/d)', 0.0)
    conc_angatuba   = dict_param_add_ORT.get('Concentração Licor Preto recebido de Angatuba (%)', 0.0)
    tss_angatuba_dia = volume_angatuba * conc_angatuba  # já é t_ss/dia

    fibras_cdr = ['CKB-FC', 'CKB-FL', 'CKN-FC', 'CKN-FL', 'BCTMP']
    tss_cel_fixo = 0.0
    for fibra in fibras_cdr:
        coef_ss = dict_consumo_especifico.get(('ORT', fibra), {}).get('SOLIDO-SECO', 0.0)
        consumo_cel_periodo = dict_consumo_fibras_ORT.get(fibra, 0.0)  # base DIAS_PERIODO
        tss_cel_fixo += -1 * coef_ss * consumo_cel_periodo
    # tss_cel_fixo está em base DIAS_PERIODO → dividir para obter t_ss/dia
    fixed_cdr = tss_cel_fixo / DIAS_PERIODO + tss_angatuba_dia

    # Evaporação ORT fixa — proveniente da celulose fixa + Angatuba
    mapa_conc_digestor = {
        'CKB-FC': 'Concentração Licor Preto gerado no digestor 1 (%)',
        'CKB-FL': 'Concentração Licor Preto gerado no digestor 2 (%)',
        'CKN-FC': 'Concentração Licor Preto gerado no digestor 3 (%)',
        'CKN-FL': 'Concentração Licor Preto gerado no digestor 4 (%)',
        'BCTMP':  'Concentração Licor Preto gerado na BCTMP (%)',
    }
    conc_saida_evap = dict_param_add_ORT.get('Concentração Licor Preto na Saída Evap1', 1.0)
    volume_lp_fixo = 0.0
    tss_total_fixo = 0.0
    for fibra, chave_conc in mapa_conc_digestor.items():
        coef_ss = dict_consumo_especifico.get(('ORT', fibra), {}).get('SOLIDO-SECO', 0.0)
        consumo_cel_periodo = dict_consumo_fibras_ORT.get(fibra, 0.0)  # base DIAS_PERIODO
        tss_fibra = -1 * coef_ss * consumo_cel_periodo  # base DIAS_PERIODO
        conc_digestor = dict_param_add_ORT.get(chave_conc, 1.0)
        if conc_digestor > 0:
            volume_lp_fixo += tss_fibra / conc_digestor
        tss_total_fixo += tss_fibra
    # volume_lp_fixo e tss_total_fixo estão em base DIAS_PERIODO
    # Adiciona contribuição de Angatuba: taxa diária × DIAS_PERIODO para mesma base
    volume_lp_fixo += volume_angatuba * DIAS_PERIODO
    tss_total_fixo += tss_angatuba_dia * DIAS_PERIODO
    # evap em base DIAS_PERIODO → converter para t_H2O/h (taxa horária instantânea)
    evap_periodo_fixo = (volume_lp_fixo - tss_total_fixo / conc_saida_evap) if conc_saida_evap > 0 else 0.0
    fixed_evap_ort = evap_periodo_fixo / (DIAS_PERIODO * 24.0) if DIAS_PERIODO > 0 else 0.0

    return {
        'Outorga Captação':  fixed_capt,
        'Outorga Emissario': fixed_emis,
        'Caustificação':     fixed_caus,
        'CKB-FC':            fixed_ckbfc,
        'CKB-FL':            fixed_ckbfl,
        'CDR':               fixed_cdr,
        'Evaporação ORT':    fixed_evap_ort,
    }


FIXED_ORT_PER_DAY = _compute_fixed_ort_per_day()


# ═══════════════════════════════════════════════════════
# SEÇÃO 6B — VALIDAÇÃO DO PASSO 1 (print de sanidade)
# ═══════════════════════════════════════════════════════

def _norm(s):
    return unicodedata.normalize('NFD', str(s).lower()).encode('ascii', 'ignore').decode()


def _validar_coeficientes():
    """
    Compara consumo total calculado por lotes contra os totais do dashboard_data_ort.json.
    Mostra separadamente a parte variável (lotes) e a parte fixa (celulose).
    """
    print("\n" + "="*70)
    print("VALIDAÇÃO PASSO 1 — Coeficientes por lote vs totais do dashboard ORT")
    print("="*70)
    print(f"\n  Termos fixos ORT (celulose/dia ou m³/h):")
    for k, v in FIXED_ORT_PER_DAY.items():
        print(f"    {k:<22}: {v:>12.4f}")

    restricoes = list(limites_ORT.keys())
    consumo_var = defaultdict(float)
    for prod, maq_dict in meta_anual.items():
        for maq, meta in maq_dict.items():
            if maq not in MAQUINAS_ORT:
                continue
            if dict_prod_por_MP.get((maq, prod), 0) != 1:
                continue
            n_lotes_aprox = meta / LOTE_T
            produtividade_pm = dict_produtividade_bruta.get((maq, prod), 0.0)
            _setup_val = PERDA_SETUP_T.get(maq, PERDA_SETUP_T_DEFAULT) if isinstance(PERDA_SETUP_T,
                                                                                     dict) else PERDA_SETUP_T
            horas_lote_pm = ((LOTE_T + _setup_val) / produtividade_pm
                             if produtividade_pm > 0 else 0.0)

            for r in restricoes:
                c = consumo_lote_restricao(prod, maq, r, 'ORT')
                if r in RESTRICOES_HORARIAS:
                    volume_total_ano = n_lotes_aprox * horas_lote_pm * c
                    consumo_var[r] += volume_total_ano / (365.0 * 24.0)
                else:
                    consumo_var[r] += c * n_lotes_aprox

    bal_norm = {_norm(k): v for k, v in balanco_ORT_dict.items()}
    print(f"\n  ORT:")
    for r in restricoes:
        lim = limites_ORT.get(r, 0)
        if lim >= 1e8:
            continue
        calc_var = consumo_var[r]
        fixed    = FIXED_ORT_PER_DAY.get(r, 0.0)
        calc_total = calc_var + fixed
        dash_entry = bal_norm.get(_norm(r), {})
        usado_dash = dash_entry.get('usado', None)
        if usado_dash is not None:
            ratio = calc_total / usado_dash if abs(usado_dash) > 1e-9 else float('inf')
            flag = "  OK" if 0.5 < ratio < 2.0 else "  !! DIVERGE"
            print(f"    {r:<22}: var={calc_var:>9.4f}  fix={fixed:>9.4f}  "
                  f"total={calc_total:>9.4f}  dash={usado_dash:>9.4f}  "
                  f"ratio={ratio:.3f}{flag}")
        else:
            print(f"    {r:<22}: var={calc_var:>9.4f}  fix={fixed:>9.4f}  "
                  f"total={calc_total:>9.4f}  (sem referência no dashboard)")

_validar_coeficientes()

# ═══════════════════════════════════════════════════════
# SEÇÃO 7 — ETAPA 1: MILP por ciclo (n_lotes[p,m])
# ═══════════════════════════════════════════════════════

def solve_ciclo(unit_name, maquinas_unit, limites, saldo_ciclo, meta_anual_unit,
                produzido_acumulado, dias_ciclo, dias_acumulados, num_ciclo,
                cortes_extras_t=None):
    """
    Etapa 1: resolve o MILP para um ciclo de dias_ciclo dias.

    saldo_ciclo[p][m]         = quanto ainda falta produzir no ano
    meta_anual_unit[p][m]     = meta anual original (para R4)
    produzido_acumulado[p][m] = já produzido nos ciclos anteriores
    dias_acumulados           = total de dias cobertos até o fim deste ciclo
    cortes_extras_t[(p,m)]    = toneladas a remover do saldo efetivo de (p,m)
                                antes de resolver (substitui cortes em lotes inteiros)

    Retorna: (n_lotes_resultado, volume_parcial_res, status_str, deficits_r4_aceitos)
      n_lotes_resultado[p][m]  = int   (lotes completos de LOTE_T)
      volume_parcial_res[(p,m)] = float (toneladas do lote parcial, 0 se inexistente)
    """
    # Saldo efetivo após descontar cortes em toneladas
    saldo_efetivo_raw = {}
    for p, maq_dict in saldo_ciclo.items():
        for m, saldo in maq_dict.items():
            corte = (cortes_extras_t or {}).get((p, m), 0.0)
            saldo_efetivo_raw[(p, m)] = max(0.0, saldo - corte)

    pares_elegiveis = [
        (p, m)
        for p, maq_dict in saldo_ciclo.items()
        for m, saldo in maq_dict.items()
        if m in maquinas_unit
        and saldo_efetivo_raw.get((p, m), 0) > 1e-3
        and dict_prod_por_MP.get((m, p), 0) == 1
        and dict_produtividade_bruta.get((m, p), 0) > 0
    ]

    if not pares_elegiveis:
        print(f"    {unit_name}: nenhum par elegível no ciclo {num_ciclo}")
        return {}, {}, 'sem_pares', {}

    # Horas disponíveis no ciclo por máquina (pro-rata do ano)
    horas_ciclo = {
        m: dict_tempo_carga[m] * dict_taxa_DISP[m] * (dias_ciclo / 365.0)
        for m in maquinas_unit
        if m in dict_tempo_carga
    }

    # Horas por lote CHEIO (LOTE_T + setup, por máquina)
    horas_por_lote = {
        (p, m): (LOTE_T + (PERDA_SETUP_T.get(m, PERDA_SETUP_T_DEFAULT) if isinstance(PERDA_SETUP_T, dict) else PERDA_SETUP_T)) / dict_produtividade_bruta[(m, p)]
        for (p, m) in pares_elegiveis
    }
    # Setup em horas para o lote parcial (setup CHEIO, independente do volume)
    setup_horas = {
        (p, m): (PERDA_SETUP_T.get(m, PERDA_SETUP_T_DEFAULT) if isinstance(PERDA_SETUP_T, dict) else PERDA_SETUP_T) / dict_produtividade_bruta[(m, p)]
        for (p, m) in pares_elegiveis
    }
    # Horas por tonelada produzida (sem setup)
    prod_inv = {
        (p, m): 1.0 / dict_produtividade_bruta[(m, p)]
        for (p, m) in pares_elegiveis
    }

    # Saldo efetivo e teto de lotes inteiros (floor; parcial fecha o restante)
    saldo_efetivo = {(p, m): saldo_efetivo_raw[(p, m)] for (p, m) in pares_elegiveis}
    lotes_teto = {
        (p, m): max(0, int(saldo_efetivo[(p, m)] // LOTE_T))
        for (p, m) in pares_elegiveis
    }

    # Piso do lote parcial: min(50%*LOTE_T, saldo_efetivo) para suportar nicho
    piso_parcial = {
        (p, m): min(FRACAO_MIN_LOTE_PARCIAL * LOTE_T, saldo_efetivo[(p, m)])
        for (p, m) in pares_elegiveis
    }

    # Limite das restrições de balanço para este ciclo
    limite_ciclo = {}
    for r, lim_base in limites.items():
        if lim_base >= 1e8:
            continue
        if r in RESTRICOES_HORARIAS:
            lim_periodo = lim_base * 24.0 * dias_ciclo
        else:
            lim_periodo = lim_base * dias_ciclo
        fixed_base = FIXED_ORT_PER_DAY.get(r, 0.0) if unit_name == 'ORT' else 0.0
        if r in RESTRICOES_HORARIAS:
            lim_periodo -= fixed_base * 24.0 * dias_ciclo
        else:
            lim_periodo -= fixed_base * dias_ciclo
        if lim_periodo > 1e-6:
            limite_ciclo[r] = lim_periodo

    model = pyo.ConcreteModel()
    idx = list(pares_elegiveis)

    model.n_lotes = pyo.Var(idx, domain=pyo.NonNegativeIntegers, initialize=0)
    for (p, m) in idx:
        model.n_lotes[p, m].setub(lotes_teto[(p, m)])

    # volume_parcial: contínuo em [0, LOTE_T]; sem binário para manter MILP rápido.
    # Piso (FRACAO_MIN) é enforçado no pós-processamento ao extrair o resultado.
    model.volume_parcial = pyo.Var(idx, domain=pyo.NonNegativeReals, initialize=0)
    for (p, m) in idx:
        model.volume_parcial[p, m].setub(LOTE_T)
    model.deficit_r4 = pyo.Var(idx, domain=pyo.NonNegativeReals)

    # Saldo: n_lotes*LOTE_T + volume_parcial <= saldo_efetivo (Caso A)
    def r_saldo_rule(mdl, p, m):
        return mdl.n_lotes[p, m] * LOTE_T + mdl.volume_parcial[p, m] <= saldo_efetivo[(p, m)]
    model.R_saldo = pyo.Constraint(idx, rule=r_saldo_rule)

    # R1 — Capacidade de horas no ciclo por máquina.
    # Setup do lote parcial é aproximado como proporcional ao volume (setup/LOTE_T * vol),
    # o que subestima em ≤5% para lotes pequenos — aceitável dado que pós-processamento
    # detecta e ajusta quaisquer violações reais.
    def r1_rule(mdl, maq):
        pares_maq = [(p, m) for (p, m) in idx if m == maq]
        if not pares_maq:
            return pyo.Constraint.Skip
        cap = horas_ciclo.get(maq, 0)
        if cap <= 0:
            return pyo.Constraint.Skip
        return (
            sum(
                mdl.n_lotes[p, m] * horas_por_lote[(p, m)]
                + mdl.volume_parcial[p, m] * (prod_inv[(p, m)] + setup_horas[(p, m)] / LOTE_T)
                for (p, m) in pares_maq
            ) <= cap
        )
    model.R1 = pyo.Constraint(maquinas_unit, rule=r1_rule)

    # R3 — Restrições de balanço agregadas no ciclo
    restricoes_ativas = list(limite_ciclo.keys())

    coef_r3 = {}
    coef_r3_per_ton = {}
    for (p, m) in idx:
        h_lote = horas_por_lote[(p, m)]
        for r in restricoes_ativas:
            c = consumo_lote_restricao(p, m, r, unit_name)
            if r in RESTRICOES_HORARIAS:
                coef_r3[(p, m, r)]         = c * h_lote
                coef_r3_per_ton[(p, m, r)] = c * h_lote / LOTE_T
            else:
                coef_r3[(p, m, r)]         = c * h_lote / 24.0
                coef_r3_per_ton[(p, m, r)] = c * h_lote / 24.0 / LOTE_T

    def r3_rule(mdl, r):
        lim = limite_ciclo.get(r, 0)
        if lim <= 0:
            return pyo.Constraint.Skip
        termos = []
        for (p, m) in idx:
            c_full = coef_r3.get((p, m, r), 0)
            c_ton  = coef_r3_per_ton.get((p, m, r), 0)
            if c_full > 1e-9:
                termos.append(mdl.n_lotes[p, m] * c_full)
            if c_ton > 1e-9:
                termos.append(mdl.volume_parcial[p, m] * c_ton)
        if not termos:
            return pyo.Constraint.Skip
        return sum(termos) <= lim
    model.R3 = pyo.Constraint(restricoes_ativas, rule=r3_rule)

    # R4 — Ritmo mínimo acumulado (vale desde o ciclo 1)
    def r4_rule(mdl, p, m):
        meta = meta_anual_unit.get(p, {}).get(m, 0)
        if meta <= 0:
            return pyo.Constraint.Skip
        acum = produzido_acumulado.get(p, {}).get(m, 0)
        rhs = meta * (dias_acumulados / float(DIAS_PERIODO)) * FATOR_TOLERANCIA_RITMO
        return (acum + mdl.n_lotes[p, m] * LOTE_T + mdl.volume_parcial[p, m]
                + mdl.deficit_r4[p, m] >= rhs)
    model.R4 = pyo.Constraint(idx, rule=r4_rule)

    # R_teto — Teto de produção por ciclo (distribui cross-ciclo)
    if FATOR_SPREAD:
        def r_teto_rule(mdl, p, m):
            meta = meta_anual_unit.get(p, {}).get(m, 0)
            if meta <= 0:
                return pyo.Constraint.Skip
            teto_t = meta * (dias_ciclo / float(DIAS_PERIODO)) * FATOR_SPREAD
            teto_t = max(teto_t, LOTE_T)  # garante ao menos 1 lote
            return (mdl.n_lotes[p, m] * LOTE_T + mdl.volume_parcial[p, m]
                    <= teto_t)
        model.R_teto = pyo.Constraint(idx, rule=r_teto_rule)

    # Objetivo: maximizar cobertura percentual (peso = 1/meta_anual)
    peso_cobertura = {}
    for (p, m) in idx:
        meta_pm = meta_anual_unit.get(p, {}).get(m, LOTE_T)
        peso_cobertura[(p, m)] = 1.0 / meta_pm if meta_pm > 0 else 0.0

    PESO_PENALIDADE_R4 = 1000.0
    model.obj = pyo.Objective(
        expr=(
            sum(
                (model.n_lotes[p, m] * LOTE_T + model.volume_parcial[p, m])
                * peso_cobertura[(p, m)]
                for (p, m) in idx
            )
            - PESO_PENALIDADE_R4 * sum(model.deficit_r4[p, m] for (p, m) in idx)
        ),
        sense=pyo.maximize,
    )

    solver = SolverFactory('glpk', executable=get_glpk_path())
    solver.options['tmlim']  = 120
    solver.options['mipgap'] = 0.05
    results = solver.solve(model, tee=False)

    from pyomo.opt import TerminationCondition as TC
    tc = results.solver.termination_condition
    if tc == TC.optimal:
        status_str = 'otimo'
    elif tc in (TC.feasible, TC.maxTimeLimit):
        status_str = 'subotimo'
    else:
        status_str = 'infeasivel'
        print(f"    AVISO: {unit_name} ciclo {num_ciclo} — {tc}")

    n_lotes_res = {}
    volume_parcial_res = {}
    deficits_r4_aceitos = {}
    if status_str != 'infeasivel':
        for (p, m) in idx:
            nl = pyo.value(model.n_lotes[p, m])
            if nl is not None and nl > 0.5:
                n_lotes_res.setdefault(p, {})[m] = int(round(nl))
            vp = pyo.value(model.volume_parcial[p, m])
            # Piso: descarta lotes parciais abaixo de FRACAO_MIN (sem binário no MILP,
            # o solver pode retornar volumes ínfimos; piso_parcial[p,m] já suporta nicho)
            if vp is not None and vp >= piso_parcial[(p, m)] - 1e-3:
                volume_parcial_res[(p, m)] = round(vp, 3)
            deficit = pyo.value(model.deficit_r4[p, m])
            if deficit is not None and deficit > 0.5:
                deficits_r4_aceitos[(p, m)] = round(deficit, 2)

    if deficits_r4_aceitos:
        print(f"    AVISO: {unit_name} ciclo {num_ciclo} — ritmo (R4) não "
              f"cumprido para {len(deficits_r4_aceitos)} par(es); déficits aceitos:")
        for (p, m), d in sorted(deficits_r4_aceitos.items(), key=lambda x: -x[1])[:10]:
            print(f"      {p} MP{m}: déficit={d:.1f}t")

    return n_lotes_res, volume_parcial_res, status_str, deficits_r4_aceitos


# ═══════════════════════════════════════════════════════
# SEÇÃO 7B — VERIFICAÇÃO DE PICO INSTANTÂNEO (restrições horárias)
# ═══════════════════════════════════════════════════════

def verificar_picos_horarios(schedule_ciclo, unit_name, limites):
    """
    Verifica, para cada dia do schedule, se o consumo de QUALQUER
    restrição (horária OU diária) excede o limite correspondente.

    Retorna:
      violacoes: dict {dia: {restricao: usado}} apenas para dias/restrições
                 que excedem o limite
      contribuicoes: dict {(dia, restricao): [(produto, maquina, valor), ...]}
      violacao_estrutural: bool
    """
    restricoes_ativas = [r for r in limites if limites.get(r, 1e9) < 1e8]
    if not restricoes_ativas:
        return {}, {}, False

    consumo_dia = defaultdict(lambda: defaultdict(float))
    contribuicoes = defaultdict(list)
    max_por_maquina_dia = defaultdict(lambda: defaultdict(dict))
    violacao_estrutural = False

    for maq, lotes in schedule_ciclo.items():
        for lote in lotes:
            produto   = lote['produto']
            dia_i     = lote['dia_inicio']
            dia_f     = lote['dia_fim']
            vol_lote  = lote.get('volume_t', LOTE_T)
            span_dias = max(1, dia_f - dia_i + 1)
            for r in restricoes_ativas:
                c_lote = consumo_lote_restricao(produto, maq, r, unit_name, vol=vol_lote)
                lim = limites.get(r, 1e9)
                if r in RESTRICOES_HORARIAS:
                    valor_no_dia = c_lote
                    if lim < 1e8 and valor_no_dia > lim:
                        violacao_estrutural = True
                    for dia in range(dia_i, dia_f + 1):
                        atual = max_por_maquina_dia[dia][r].get(maq, 0.0)
                        if valor_no_dia > atual:
                            max_por_maquina_dia[dia][r][maq] = valor_no_dia
                        contribuicoes[(dia, r)].append((produto, maq, valor_no_dia))
                else:
                    valor_no_dia = c_lote * 365.0 / span_dias
                    for dia in range(dia_i, dia_f + 1):
                        consumo_dia[dia][r] += valor_no_dia
                        contribuicoes[(dia, r)].append((produto, maq, valor_no_dia))

    for dia, restr_dict in max_por_maquina_dia.items():
        for r, por_maquina in restr_dict.items():
            consumo_dia[dia][r] += sum(por_maquina.values())

    violacoes = {}
    for dia, restr_dict in consumo_dia.items():
        for r, usado in restr_dict.items():
            if unit_name == 'ORT':
                usado = usado + FIXED_ORT_PER_DAY.get(r, 0.0)
            lim = limites.get(r, 1e9)
            if lim < 1e8 and usado > lim * 1.001:
                violacoes.setdefault(dia, {})[r] = usado

    return violacoes, contribuicoes, violacao_estrutural


def escolher_corte(violacoes, contribuicoes, idx, limites, unit_name):
    """
    Identifica a restrição/máquina mais violada e retorna exatamente
    QUANTAS TONELADAS cortar de cada par (produto, máquina) para eliminar
    o excesso — com precisão de fração de tonelada (lote parcial).

    Retorna lista de tuplas [(produto, maquina, toneladas_a_remover), ...].
    """
    if not violacoes:
        return []

    pior_dia, pior_r, pior_usado = None, None, -1
    for dia, restr_dict in violacoes.items():
        for r, usado in restr_dict.items():
            if usado > pior_usado:
                pior_usado = usado
                pior_dia, pior_r = dia, r

    pares = contribuicoes.get((pior_dia, pior_r), [])
    if not pares:
        return []

    eh_horaria = pior_r in RESTRICOES_HORARIAS
    por_maquina = {}
    if eh_horaria:
        for produto, maq, valor in pares:
            if maq not in por_maquina or valor > por_maquina[maq][1]:
                por_maquina[maq] = (produto, valor)
    else:
        soma_por_maq = defaultdict(float)
        produto_repr_por_maq = {}
        for produto, maq, valor in pares:
            soma_por_maq[maq] += valor
            if maq not in produto_repr_por_maq or valor > produto_repr_por_maq[maq][1]:
                produto_repr_por_maq[maq] = (produto, valor)
        for maq, soma in soma_por_maq.items():
            por_maquina[maq] = (produto_repr_por_maq[maq][0], soma)

    if not por_maquina:
        return []

    maquina_alvo = max(por_maquina.keys(), key=lambda m: por_maquina[m][1])

    lim = limites.get(pior_r, 1e9)
    if lim >= 1e8:
        return []
    excesso_total = pior_usado - lim

    cortes = []  # [(produto, maquina, toneladas_a_remover), ...]

    if eh_horaria:
        # HORÁRIAS: taxa c_lote (m³/h) é quase proporcional ao volume.
        # c_per_ton = c_lote / LOTE_T; tons_a_remover = excesso / c_per_ton.
        candidatos = []
        for (p, m) in idx:
            if m != maquina_alvo:
                continue
            c_lote = consumo_lote_restricao(p, m, pior_r, unit_name)
            if c_lote > 1e-9:
                candidatos.append((p, m, c_lote))
        candidatos.sort(key=lambda x: -x[2])  # maior taxa primeiro

        excesso_restante = excesso_total
        for (p, m, c_lote) in candidatos:
            if excesso_restante <= 1e-6:
                break
            c_per_ton = c_lote / LOTE_T
            tons_remover = excesso_restante / c_per_ton
            if tons_remover > 1e-6:
                cortes.append((p, m, tons_remover))
                excesso_restante -= tons_remover * c_per_ton  # ≈ 0
    else:
        # DIÁRIAS: 'valor' já está na escala concentrada (c_lote*365/span_dias).
        # c_per_ton = c_lote_escala / LOTE_T; tons = excesso / c_per_ton.
        contrib_real_por_prod = defaultdict(float)
        contagem_lotes_por_prod = defaultdict(int)
        for produto, maq, valor in pares:
            if maq == maquina_alvo:
                contrib_real_por_prod[produto] += valor
                contagem_lotes_por_prod[produto] += 1

        produtos_ordenados = sorted(contrib_real_por_prod.items(), key=lambda x: -x[1])
        excesso_restante = excesso_total
        for produto, valor_total in produtos_ordenados:
            if excesso_restante <= 1e-6:
                break
            n_lotes_deste = max(1, contagem_lotes_por_prod[produto])
            c_lote_escala = valor_total / n_lotes_deste
            if c_lote_escala <= 1e-9:
                continue
            c_per_ton = c_lote_escala / LOTE_T
            tons_remover = excesso_restante / c_per_ton
            if tons_remover > 1e-6:
                cortes.append((produto, maquina_alvo, tons_remover))
                excesso_restante -= tons_remover * c_per_ton  # ≈ 0
            if len(cortes) >= 3:
                break

    return cortes


MAX_DESLOCAMENTOS   = 60  # máx deslocamentos temporais por iteração MILP
MAX_ITERACOES_PICO  = 25  # limite de segurança contra loop infinito (fallback corte)


def _repacotar_lotes_a_partir(lotes, dia_inicio, hora_inicio, horas_dia_maq, dia_max):
    """Repacota `lotes` sequencialmente a partir de (dia_inicio, hora_inicio).
    Retorna lista reempacotada ou None se extrapolar dia_max."""
    resultado = []
    dia_atual  = dia_inicio
    hora_atual = hora_inicio
    for lote in lotes:
        duracao = lote['duracao_h']
        if dia_atual > dia_max:
            return None
        dia_inicio_lote = dia_atual
        hora_inicio_lote = hora_atual
        horas_restantes = duracao
        horas_por_dia_lote: dict = {}
        while horas_restantes > 1e-6:
            if dia_atual > dia_max:
                return None
            cap = horas_dia_maq - hora_atual
            if cap <= 1e-6:
                dia_atual  += 1
                hora_atual  = 0.0
                if dia_atual > dia_max:
                    return None
                cap = horas_dia_maq
            h_neste = min(horas_restantes, cap)
            horas_por_dia_lote[dia_atual] = horas_por_dia_lote.get(dia_atual, 0.0) + h_neste
            hora_atual        += h_neste
            horas_restantes   -= h_neste
            if hora_atual >= horas_dia_maq - 1e-6:
                dia_atual  += 1
                hora_atual  = 0.0
        dia_fim_lote = max(horas_por_dia_lote) if horas_por_dia_lote else dia_inicio_lote
        entry = dict(lote)
        entry['dia_inicio']    = dia_inicio_lote
        entry['dia_fim']       = dia_fim_lote
        entry['hora_inicio']   = round(hora_inicio_lote, 3)
        entry['horas_por_dia'] = {str(d): round(h, 3) for d, h in horas_por_dia_lote.items()}
        resultado.append(entry)
    return resultado


def _deslocar_lote_no_schedule(schedule_maq, idx_lote, horas_dia_maq, dia_max):
    """Desloca o lote em idx_lote para o início do dia seguinte ao seu dia_inicio,
    repacotando em cascata todos os lotes subsequentes da mesma máquina.
    Retorna novo schedule ou None se extrapolar dia_max."""
    if idx_lote >= len(schedule_maq):
        return None
    novo_dia = schedule_maq[idx_lote]['dia_inicio'] + 1
    if novo_dia > dia_max:
        return None
    novos = _repacotar_lotes_a_partir(
        schedule_maq[idx_lote:], novo_dia, 0.0, horas_dia_maq, dia_max)
    if novos is None:
        return None
    return schedule_maq[:idx_lote] + novos


def _tentar_deslocamento_temporal(schedule_ciclo, dia_violado, restricao, unit_name,
                                   maquinas_unit, dia_max):
    """Identifica todos os lotes rodando no dia_violado com contribuição > 0
    para a restricao, escolhe o de MENOR volume e o desloca para o dia seguinte.
    Retorna (schedule_novo, sucesso)."""
    candidatos = []
    for maq in maquinas_unit:
        for i, lote in enumerate(schedule_ciclo.get(maq, [])):
            if lote['dia_inicio'] <= dia_violado <= lote['dia_fim']:
                p = lote['produto']
                v = lote.get('volume_t', LOTE_T)
                if consumo_lote_restricao(p, maq, restricao, unit_name, vol=v) > 1e-9:
                    candidatos.append((v, i, maq))
    if not candidatos:
        return schedule_ciclo, False
    candidatos.sort(key=lambda x: x[0])  # menor volume primeiro
    for _vol, idx_c, maq_c in candidatos:
        novo_sched_maq = _deslocar_lote_no_schedule(
            schedule_ciclo[maq_c], idx_c, horas_dia.get(maq_c, 0.0), dia_max)
        if novo_sched_maq is not None:
            novo = dict(schedule_ciclo)
            novo[maq_c] = novo_sched_maq
            return novo, True
    return schedule_ciclo, False


def _loop_deslocamento_horario(schedule_ciclo, unit_name, limites, maquinas_unit, dia_max):
    """Itera deslocamentos temporais para resolver violações de restrições HORÁRIAS.
    Retorna (schedule, n_deslocamentos_aplicados, ainda_ha_horaria_violada)."""
    n_desl = 0
    for _ in range(MAX_DESLOCAMENTOS):
        violacoes, _, viol_estrutural = verificar_picos_horarios(
            schedule_ciclo, unit_name, limites)
        if viol_estrutural:
            # Lote isolado já excede o limite — deslocamento não resolve
            return schedule_ciclo, n_desl, True
        viol_horarias = {
            (dia, r): usado
            for dia, rd in violacoes.items()
            for r, usado in rd.items()
            if r in RESTRICOES_HORARIAS
        }
        if not viol_horarias:
            return schedule_ciclo, n_desl, False  # resolveu
        (pior_dia, pior_r), _ = max(viol_horarias.items(), key=lambda x: x[1])
        schedule_ciclo, ok = _tentar_deslocamento_temporal(
            schedule_ciclo, pior_dia, pior_r, unit_name, maquinas_unit, dia_max)
        if not ok:
            return schedule_ciclo, n_desl, True  # não conseguiu deslocar nenhum
        n_desl += 1
    return schedule_ciclo, n_desl, True  # esgotou iterações


# ═══════════════════════════════════════════════════════
# CAMADA 3 — SWAP DE COMPOSIÇÃO (reordenação dentro da mesma máquina)
# ═══════════════════════════════════════════════════════

def _taxa_horaria_por_maquina_no_dia(schedule_ciclo, dia, restricao, unit_name):
    """Retorna {maq: (produto_dominante, taxa_max)} para uma restrição horária num dia."""
    por_maq = {}
    for maq, lotes in schedule_ciclo.items():
        for lote in lotes:
            if lote['dia_inicio'] <= dia <= lote['dia_fim']:
                t = consumo_lote_restricao(
                    lote['produto'], maq, restricao, unit_name,
                    vol=lote.get('volume_t', LOTE_T))
                if t > por_maq.get(maq, (None, 0.0))[1]:
                    por_maq[maq] = (lote['produto'], t)
    return por_maq


def _get_product_block(lotes_maq, produto, dia):
    """Retorna (start_idx, end_idx) do bloco contíguo de `produto` que cobre `dia`.
    Expande para incluir todos os lotes consecutivos do mesmo produto."""
    anchor = next((i for i, l in enumerate(lotes_maq)
                   if l['produto'] == produto
                   and l['dia_inicio'] <= dia <= l['dia_fim']), None)
    if anchor is None:
        return None, None
    start = anchor
    while start > 0 and lotes_maq[start - 1]['produto'] == produto:
        start -= 1
    end = anchor
    while end + 1 < len(lotes_maq) and lotes_maq[end + 1]['produto'] == produto:
        end += 1
    return start, end


def _simular_group_reorder(lotes_maq, block_a_start, block_a_end,
                            block_b_start, block_b_end, horas_dia_maq, dia_max):
    """
    Reordena dois blocos: coloca block_b ANTES de block_a (preservando o meio).
    Assume block_b_start > block_a_end.

    Estratégia de repack: repacota APENAS a janela [block_a_start .. block_b_end]
    (lotes B + meio + A). O sufixo pós-block_b_end mantém as posições originais.
    Isso evita falha de overflow quando o sufixo já está no limite de dia_max.
    Usa dia_max_local = fim original do último lote de block_b como teto do repack.
    """
    if block_b_start <= block_a_end:
        return None
    prefix  = lotes_maq[:block_a_start]
    bloco_A = lotes_maq[block_a_start:block_a_end + 1]
    meio    = lotes_maq[block_a_end + 1:block_b_start]
    bloco_B = lotes_maq[block_b_start:block_b_end + 1]
    suffix  = lotes_maq[block_b_end + 1:]

    # Teto local = fim do último lote do bloco B original (a janela deve caber nele)
    dia_max_local = lotes_maq[block_b_end].get('dia_fim', dia_max)

    janela = bloco_B + meio + bloco_A
    dia_r  = lotes_maq[block_a_start]['dia_inicio']
    hora_r = lotes_maq[block_a_start].get('hora_inicio', 0.0)
    repacked_janela = _repacotar_lotes_a_partir(
        janela, dia_r, hora_r, horas_dia_maq, dia_max_local)
    if repacked_janela is None:
        return None
    return prefix + repacked_janela + suffix


def _seq_unica(seq):
    """Remove repetições consecutivas de uma sequência de produtos."""
    out = []
    for p in seq:
        if not out or p != out[-1]:
            out.append(p)
    return out


def _tentar_swap_composicao_horario(schedule_ciclo, dia_violado, restricao, unit_name,
                                     maquinas_unit, limites, dia_max):
    """
    Tenta resolver violação horária no dia_violado por REORDENAÇÃO DE BLOCOS de produto
    na mesma máquina: move um bloco de produto de baixa taxa para antes do bloco de
    alta taxa ativo no dia_violado (sem alterar volumes ou metas).

    Estratégia:
      1. Identifica o produto dominante (alta taxa) em cada máquina no dia_violado
      2. Localiza o BLOCO inteiro desse produto na fila da máquina
      3. Busca outros blocos de produto na mesma máquina com taxa menor
      4. Pré-filtro analítico: redução suficiente para resolver D
      5. Simulação completa (repack + verificar_picos) para os top-5 candidatos
      6. Aceita o primeiro que resolve D sem criar nova violação em dia limpo

    Retorna (schedule_novo, sucesso).
    """
    lim = limites.get(restricao, 1e9)
    if lim >= 1e8:
        return schedule_ciclo, False

    por_maq_D = _taxa_horaria_por_maquina_no_dia(
        schedule_ciclo, dia_violado, restricao, unit_name)
    total_D = sum(t for _, t in por_maq_D.values())
    if total_D <= lim * 1.001:
        return schedule_ciclo, False

    viol_antes, _, _ = verificar_picos_horarios(schedule_ciclo, unit_name, limites)

    candidatos = []
    for maq in maquinas_unit:
        if maq not in por_maq_D:
            continue
        prod_alto, taxa_a = por_maq_D[maq]
        if taxa_a <= 1e-9:
            continue

        lotes_maq = schedule_ciclo.get(maq, [])
        ba_start, ba_end = _get_product_block(lotes_maq, prod_alto, dia_violado)
        if ba_start is None:
            continue

        # Busca blocos distintos com taxa menor que aparecem DEPOIS do bloco alto
        produtos_vistos = set()
        for idx_b in range(ba_end + 1, len(lotes_maq)):
            prod_b = lotes_maq[idx_b]['produto']
            if prod_b in produtos_vistos or prod_b == prod_alto:
                continue
            produtos_vistos.add(prod_b)

            taxa_b = consumo_lote_restricao(
                prod_b, maq, restricao, unit_name,
                vol=lotes_maq[idx_b].get('volume_t', LOTE_T))
            reducao = taxa_a - taxa_b
            if reducao <= 1e-6:
                continue
            if total_D - reducao > lim * 1.001:
                continue

            bb_start, bb_end = _get_product_block(
                lotes_maq, prod_b, lotes_maq[idx_b]['dia_inicio'])
            if bb_start is None or bb_start <= ba_end:
                continue

            # Pré-filtro geométrico: Block B deve ter duração suficiente para cobrir
            # dia_violado sozinho (caso contrário o segmento "meio" cobre dia_violado
            # com taxa possivelmente alta → D continua violado após o reorder).
            hd_maq = horas_dia.get(maq, 0.0)
            if hd_maq > 0:
                dur_B_h = sum(lotes_maq[i]['duracao_h']
                              for i in range(bb_start, bb_end + 1))
                block_a_day0 = lotes_maq[ba_start]['dia_inicio']
                hora_a0      = lotes_maq[ba_start].get('hora_inicio', 0.0)
                h_ate_fim_D  = (dia_violado - block_a_day0 + 1) * hd_maq - hora_a0
                if dur_B_h < h_ate_fim_D - 1e-3:
                    continue  # Block B não cobre dia_violado inteiramente

            candidatos.append((reducao, maq, ba_start, ba_end, bb_start, bb_end))

    if not candidatos:
        return schedule_ciclo, False

    candidatos.sort(reverse=True)
    for reducao, maq, ba_start, ba_end, bb_start, bb_end in candidatos[:5]:
        lotes_maq = schedule_ciclo.get(maq, [])
        hd = horas_dia.get(maq, 0.0)
        novos_lotes = _simular_group_reorder(
            lotes_maq, ba_start, ba_end, bb_start, bb_end, hd, dia_max)
        if novos_lotes is None:
            continue

        # Filtro de custo de setup para ORT: rejeita swap que piora setup além do limiar
        if unit_name == 'ORT' and maq in MAQUINAS_ORT:
            seq_antes = [l['produto'] for l in lotes_maq]
            seq_depois = [l['produto'] for l in novos_lotes]
            custo_antes = custo_sequencia_setup(maq, _seq_unica(seq_antes))
            custo_depois = custo_sequencia_setup(maq, _seq_unica(seq_depois))
            delta_setup = custo_depois - custo_antes
            if delta_setup > LIMIAR_PENALIDADE_SWAP_T:
                continue  # swap muito caro em termos de setup — rejeitar

        sched_test = dict(schedule_ciclo)
        sched_test[maq] = novos_lotes
        viol_test, _, _ = verificar_picos_horarios(sched_test, unit_name, limites)

        if viol_test.get(dia_violado, {}).get(restricao) is not None:
            continue  # D ainda violado

        criou_nova = any(
            restricao in rd and viol_antes.get(d, {}).get(restricao) is None
            for d, rd in viol_test.items()
        )
        if criou_nova:
            continue

        novo_sched = dict(schedule_ciclo)
        novo_sched[maq] = novos_lotes
        return novo_sched, True

    return schedule_ciclo, False


def _loop_swap_composicao(schedule_ciclo, unit_name, limites, maquinas_unit, dia_max):
    """
    Camada 3 (nova, roda ANTES do deslocamento temporal): resolve violações HORÁRIAS
    reordenando lotes dentro da mesma máquina.
    Retorna (schedule, n_swaps, ainda_ha_horaria_violada).
    """
    n_swaps = 0
    for _ in range(MAX_ITERACOES_PICO):
        violacoes, _, viol_estrutural = verificar_picos_horarios(
            schedule_ciclo, unit_name, limites)
        if viol_estrutural:
            return schedule_ciclo, n_swaps, True
        viol_hor = {
            (d, r): v
            for d, rd in violacoes.items()
            for r, v in rd.items()
            if r in RESTRICOES_HORARIAS
        }
        if not viol_hor:
            return schedule_ciclo, n_swaps, False
        (pior_dia, pior_r), _ = max(viol_hor.items(), key=lambda x: x[1])
        schedule_ciclo, ok = _tentar_swap_composicao_horario(
            schedule_ciclo, pior_dia, pior_r, unit_name, maquinas_unit, limites, dia_max)
        if not ok:
            return schedule_ciclo, n_swaps, True
        n_swaps += 1
    return schedule_ciclo, n_swaps, True


def _build_schedule_ciclo(n_lotes_res, volume_parcial_res, maquinas_unit, dias_do_ciclo,
                           limites, unit_name, ultimo_produto_por_maquina=None):
    """Constrói o schedule do ciclo a partir do resultado do MILP."""
    schedule = {}
    for maq in maquinas_unit:
        produtos_maq = {
            p: n
            for p, maq_dict in n_lotes_res.items()
            for m2, n in maq_dict.items()
            if m2 == maq and n > 0
        }
        vol_parcial_maq = {
            p: vol
            for (p, m2), vol in volume_parcial_res.items()
            if m2 == maq and vol > 1e-6
        }
        if not produtos_maq and not vol_parcial_maq:
            schedule[maq] = []
            continue
        _setup_maq = PERDA_SETUP_T.get(maq, PERDA_SETUP_T_DEFAULT) if isinstance(PERDA_SETUP_T, dict) else PERDA_SETUP_T
        hpl_maq = {
            p: (LOTE_T + _setup_maq) / dict_produtividade_bruta.get((maq, p), 1.0)
            for p in produtos_maq
        }
        prod_bruta_maq = {
            p: dict_produtividade_bruta.get((maq, p), 0)
            for p in set(list(produtos_maq.keys()) + list(vol_parcial_maq.keys()))
        }
        hd = horas_dia.get(maq, 0)
        max_lotes_por_dia_maq = {}
        for p in produtos_maq:
            limite_seguro_p = None
            for r, lim_base in limites.items():
                if r in RESTRICOES_HORARIAS or lim_base >= 1e8:
                    continue
                c_lote = consumo_lote_restricao(p, maq, r, unit_name)
                consumo_total_lote = c_lote * 365.0
                if consumo_total_lote > 1e-9:
                    max_seguro_r = max(1, int(lim_base / consumo_total_lote))
                    if limite_seguro_p is None or max_seguro_r < limite_seguro_p:
                        limite_seguro_p = max_seguro_r
            if limite_seguro_p is not None:
                max_lotes_por_dia_maq[p] = limite_seguro_p
        # Último produto do ciclo anterior nesta máquina (para continuidade de setup)
        _upm = ultimo_produto_por_maquina or {}
        p_ultimo_anterior = _upm.get(maq) if unit_name == 'ORT' else None

        schedule[maq] = empacotar_lotes(
            produtos_e_quantidades=produtos_maq,
            dias_disponiveis=dias_do_ciclo,
            horas_por_lote_maq=hpl_maq,
            horas_dia_maq=hd,
            janela_max=JANELA_MAX_DIAS,
            max_lotes_por_dia=max_lotes_por_dia_maq,
            volumes_parciais=vol_parcial_maq,
            produtividade_maq=prod_bruta_maq,
            maquina=maq,
            p_inicial=p_ultimo_anterior,
        )
    return schedule


def solve_ciclo_com_pico_controlado(
    unit_name, maquinas_unit, limites, saldo_ciclo, meta_anual_unit,
    produzido_acumulado, dias_ciclo, dias_acumulados, num_ciclo,
    ultimo_produto_por_maquina=None,
):
    """
    Envolve solve_ciclo() + empacotar_lotes() em um loop que:
      1. Tenta resolver violações HORÁRIAS por DESLOCAMENTO TEMPORAL (sem re-MILP)
      2. Cai para CORTE DE TONELADAS (re-MILP) como fallback

    Retorna: (n_lotes_res, volume_parcial_res, status_str, schedule_ciclo,
              pico_nao_resolvido, violacoes_pico_finais, n_iteracoes, stats_desl)
      stats_desl = {'n_deslocamentos': int, 'n_hor_resolvidas_desl': int,
                    'n_fallbacks_corte': int}
    """
    cortes_extras_t = {}
    n_iteracoes = 0
    n_lotes_res = {}
    volume_parcial_res = {}
    status_str = 'sem_pares'
    schedule_ciclo = {}
    violacoes = {}
    _corte_anterior = None
    dia_max_ciclo = dias_acumulados  # limite do ciclo para deslocamentos

    stats = {'n_deslocamentos': 0, 'n_hor_resolvidas_desl': 0, 'n_fallbacks_corte': 0,
             'n_swaps': 0, 'n_hor_resolvidas_swap': 0}

    for it in range(MAX_ITERACOES_PICO):
        n_iteracoes = it + 1

        n_lotes_res, volume_parcial_res, status_str, _deficits = solve_ciclo(
            unit_name, maquinas_unit, limites,
            saldo_ciclo, meta_anual_unit, produzido_acumulado,
            dias_ciclo, dias_acumulados, num_ciclo,
            cortes_extras_t=cortes_extras_t,
        )

        tem_producao = bool(n_lotes_res) or bool(volume_parcial_res)
        if status_str == 'infeasivel' or not tem_producao:
            return (n_lotes_res, volume_parcial_res, status_str, {},
                    False, {}, n_iteracoes, stats)

        dias_do_ciclo = list(range(dias_acumulados - dias_ciclo + 1, dias_acumulados + 1))
        schedule_ciclo = _build_schedule_ciclo(
            n_lotes_res, volume_parcial_res, maquinas_unit,
            dias_do_ciclo, limites, unit_name,
            ultimo_produto_por_maquina=ultimo_produto_por_maquina or {},
        )

        # ── FASE 0 (nova): swap de composição para restrições HORÁRIAS ──
        viol_pre_swap, _, _ = verificar_picos_horarios(schedule_ciclo, unit_name, limites)
        n_hor_pre_swap = sum(
            1 for rd in viol_pre_swap.values() for r in rd if r in RESTRICOES_HORARIAS)

        schedule_ciclo, n_swaps_it, _ = _loop_swap_composicao(
            schedule_ciclo, unit_name, limites, maquinas_unit, dia_max_ciclo)

        stats['n_swaps'] += n_swaps_it
        if n_swaps_it > 0:
            viol_pos_swap, _, _ = verificar_picos_horarios(schedule_ciclo, unit_name, limites)
            n_hor_pos_swap = sum(
                1 for rd in viol_pos_swap.values() for r in rd if r in RESTRICOES_HORARIAS)
            stats['n_hor_resolvidas_swap'] += max(0, n_hor_pre_swap - n_hor_pos_swap)
            print(f"    {unit_name} ciclo {num_ciclo} it{n_iteracoes}: "
                  f"swap composição: {n_swaps_it} swap(s), "
                  f"viol.hor {n_hor_pre_swap}→{n_hor_pos_swap}")

        # ── FASE 1: deslocamento temporal para restrições HORÁRIAS ──
        viol_antes, _, _ = verificar_picos_horarios(schedule_ciclo, unit_name, limites)
        n_hor_antes = sum(
            1 for rd in viol_antes.values() for r in rd if r in RESTRICOES_HORARIAS)

        schedule_ciclo, n_desl, ainda_horaria = _loop_deslocamento_horario(
            schedule_ciclo, unit_name, limites, maquinas_unit, dia_max_ciclo)

        stats['n_deslocamentos'] += n_desl

        viol_depois, _, _ = verificar_picos_horarios(schedule_ciclo, unit_name, limites)
        n_hor_depois = sum(
            1 for rd in viol_depois.values() for r in rd if r in RESTRICOES_HORARIAS)
        stats['n_hor_resolvidas_desl'] += max(0, n_hor_antes - n_hor_depois)

        if n_desl > 0:
            print(f"    {unit_name} ciclo {num_ciclo} it{n_iteracoes}: "
                  f"deslocamento temporal: {n_desl} lote(s) deslocado(s), "
                  f"viol.hor {n_hor_antes}→{n_hor_depois}")

        # ── Verificação final pós-deslocamento ──
        violacoes, contribuicoes, viol_estrutural = verificar_picos_horarios(
            schedule_ciclo, unit_name, limites)

        if not violacoes:
            if it > 0 or n_desl > 0:
                print(f"    {unit_name} ciclo {num_ciclo}: "
                      f"pico resolvido em it{n_iteracoes} "
                      f"(desl={stats['n_deslocamentos']} cortes={stats['n_fallbacks_corte']})")
            return (n_lotes_res, volume_parcial_res, status_str, schedule_ciclo,
                    False, {}, n_iteracoes, stats)

        if viol_estrutural:
            print(f"    AVISO: {unit_name} ciclo {num_ciclo} — violação ESTRUTURAL de pico "
                  f"(1 lote isolado já excede o limite)")
            return (n_lotes_res, volume_parcial_res, status_str, schedule_ciclo,
                    True, violacoes, n_iteracoes, stats)

        # ── FASE 2 (fallback): corte de toneladas — re-MILP ──
        idx_ciclo = [
            (p, m)
            for p, maq_dict in saldo_ciclo.items()
            for m, saldo in maq_dict.items()
            if m in maquinas_unit
            and saldo > 1e-3
            and dict_prod_por_MP.get((m, p), 0) == 1
            and dict_produtividade_bruta.get((m, p), 0) > 0
        ]

        cortes_aplicar = escolher_corte(violacoes, contribuicoes, idx_ciclo,
                                        limites, unit_name)
        if not cortes_aplicar:
            return (n_lotes_res, volume_parcial_res, status_str, schedule_ciclo,
                    True, violacoes, n_iteracoes, stats)

        corte_key = tuple(sorted((p, m, round(t, 1)) for p, m, t in cortes_aplicar))
        if corte_key == _corte_anterior:
            print(f"    {unit_name} ciclo {num_ciclo} it{n_iteracoes}: "
                  f"pico sem progresso após corte repetido — estrutural residual")
            return (n_lotes_res, volume_parcial_res, status_str, schedule_ciclo,
                    True, violacoes, n_iteracoes, stats)
        _corte_anterior = corte_key

        stats['n_fallbacks_corte'] += 1
        for (p, m, tons) in cortes_aplicar:
            cortes_extras_t[(p, m)] = cortes_extras_t.get((p, m), 0.0) + tons
        resumo_corte = ", ".join(
            f"{p}/MP{m}(-{tons:.1f}t)" for p, m, tons in cortes_aplicar)
        print(f"    {unit_name} ciclo {num_ciclo} it{n_iteracoes}: "
              f"fallback corte: {resumo_corte}")

    print(f"    AVISO: {unit_name} ciclo {num_ciclo} — não convergiu em "
          f"{MAX_ITERACOES_PICO} iterações; pico residual mantido")
    return (n_lotes_res, volume_parcial_res, status_str, schedule_ciclo,
            True, violacoes, n_iteracoes, stats)


# ═══════════════════════════════════════════════════════
# SEÇÃO 8 — ETAPA 2: EMPACOTAMENTO DE LOTES
# ═══════════════════════════════════════════════════════

def empacotar_lotes(
    produtos_e_quantidades: dict,
    dias_disponiveis: list,
    horas_por_lote_maq: dict,
    horas_dia_maq: float,
    janela_max: int = JANELA_MAX_DIAS,
    max_lotes_por_dia: dict = None,
    volumes_parciais: dict = None,
    produtividade_maq: dict = None,
    maquina: int = None,
    p_inicial: str = None,
) -> list:
    """
    Etapa 2 (Fase 1): empacota lotes numa máquina, agrupando todos os
    lotes do mesmo produto antes de trocar.

    volumes_parciais:  {produto: volume_t}  — lotes parciais a agendar após os cheios.
    produtividade_maq: {produto: t/h}       — necessário para calcular duração do parcial.
    """
    schedule = []

    # Ordenação base: decrescente por número de lotes
    produtos_base = sorted(
        [(p, n) for p, n in produtos_e_quantidades.items() if n > 0],
        key=lambda x: -x[1]
    )

    # Para ORT: reordenar pelo menor custo de transição (vizinho mais próximo)
    if maquina in MAQUINAS_ORT and len(produtos_base) > 1:
        nao_visitados = list(produtos_base)
        produtos_ordenados = []
        # Ponto de partida: produto com mais lotes (primeiro da lista base)
        # ou continuação do produto anterior ao ciclo (p_inicial)
        if p_inicial is not None:
            atual = p_inicial
        else:
            atual = nao_visitados[0][0]
            produtos_ordenados.append(nao_visitados.pop(0))

        while nao_visitados:
            melhor_idx = min(
                range(len(nao_visitados)),
                key=lambda i: get_setup_t(maquina, atual, nao_visitados[i][0])
            )
            proximo = nao_visitados.pop(melhor_idx)
            produtos_ordenados.append(proximo)
            atual = proximo[0]
    else:
        produtos_ordenados = produtos_base

    if not produtos_ordenados and not volumes_parciais:
        return schedule
    if not dias_disponiveis:
        return schedule

    max_limites = max_lotes_por_dia or {}

    # ── Pré-cálculo do espaçamento uniforme ────────────────────────────
    _sequencia_lotes = []
    for produto, n_lotes in produtos_ordenados:
        for _ in range(n_lotes):
            _sequencia_lotes.append(produto)
    _N_lotes_total = len(_sequencia_lotes)

    _total_dias_disp = len(dias_disponiveis)
    _fator_disp = FATOR_DISPERSAO_ORT if DISPERSAO_UNIFORME_ORT else 0.0
    _intervalo_dias  = (
        (_total_dias_disp * _fator_disp / _N_lotes_total)
        if (_N_lotes_total > 0 and _fator_disp > 0 and _total_dias_disp > 0)
        else 0.0
    )
    _lote_global_idx = 0
    # ────────────────────────────────────────────────────────────────────

    idx_dia = 0
    hora_atual = 0.0

    def _agendar_lote(produto, hpl, lote_num, volume_t=None, eh_parcial=False):
        """Agenda um lote com `hpl` horas; atualiza idx_dia e hora_atual no escopo pai."""
        nonlocal idx_dia, hora_atual
        if idx_dia >= len(dias_disponiveis):
            return

        dia_inicio_abs = dias_disponiveis[idx_dia]
        hora_inicio = hora_atual
        horas_restantes = hpl
        horas_por_dia_lote = {}

        while horas_restantes > 1e-6 and idx_dia < len(dias_disponiveis):
            cap_restante_dia = horas_dia_maq - hora_atual
            if cap_restante_dia <= 1e-6:
                idx_dia += 1
                hora_atual = 0.0
                if idx_dia < len(dias_disponiveis):
                    cap_restante_dia = horas_dia_maq
                else:
                    break
            horas_neste_dia = min(horas_restantes, cap_restante_dia)
            dia_abs_atual = dias_disponiveis[idx_dia]
            horas_por_dia_lote[dia_abs_atual] = (
                horas_por_dia_lote.get(dia_abs_atual, 0.0) + horas_neste_dia)
            hora_atual += horas_neste_dia
            horas_restantes -= horas_neste_dia
            if hora_atual >= horas_dia_maq - 1e-6:
                idx_dia += 1
                hora_atual = 0.0

        dia_fim_abs = dias_disponiveis[min(idx_dia, len(dias_disponiveis) - 1)]

        entry = {
            'produto':       produto,
            'lote_num':      lote_num,
            'dia_inicio':    dia_inicio_abs,
            'dia_fim':       dia_fim_abs,
            'hora_inicio':   round(hora_inicio, 3),
            'duracao_h':     round(hpl, 3),
            'horas_por_dia': {str(d): round(h, 3) for d, h in horas_por_dia_lote.items()},
        }
        if eh_parcial and volume_t is not None:
            entry['parcial']  = True
            entry['volume_t'] = round(volume_t, 3)
        else:
            entry['volume_t'] = LOTE_T
        schedule.append(entry)

    for produto, n_lotes in produtos_ordenados:
        hpl = horas_por_lote_maq.get(produto, 0)
        if hpl <= 0:
            continue

        limite_dia_produto = max_limites.get(produto)
        lotes_hoje = 0
        dia_civil_referencia = dias_disponiveis[idx_dia] if idx_dia < len(dias_disponiveis) else None

        for lote_num in range(1, n_lotes + 1):
            if idx_dia >= len(dias_disponiveis):
                break

            # ── Avanço de ponteiro para espaçamento uniforme ───────────
            if _intervalo_dias > 0 and _lote_global_idx > 0:
                _dia_alvo_abs  = dias_disponiveis[0] + _lote_global_idx * _intervalo_dias
                _dia_atual_abs = dias_disponiveis[idx_dia]
                if _dia_atual_abs < _dia_alvo_abs:
                    _idx_alvo = 0
                    for _di, _d in enumerate(dias_disponiveis):
                        if _d >= _dia_alvo_abs:
                            _idx_alvo = _di
                            break
                    else:
                        _idx_alvo = len(dias_disponiveis) - 1
                    if _idx_alvo > idx_dia:
                        idx_dia    = _idx_alvo
                        hora_atual = 0.0
            # ──────────────────────────────────────────────────────────

            dia_civil_atual = dias_disponiveis[idx_dia]
            if dia_civil_atual != dia_civil_referencia:
                dia_civil_referencia = dia_civil_atual
                lotes_hoje = 0

            if limite_dia_produto is not None and lotes_hoje >= limite_dia_produto:
                idx_dia += 1
                hora_atual = 0.0
                lotes_hoje = 0
                if idx_dia >= len(dias_disponiveis):
                    break
                dia_civil_referencia = dias_disponiveis[idx_dia]

            _agendar_lote(produto, hpl, lote_num, volume_t=LOTE_T)
            lotes_hoje += 1
            _lote_global_idx += 1

    # Agendar lotes parciais após os lotes cheios
    if volumes_parciais and produtividade_maq:
        for produto, vol_p in volumes_parciais.items():
            if vol_p <= 1e-6 or idx_dia >= len(dias_disponiveis):
                continue
            prod_rate = produtividade_maq.get(produto, 0)
            if prod_rate <= 0:
                continue
            setup_t_maq = PERDA_SETUP_T.get(maquina, PERDA_SETUP_T_DEFAULT) if isinstance(PERDA_SETUP_T, dict) else PERDA_SETUP_T
            hpl_parcial = (vol_p + setup_t_maq) / prod_rate
            _agendar_lote(produto, hpl_parcial, lote_num=1,
                          volume_t=vol_p, eh_parcial=True)

    return schedule


# ═══════════════════════════════════════════════════════
# SEÇÃO 9 — VALIDAÇÃO DIÁRIA DE BALANÇO
# ═══════════════════════════════════════════════════════

def validar_balanco_diario(schedule_maquinas, unit_name, restricoes, limites_diarios):
    """
    Recebe schedule_maquinas = {maquina: [lotes]}, onde cada lote tem
    'dia_inicio', 'dia_fim', 'produto', 'duracao_h'.
    """
    consumo_por_dia = defaultdict(lambda: defaultdict(float))
    max_por_maquina_dia = defaultdict(lambda: defaultdict(dict))

    for maq, lotes in schedule_maquinas.items():
        for lote in lotes:
            produto   = lote['produto']
            dia_i     = lote['dia_inicio']
            dia_f     = lote['dia_fim']
            vol_lote  = lote.get('volume_t', LOTE_T)

            span_dias = max(1, dia_f - dia_i + 1)
            frac = 1.0 / span_dias

            for dia in range(dia_i, dia_f + 1):
                for r in restricoes:
                    c_lote = consumo_lote_restricao(produto, maq, r, unit_name, vol=vol_lote)
                    if r in RESTRICOES_HORARIAS:
                        atual = max_por_maquina_dia[dia][r].get(maq, 0.0)
                        if c_lote > atual:
                            max_por_maquina_dia[dia][r][maq] = c_lote
                    else:
                        consumo_por_dia[dia][r] += c_lote * 365.0 * frac

    for dia, restr_dict in max_por_maquina_dia.items():
        for r, por_maquina in restr_dict.items():
            consumo_por_dia[dia][r] += sum(por_maquina.values())

    violacoes = []
    for dia, consumos in sorted(consumo_por_dia.items()):
        for r, usado in consumos.items():
            if unit_name == 'ORT':
                usado = usado + FIXED_ORT_PER_DAY.get(r, 0.0)
            lim = limites_diarios.get(r, 1e9)
            if lim >= 1e8:
                continue
            if usado > lim * 1.001:
                excesso_pct = 100.0 * (usado - lim) / lim
                violacoes.append({
                    'dia':         dia,
                    'restricao':   r,
                    'usado':       round(usado, 6),
                    'limite':      round(lim, 6),
                    'excesso_pct': round(excesso_pct, 2),
                })

    return violacoes


# ═══════════════════════════════════════════════════════
# SEÇÃO 10 — LOOP DE ENCADEAMENTO ANUAL
# ═══════════════════════════════════════════════════════

def gerar_ciclos_ano(janela=JANELA_MAX_DIAS, total_dias=365):
    """Gera lista de (dia_inicio, dia_fim) para cada ciclo."""
    ciclos = []
    d = 1
    while d <= total_dias:
        fim = min(d + janela - 1, total_dias)
        ciclos.append((d, fim))
        d = fim + 1
    return ciclos


def sequenciar_ano(unit_name, maquinas_unit, limites):
    ciclos_ano = gerar_ciclos_ano(janela=JANELA_MAX_DIAS, total_dias=DIAS_PERIODO)
    n_ciclos   = len(ciclos_ano)

    produzido_acumulado = {}
    resultados_ciclos = []
    stats_ano = {'n_deslocamentos': 0, 'n_hor_resolvidas_desl': 0, 'n_fallbacks_corte': 0,
                 'n_swaps': 0, 'n_hor_resolvidas_swap': 0}
    # Último produto produzido por máquina (para setup entre ciclos em ORT)
    _ultimo_produto_por_maquina = {}

    meta_anual_unit = {
        p: {m: v for m, v in maq_dict.items() if m in maquinas_unit}
        for p, maq_dict in meta_anual.items()
        if any(m in maquinas_unit for m in maq_dict)
    }

    for ciclo_idx, (dia_ini, dia_fim) in enumerate(ciclos_ano):
        num_ciclo    = ciclo_idx + 1
        dias_ciclo   = dia_fim - dia_ini + 1
        dias_acum    = dia_fim
        ultimo_ciclo = (num_ciclo == n_ciclos)

        print(f"\n  {unit_name} — Ciclo {num_ciclo}/{n_ciclos}  "
              f"(dias {dia_ini}-{dia_fim}, {dias_ciclo} dias)")

        saldo_ciclo = {}
        for p, maq_dict in meta_anual_unit.items():
            for m, meta in maq_dict.items():
                acum = produzido_acumulado.get(p, {}).get(m, 0.0)
                saldo = meta - acum
                if saldo > 1e-3:
                    saldo_ciclo.setdefault(p, {})[m] = saldo

        residuos_tratados = []
        if ultimo_ciclo:
            for p in list(saldo_ciclo.keys()):
                for m in list(saldo_ciclo.get(p, {}).keys()):
                    s = saldo_ciclo[p][m]
                    if 0 < s < LOTE_T:
                        residuos_tratados.append({
                            'produto': p, 'maquina': m,
                            'residuo_t': round(s, 2),
                            'opcao': 'A_lote_reduzido',
                        })

        (n_lotes_res, volume_parcial_res_ciclo, status, schedule_ciclo_pronto,
         pico_nao_resolvido, violacoes_pico, n_iteracoes_pico, stats_ciclo) = (
            solve_ciclo_com_pico_controlado(
                unit_name, maquinas_unit, limites,
                saldo_ciclo, meta_anual_unit,
                produzido_acumulado,
                dias_ciclo, dias_acum, num_ciclo,
                ultimo_produto_por_maquina=_ultimo_produto_por_maquina,
            )
        )
        for k in stats_ano:
            stats_ano[k] += stats_ciclo.get(k, 0)

        # Produção deste ciclo: lotes cheios + lote parcial por par
        prod_ciclo_pm = {}
        for p, maq_dict in n_lotes_res.items():
            for m, nl in maq_dict.items():
                prod_ciclo_pm[(p, m)] = nl * LOTE_T
        for (p, m), vp in volume_parcial_res_ciclo.items():
            prod_ciclo_pm[(p, m)] = prod_ciclo_pm.get((p, m), 0.0) + vp

        # Verificar ritmo mínimo (R4) — informativo, vale desde o ciclo 1
        ritmo_ok = True
        pares_ritmo_nok = []
        for p, maq_dict in meta_anual_unit.items():
            for m, meta in maq_dict.items():
                if meta <= 0:
                    continue
                if dict_prod_por_MP.get((m, p), 0) != 1:
                    continue
                if dict_produtividade_bruta.get((m, p), 0) <= 0:
                    continue
                acum = produzido_acumulado.get(p, {}).get(m, 0.0)
                prod_ciclo = prod_ciclo_pm.get((p, m), 0.0)
                alcancado  = acum + prod_ciclo
                minimo     = meta * (dias_acum / 365.0) * FATOR_TOLERANCIA_RITMO
                if alcancado < minimo - 1e-3:
                    ritmo_ok = False
                    pares_ritmo_nok.append({
                        'produto': p, 'maquina': m,
                        'alcancado': round(alcancado, 2),
                        'minimo_esperado': round(minimo, 2),
                        'deficit_t': round(minimo - alcancado, 2),
                    })

        schedule_ciclo = schedule_ciclo_pronto

        utilizacao_horas_ciclo = {}
        for maq in maquinas_unit:
            cap_dia = horas_dia.get(maq, 0.0)
            usado_por_dia = defaultdict(float)
            for lote in schedule_ciclo.get(maq, []):
                for dia_str, h in lote.get('horas_por_dia', {}).items():
                    usado_por_dia[dia_str] += h
            utilizacao_horas_ciclo[str(maq)] = {
                dia_str: {
                    'usado_h': round(usado, 3),
                    'capacidade_h': round(cap_dia, 3),
                    'pct': round(100.0 * usado / cap_dia, 1) if cap_dia > 1e-9 else 0.0,
                }
                for dia_str, usado in usado_por_dia.items()
            }

        violacoes = validar_balanco_diario(
            schedule_ciclo, unit_name,
            list(limites.keys()), limites,
        )

        # Acumular produção (lotes cheios + parciais)
        for (p, m), vol_prod in prod_ciclo_pm.items():
            produzido_acumulado.setdefault(p, {})[m] = (
                produzido_acumulado.get(p, {}).get(m, 0.0) + vol_prod
            )

        # Atualizar último produto por máquina (para setup entre ciclos em ORT)
        if unit_name == 'ORT':
            for maq, lotes in schedule_ciclo.items():
                if lotes:
                    _ultimo_produto_por_maquina[maq] = lotes[-1]['produto']

        total_lotes_ciclo = sum(
            nl for maq_dict in n_lotes_res.values() for nl in maq_dict.values())
        vol_parcial_ciclo = sum(volume_parcial_res_ciclo.values())

        print(f"    status={status}  lotes={total_lotes_ciclo}  "
              f"parcial={vol_parcial_ciclo:.0f}t  "
              f"violações_diárias={len(violacoes)}  "
              f"ritmo_ok={ritmo_ok}")

        resultados_ciclos.append({
            'ciclo':              num_ciclo,
            'dias':               [dia_ini, dia_fim],
            'dias_ciclo':         dias_ciclo,
            'status':             status,
            'ritmo_ok':           ritmo_ok,
            'pares_ritmo_nok':    pares_ritmo_nok,
            'n_lotes':            n_lotes_res,
            'volume_parcial':     {f"{p}|{m}": vp for (p, m), vp in volume_parcial_res_ciclo.items()},
            'schedule':           schedule_ciclo,
            'violacoes_diarias':  violacoes,
            'residuos_tratados':  residuos_tratados,
            'utilizacao_horas':   utilizacao_horas_ciclo,
            'pico_nao_resolvido': pico_nao_resolvido,
            'violacoes_pico':     violacoes_pico,
            'n_iteracoes_pico':   n_iteracoes_pico,
            'stats_deslocamento': stats_ciclo,
        })

    return resultados_ciclos, produzido_acumulado, stats_ano


# ═══════════════════════════════════════════════════════
# SEÇÃO 11 — EXECUÇÃO PRINCIPAL
# ═══════════════════════════════════════════════════════

print("\n" + "="*70)
print("SEQUENCIADOR POR LOTE ORT — Fase 1")
print(f"  LOTE_T={LOTE_T}t | PERDA_SETUP_T={PERDA_SETUP_T}t | "
      f"DIAS_PERIODO={DIAS_PERIODO}d | JANELA={JANELA_MAX_DIAS}d | "
      f"FATOR_RITMO={FATOR_TOLERANCIA_RITMO}")
print(f"  Máquinas: {MAQUINAS_ORT}")
print("="*70)

print("\n--- UNIDADE ORT ---")
ciclos_ORT, acum_ORT, stats_desl_ORT = sequenciar_ano('ORT', MAQUINAS_ORT, limites_ORT)

# ═══════════════════════════════════════════════════════
# SEÇÃO 12 — PRINTS NO TERMINAL
# ═══════════════════════════════════════════════════════

def _print_resumo_ciclos(ciclos, unit_name):
    print(f"\n{'='*70}")
    print(f"RESUMO CICLOS — {unit_name}")
    print(f"{'='*70}")
    print(f"  {'Ciclo':>5}  {'Dias':>10}  {'Status':>12}  "
          f"{'Lotes':>6}  {'Prod(t)':>10}  {'Viol.dia':>8}  {'Ritmo':>6}  {'Pico':>6}")
    print(f"  {'-'*75}")
    for c in ciclos:
        d = c['dias']
        nl = sum(n for md in c['n_lotes'].values() for n in md.values())
        viol = len(c['violacoes_diarias'])
        ritmo = 'OK' if c['ritmo_ok'] else 'NOK'
        pico = 'NOK' if c.get('pico_nao_resolvido') else 'OK'
        vol_par = sum(c.get('volume_parcial', {}).values())
        prod_t = nl * LOTE_T + vol_par
        print(f"  {c['ciclo']:>5}  {d[0]:>4}-{d[1]:<4}  {c['status']:>12}  "
              f"{nl:>6}  {prod_t:>10,.0f}  {viol:>8}  {ritmo:>6}  {pico:>6}")


def _print_producao_anual(acum, unit_name):
    print(f"\n{'='*70}")
    print(f"PRODUÇÃO ACUMULADA vs META ANUAL — {unit_name}")
    print(f"{'='*70}")
    print(f"  {'Produto':<18} {'MP':>4}  {'Produzido':>12}  {'Meta':>12}  "
          f"{'Cobert':>8}  {'Δt':>10}")
    print(f"  {'-'*65}")
    for p in sorted(acum.keys()):
        for m, prod in sorted(acum[p].items()):
            meta = meta_anual.get(p, {}).get(m, 0)
            delta = prod - meta
            cobert = 100.0 * prod / meta if meta > 0 else 0.0
            flag = '  !' if abs(delta) > LOTE_T else ''
            print(f"  {p:<18} {m:>4}  {prod:>12,.1f}  {meta:>12,.1f}  "
                  f"{cobert:>7.1f}%  {delta:>+10.1f}{flag}")


_print_resumo_ciclos(ciclos_ORT, 'ORT')
_print_producao_anual(acum_ORT, 'ORT')

# ═══════════════════════════════════════════════════════
# SEÇÃO 13 — FACTIBILIDADE ANUAL
# ═══════════════════════════════════════════════════════

ciclos_infeasiveis = [
    c for c in ciclos_ORT
    if c['status'] == 'infeasivel'
]
ciclos_ritmo_nok = [
    c for c in ciclos_ORT
    if not c['ritmo_ok']
]

ciclos_pico_nao_resolvido = [
    c for c in ciclos_ORT
    if c.get('pico_nao_resolvido')
]
pares_problematicos = []
for c in ciclos_ritmo_nok:
    pares_problematicos.extend(c['pares_ritmo_nok'])

pares_nao_sequenciaveis = []
for p, maq_dict in meta_anual.items():
    for m, meta in maq_dict.items():
        if meta <= 0:
            continue
        if m not in MAQUINAS_ORT:
            continue
        has_map  = dict_prod_por_MP.get((m, p), 0) == 1
        has_prod = dict_produtividade_bruta.get((m, p), 0) > 0
        if not has_map or not has_prod:
            pares_nao_sequenciaveis.append({
                'produto': p, 'maquina': m, 'meta_t': round(meta, 2),
                'sem_mapeamento_MP': not has_map,
                'sem_produtividade': not has_prod,
            })

factibilidade_anual = (len(ciclos_infeasiveis) == 0 and len(ciclos_ritmo_nok) == 0
                        and len(ciclos_pico_nao_resolvido) == 0)

print(f"\n{'='*70}")
print("FACTIBILIDADE ANUAL")
print(f"{'='*70}")
if factibilidade_anual:
    print("  RESULTADO: FACTÍVEL — o mix anual é fisicamente sequenciável "
          "dentro das janelas de 60 dias e restrições de balanço.")
else:
    print("  RESULTADO: INFACTÍVEL ou COM RESTRIÇÕES")
    if ciclos_infeasiveis:
        print(f"    Ciclos sem solução viável: "
              f"{[c['ciclo'] for c in ciclos_infeasiveis]}")
    if ciclos_ritmo_nok:
        print(f"    Ciclos com ritmo mínimo não atingido: "
              f"{[c['ciclo'] for c in ciclos_ritmo_nok]}")
    if ciclos_pico_nao_resolvido:
        print(f"    Ciclos com pico instantâneo não resolvido: "
              f"{[c['ciclo'] for c in ciclos_pico_nao_resolvido]}")
        for c in ciclos_pico_nao_resolvido[:5]:
            print(f"      Ciclo {c['ciclo']}: {c['violacoes_pico']}")
    if pares_problematicos:
        print(f"    Pares com déficit de ritmo:")
        for par in pares_problematicos[:10]:
            print(f"      {par['produto']} MP{par['maquina']}: "
                  f"alcançado={par['alcancado']:.0f}t "
                  f"mínimo={par['minimo_esperado']:.0f}t "
                  f"déficit={par['deficit_t']:.0f}t")

if pares_nao_sequenciaveis:
    meta_total_ns = sum(p['meta_t'] for p in pares_nao_sequenciaveis)
    print(f"\n  AVISO: {len(pares_nao_sequenciaveis)} pares têm meta no LP mas "
          f"não são sequenciáveis (sem mapeamento/produtividade): "
          f"{meta_total_ns:,.0f} t")
    for par in pares_nao_sequenciaveis[:5]:
        motivo = ('sem_mapeamento_MP' if par['sem_mapeamento_MP'] else 'sem_produtividade')
        print(f"    {par['produto']} MP{par['maquina']}: "
              f"meta={par['meta_t']:,.0f}t  [{motivo}]")

# ─────────────────────────────────────────────────────
# COMPARATIVO ANTES / DEPOIS DO DESLOCAMENTO TEMPORAL
# ─────────────────────────────────────────────────────
ANTES_TOTAL_VIOL   = 0   # baseline a calibrar para ORT isolado
ANTES_HOR_VIOL     = 0
ANTES_NAO_HOR_VIOL = 0
ANTES_COB_PCT      = 0.0

_all_viol = [v for c in ciclos_ORT for v in c['violacoes_diarias']]
depois_viol_ORT = sum(len(c['violacoes_diarias']) for c in ciclos_ORT)
depois_total    = depois_viol_ORT

hor_res_after = {r: 0 for r in RESTRICOES_HORARIAS}
nao_hor_after = 0
for v in _all_viol:
    if v['restricao'] in RESTRICOES_HORARIAS:
        hor_res_after[v['restricao']] = hor_res_after.get(v['restricao'], 0) + 1
    else:
        nao_hor_after += 1
depois_hor_total = sum(hor_res_after.values())

total_prod_depois, total_meta_depois = 0.0, 0.0
for p, maq_dict in acum_ORT.items():
    for m, prod in maq_dict.items():
        total_prod_depois += prod
        total_meta_depois += meta_anual.get(p, {}).get(m, 0)
depois_cob = 100.0 * total_prod_depois / total_meta_depois if total_meta_depois > 0 else 0

desl_ort     = stats_desl_ORT['n_deslocamentos']
hor_res_ort  = stats_desl_ORT['n_hor_resolvidas_desl']
fb_ort       = stats_desl_ORT['n_fallbacks_corte']
swap_ort     = stats_desl_ORT.get('n_swaps', 0)
hor_swap_ort = stats_desl_ORT.get('n_hor_resolvidas_swap', 0)

# Validação de integridade de volumes ORT
_vol_schedule_ORT: dict = {}
for c in ciclos_ORT:
    for maq, lotes in c['schedule'].items():
        maq_int = int(maq) if isinstance(maq, str) else maq
        for lote in lotes:
            chave = (lote['produto'], maq_int)
            _vol_schedule_ORT[chave] = _vol_schedule_ORT.get(chave, 0.0) + lote.get('volume_t', LOTE_T)
_vol_ok_ORT = all(
    abs(_vol_schedule_ORT.get((p, m), 0.0) - v) < 10.0
    for p, md in acum_ORT.items() for m, v in md.items()
)
if not _vol_ok_ORT:
    print("\n  Divergências schedule vs acumulado (ORT):")
    for p, md in acum_ORT.items():
        for m, v in md.items():
            sched = _vol_schedule_ORT.get((p, m), 0.0)
            diff = sched - v
            if abs(diff) >= 10.0:
                print(f"    {p} MP{m}: schedule={sched:,.1f}t  acum={v:,.1f}t  "
                      f"Δ={diff:+,.1f}t")

print(f"\n{'='*70}")
print("COMPARATIVO ANTES / DEPOIS — 3 CAMADAS DE CORREÇÃO")
print(f"  Baseline (SEM nenhuma camada): {ANTES_TOTAL_VIOL} viol. totais")
print(f"{'='*70}")
print(f"  {'Métrica':<42}  {'ANTES':>8}  {'DEPOIS':>8}  {'Delta':>8}")
print(f"  {'-'*70}")
print(f"  {'Violações diárias totais':<42}  {ANTES_TOTAL_VIOL:>8}  "
      f"{depois_total:>8}  {depois_total - ANTES_TOTAL_VIOL:>+8}")
print(f"  {'  — horárias (Evap+LicVerde+Outorga)':<42}  {ANTES_HOR_VIOL:>8}  "
      f"{depois_hor_total:>8}  {depois_hor_total - ANTES_HOR_VIOL:>+8}")
print(f"  {'  — não-horárias (PMAD/Kamyr/etc)':<42}  {ANTES_NAO_HOR_VIOL:>8}  "
      f"{nao_hor_after:>8}  {nao_hor_after - ANTES_NAO_HOR_VIOL:>+8}")
print(f"  {'Cobertura meta anual (%)':<42}  {ANTES_COB_PCT:>8.1f}  "
      f"{depois_cob:>8.1f}  {depois_cob - ANTES_COB_PCT:>+8.1f}")
print(f"\n  Camada 1 — Swap de composição:")
print(f"    ORT: {swap_ort} swaps aplicados, "
      f"{hor_swap_ort} viol.horárias resolvidas por swap")
print(f"\n  Camada 2 — Deslocamento temporal:")
print(f"    ORT: {desl_ort} lotes deslocados, "
      f"{hor_res_ort} viol.horárias resolvidas, {fb_ort} fallbacks p/ corte")
print(f"\n  Detalhe viol.horárias residuais (DEPOIS):")
for r in sorted(RESTRICOES_HORARIAS):
    n = hor_res_after.get(r, 0)
    if n > 0:
        print(f"    {r:<30}: {n}")
print(f"\n  Validação de integridade de volumes (swap não altera metas):")
print(f"    Schedule ORT consistente com acumulados: {'OK' if _vol_ok_ORT else 'FALHA'}")
print(f"{'='*70}")

# ═══════════════════════════════════════════════════════
# SEÇÃO 14 — EXPORTAÇÃO DO JSON
# ═══════════════════════════════════════════════════════

def _serializar_schedule(ciclos):
    """Converte schedule de cada ciclo para formato JSON-serializável."""
    out = []
    for c in ciclos:
        sc_json = {}
        for maq, lotes in c['schedule'].items():
            sc_json[str(maq)] = lotes
        out.append({
            'ciclo':            c['ciclo'],
            'dias':             c['dias'],
            'dias_ciclo':       c['dias_ciclo'],
            'status':           c['status'],
            'ritmo_ok':         c['ritmo_ok'],
            'pares_ritmo_nok':  c['pares_ritmo_nok'],
            'n_lotes': {
                p: {str(m): nl for m, nl in md.items()}
                for p, md in c['n_lotes'].items()
            },
            'schedule':          sc_json,
            'volume_parcial':     c.get('volume_parcial', {}),
            'violacoes_diarias':  c['violacoes_diarias'],
            'residuos_tratados':  c['residuos_tratados'],
            'pico_nao_resolvido': c.get('pico_nao_resolvido', False),
            'violacoes_pico':     c.get('violacoes_pico', {}),
            'n_iteracoes_pico':   c.get('n_iteracoes_pico', 1),
        })
    return out


schedule_anual_ORT = {str(m): [] for m in MAQUINAS_ORT}

for c in ciclos_ORT:
    for maq, lotes in c['schedule'].items():
        schedule_anual_ORT[str(maq)].extend(lotes)

diag_ORT = [v for c in ciclos_ORT for v in c['violacoes_diarias']]

total_lotes   = sum(nl for md in acum_ORT.values() for nl in md.values()) / LOTE_T
_setup_medio  = sum(PERDA_SETUP_T.values()) / len(PERDA_SETUP_T) if isinstance(PERDA_SETUP_T, dict) else PERDA_SETUP_T
total_setup_t = total_lotes * _setup_medio

seq_data = {
    'meta_dados': {
        'lote_t':                    LOTE_T,
        'perda_setup_t':             PERDA_SETUP_T,
        'perda_setup_default':       PERDA_SETUP_T_DEFAULT,
        'dias_periodo':              DIAS_PERIODO,
        'janela_max_dias':           JANELA_MAX_DIAS,
        'fator_tolerancia_ritmo':    FATOR_TOLERANCIA_RITMO,
        'fator_spread':              FATOR_SPREAD,
        'dispersao_uniforme_ort':    DISPERSAO_UNIFORME_ORT,
        'fator_dispersao_ort':       FATOR_DISPERSAO_ORT,
        'maquinas_ORT':              MAQUINAS_ORT,
        'n_ciclos_ORT':              len(ciclos_ORT),
    },
    'ciclos': {
        'ORT': _serializar_schedule(ciclos_ORT),
    },
    'schedule': {
        'ORT': schedule_anual_ORT,
    },
    'resumo_producao': {
        'ORT': {
            p: {str(m): {'produzido_t': round(v, 2),
                         'meta_t':      round(meta_anual.get(p, {}).get(m, 0), 2),
                         'cobertura_pct': round(
                             100 * v / meta_anual.get(p, {}).get(m, 1)
                             if meta_anual.get(p, {}).get(m, 0) > 0 else 0, 1)}
                for m, v in maq_dict.items()}
            for p, maq_dict in acum_ORT.items()
        },
    },
    'limites': {
        'ORT': {k: v for k, v in limites_ORT.items() if v < 1e8},
    },
    'diagnostico': {
        'ORT': diag_ORT,
        'ciclos_infeasiveis':        [c['ciclo'] for c in ciclos_infeasiveis],
        'ciclos_ritmo_nok':          [c['ciclo'] for c in ciclos_ritmo_nok],
        'ciclos_pico_nao_resolvido': [c['ciclo'] for c in ciclos_pico_nao_resolvido],
        'pares_nao_sequenciaveis':   pares_nao_sequenciaveis,
    },
    'resumo_geral': {
        'factibilidade_anual':     factibilidade_anual,
        'pares_problematicos':     pares_problematicos[:50],
        'n_pares_nao_seq':         len(pares_nao_sequenciaveis),
        'meta_t_nao_seq':          round(sum(p['meta_t'] for p in pares_nao_sequenciaveis), 1),
        'total_lotes_aprox':       round(total_lotes, 0),
        'total_perda_setup_t':     round(total_setup_t, 1),
        'n_violacoes_diarias_ORT': len(diag_ORT),
    },
}

with open('sequenciamento_ort30dias_data.json', 'w', encoding='utf-8') as f:
    json.dump(seq_data, f, ensure_ascii=False, indent=2)

print(f"\n✓ sequenciamento_ort30dias_data.json gerado")

# ═══════════════════════════════════════════════════════
# SEÇÃO 15 — DADOS PARA O DASHBOARD V2 + MATRIZ DE FEASIBILITY
# ═══════════════════════════════════════════════════════

GRID_FEASIBILITY = [
    {'lote_t': 250, 'perda_setup_t': 20},
    {'lote_t': 300, 'perda_setup_t': 20},
    {'lote_t': 300, 'perda_setup_t': 30},
    {'lote_t': 400, 'perda_setup_t': 20},
    {'lote_t': 400, 'perda_setup_t': 30},
    {'lote_t': 400, 'perda_setup_t': 40},
    {'lote_t': 500, 'perda_setup_t': 30},
    {'lote_t': 500, 'perda_setup_t': 40},
    {'lote_t': 500, 'perda_setup_t': 50},
    {'lote_t': 600, 'perda_setup_t': 40},
    {'lote_t': 600, 'perda_setup_t': 50},
    {'lote_t': 700, 'perda_setup_t': 50},
]

HORIZONTE_FEASIBILITY_DIAS = 60

print("\n" + "="*70)
print("SEÇÃO 15 — MATRIZ DE FEASIBILITY (lote_t × perda_setup_t)")
print(f"  Horizonte por combinação: {HORIZONTE_FEASIBILITY_DIAS} dias  |  "
      f"Combinações: {len(GRID_FEASIBILITY)}")
print("="*70)


def _testar_combinacao_unidade(unit_name, maquinas_unit, limites, lote_t, perda_setup_t,
                                horizonte_dias):
    global LOTE_T, PERDA_SETUP_T
    lote_t_orig, perda_setup_t_orig = LOTE_T, PERDA_SETUP_T
    LOTE_T, PERDA_SETUP_T = float(lote_t), float(perda_setup_t)

    try:
        meta_anual_unit = {
            p: {m: v for m, v in maq_dict.items() if m in maquinas_unit}
            for p, maq_dict in meta_anual.items()
            if any(m in maquinas_unit for m in maq_dict)
        }
        saldo_ciclo = {
            p: dict(maq_dict) for p, maq_dict in meta_anual_unit.items() if maq_dict
        }

        if not saldo_ciclo:
            return {
                'factivel': None, 'status': 'sem_metas',
                'cobertura_pct': 0.0, 'pct_subciclos_ok': 0.0,
            }

        n_lotes_res, volume_parcial_res, status, _deficits_r4 = solve_ciclo(
            unit_name, maquinas_unit, limites,
            saldo_ciclo, meta_anual_unit,
            produzido_acumulado={},
            dias_ciclo=horizonte_dias,
            dias_acumulados=horizonte_dias,
            num_ciclo=1,
        )

        produzido_total = 0.0
        meta_periodo_total = 0.0
        for p, maq_dict in meta_anual_unit.items():
            for m, meta in maq_dict.items():
                meta_periodo = meta * (horizonte_dias / 365.0)
                meta_periodo_total += meta_periodo
                produzido = (n_lotes_res.get(p, {}).get(m, 0) * LOTE_T
                             + volume_parcial_res.get((p, m), 0.0))
                produzido_total += produzido

        cobertura_pct = (100.0 * produzido_total / meta_periodo_total
                          if meta_periodo_total > 1e-6 else 0.0)

        factivel = status in ('otimo', 'subotimo')
        pct_subciclos_ok = 100.0 if factivel else 0.0

        return {
            'factivel': factivel,
            'status': status,
            'cobertura_pct': round(cobertura_pct, 1),
            'pct_subciclos_ok': round(pct_subciclos_ok, 1),
        }
    finally:
        LOTE_T, PERDA_SETUP_T = lote_t_orig, perda_setup_t_orig


def gerar_matriz_feasibility():
    resultados = []
    for combo in GRID_FEASIBILITY:
        lote_t = combo['lote_t']
        perda  = combo['perda_setup_t']
        print(f"\n  Testando lote_t={lote_t}  perda_setup_t={perda} ...")

        res_ort = _testar_combinacao_unidade(
            'ORT', MAQUINAS_ORT, limites_ORT, lote_t, perda, HORIZONTE_FEASIBILITY_DIAS)

        factivel_geral      = bool(res_ort.get('factivel'))
        cobertura_geral     = res_ort.get('cobertura_pct', 0.0)
        pct_subciclos_geral = res_ort.get('pct_subciclos_ok', 0.0)

        print(f"    ORT: status={res_ort.get('status')}  cobertura={res_ort.get('cobertura_pct')}%")
        print(f"    => geral: factivel={factivel_geral}  cobertura={cobertura_geral}%")

        resultados.append({
            'lote_t':           lote_t,
            'perda_setup_t':    perda,
            'factivel':         factivel_geral,
            'cobertura_pct':    cobertura_geral,
            'pct_subciclos_ok': pct_subciclos_geral,
            'detalhe': {
                'ORT': res_ort,
            },
        })

    return resultados


matriz_feasibility = gerar_matriz_feasibility()


def calcular_utilizacao_diaria_completa(ciclos, unit_name, restricoes, limites_diarios):
    consumo_por_dia = defaultdict(lambda: defaultdict(float))
    max_por_maquina_dia = defaultdict(lambda: defaultdict(dict))
    max_dia = 0

    for c in ciclos:
        schedule_ciclo = c['schedule']
        for maq, lotes in schedule_ciclo.items():
            for lote in lotes:
                produto   = lote['produto']
                dia_i     = lote['dia_inicio']
                dia_f     = lote['dia_fim']
                max_dia   = max(max_dia, dia_f)
                span_dias = max(1, dia_f - dia_i + 1)
                frac = 1.0 / span_dias
                vol_lote = lote.get('volume_t', LOTE_T)
                maq_int = int(maq) if isinstance(maq, str) else maq
                for dia in range(dia_i, dia_f + 1):
                    for r in restricoes:
                        c_lote = consumo_lote_restricao(produto, maq_int, r, unit_name, vol=vol_lote)
                        if r in RESTRICOES_HORARIAS:
                            atual = max_por_maquina_dia[dia][r].get(maq_int, 0.0)
                            if c_lote > atual:
                                max_por_maquina_dia[dia][r][maq_int] = c_lote
                        else:
                            consumo_por_dia[dia][r] += c_lote * 365.0 * frac

    for dia, restr_dict in max_por_maquina_dia.items():
        for r, por_maquina in restr_dict.items():
            consumo_por_dia[dia][r] += sum(por_maquina.values())

    serie = []
    for dia in range(1, max_dia + 1):
        item = {'dia': dia}
        for r in restricoes:
            usado = consumo_por_dia.get(dia, {}).get(r, 0.0)
            if unit_name == 'ORT':
                usado += FIXED_ORT_PER_DAY.get(r, 0.0)
            lim   = limites_diarios.get(r, 1e9)
            pct   = (100.0 * usado / lim) if lim > 1e-9 and lim < 1e8 else 0.0
            item[r] = {'usado': round(usado, 6),
                       'limite': round(lim, 6) if lim < 1e8 else None,
                       'pct': round(pct, 1)}
        serie.append(item)
    return serie


utilizacao_diaria_ORT = calcular_utilizacao_diaria_completa(
    ciclos_ORT, 'ORT', list(limites_ORT.keys()), limites_ORT)


def consolidar_utilizacao_horas(ciclos):
    consolidado = {}
    for c in ciclos:
        for maq_str, dias_dict in c.get('utilizacao_horas', {}).items():
            consolidado.setdefault(maq_str, {})
            consolidado[maq_str].update(dias_dict)
    return consolidado


utilizacao_horas_ORT = consolidar_utilizacao_horas(ciclos_ORT)


def calcular_metricas_produto_maquina(ciclos, maquinas_unit):
    metricas = defaultdict(lambda: defaultdict(lambda: {
        'horas_tot': 0.0, 'n_setups': 0, 'dias_prod': set(),
    }))
    for c in ciclos:
        for maq, lotes in c['schedule'].items():
            for lote in lotes:
                p = lote['produto']
                m = str(maq)
                metricas[p][m]['horas_tot'] += lote.get('duracao_h', 0.0)
                metricas[p][m]['n_setups']  += 1
                for d in range(lote['dia_inicio'], lote['dia_fim'] + 1):
                    metricas[p][m]['dias_prod'].add(d)

    out = {}
    for p, maq_dict in metricas.items():
        out[p] = {}
        for m, v in maq_dict.items():
            out[p][m] = {
                'horas_tot':  round(v['horas_tot'], 2),
                'n_setups':   v['n_setups'],
                'dias_prod':  len(v['dias_prod']),
            }
    return out


metricas_ORT = calcular_metricas_produto_maquina(ciclos_ORT, MAQUINAS_ORT)


def _mesclar_resumo_com_metricas(resumo_producao_unit, metricas_unit):
    out = {}
    for p, maq_dict in resumo_producao_unit.items():
        out[p] = {}
        for m, v in maq_dict.items():
            extra = metricas_unit.get(p, {}).get(
                m, {'horas_tot': 0.0, 'n_setups': 0, 'dias_prod': 0})
            out[p][m] = {**v, **extra,
                         'folga_t': round(v.get('meta_t', 0) - v.get('produzido_t', 0), 2)}
    return out


resumo_producao_v2_ORT = _mesclar_resumo_com_metricas(
    seq_data['resumo_producao']['ORT'], metricas_ORT)


def enriquecer_schedule_com_toneladas(ciclos, unit_name):
    schedule_enriquecido = {}
    for c in ciclos:
        num_ciclo = c['ciclo']
        for maq, lotes in c['schedule'].items():
            schedule_enriquecido.setdefault(maq, [])
            produto_anterior = None
            for lote in lotes:
                produto = lote['produto']
                maq_int = int(maq)
                produtividade = dict_produtividade_bruta.get((maq_int, produto), 0)
                duracao_h = lote.get('duracao_h', 0.0)
                _setup_ref = PERDA_SETUP_T.get(maq_int, PERDA_SETUP_T_DEFAULT) if isinstance(PERDA_SETUP_T, dict) else PERDA_SETUP_T
                toneladas_brutas = (duracao_h * produtividade - _setup_ref
                                    if produtividade > 0 else 0.0)

                # Setup real: depende da transição para ORT
                if unit_name == 'ORT' and maq_int in MAQUINAS_ORT:
                    setup_t = get_setup_t(maq_int, produto_anterior, produto)
                else:
                    setup_t = 0.0 if produto_anterior == produto else _setup_ref
                # Primeiro lote do ciclo: setup zero (sem produto anterior)
                if produto_anterior is None:
                    setup_t = 0.0

                lote_enriquecido = {
                    **lote,
                    'toneladas_brutas':  round(max(toneladas_brutas, 0.0), 2),
                    'semana':            ((lote['dia_inicio'] - 1) // 7) + 1,
                    'ciclo':             num_ciclo,
                    'setup_t':           round(setup_t, 1),
                    'produto_anterior':  produto_anterior,
                }
                schedule_enriquecido[maq].append(lote_enriquecido)
                produto_anterior = produto
    return schedule_enriquecido


schedule_v2_ORT = enriquecer_schedule_com_toneladas(ciclos_ORT, 'ORT')

seq_data['utilizacao_restricoes'] = {
    'ORT': utilizacao_diaria_ORT,
}

seq_data['utilizacao_horas_maquina'] = {
    'ORT': utilizacao_horas_ORT,
}

seq_data['resumo_producao_v2'] = {
    'ORT': resumo_producao_v2_ORT,
}

seq_data['schedule_v2'] = {
    'ORT': schedule_v2_ORT,
}

seq_data['feasibility_matrix'] = {
    'horizonte_dias_testado': HORIZONTE_FEASIBILITY_DIAS,
    'grid_parametros':        GRID_FEASIBILITY,
    'resultados':             matriz_feasibility,
}

seq_data['meta_dados']['total_dias_ano'] = max(
    [c['dias'][1] for c in ciclos_ORT] + [0]
) or 365
seq_data['meta_dados']['n_ciclos_max'] = len(ciclos_ORT)

with open('sequenciamento_ort30dias_data.json', 'w', encoding='utf-8') as f:
    json.dump(seq_data, f, ensure_ascii=False, indent=2)

print(f"\n✓ sequenciamento_ort30dias_data.json REGRAVADO com campos da Seção 15:")
print(f"    - utilizacao_restricoes (série diária completa, ORT)")
print(f"    - resumo_producao_v2 (com horas/setups/dias por produto-máquina)")
print(f"    - schedule_v2 (com toneladas_brutas, semana, ciclo por lote)")
print(f"    - feasibility_matrix ({len(GRID_FEASIBILITY)} combinações testadas, "
      f"horizonte={HORIZONTE_FEASIBILITY_DIAS}d)")