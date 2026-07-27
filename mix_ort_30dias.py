import pyomo.environ as pyo
from pyomo.opt import SolverFactory
import pandas as pd
from collections import defaultdict
from typing import Dict,List,Tuple,Any
import sys
import os


def get_glpk_path() -> str:
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, 'solvers', 'glpsol.exe')
    return 'glpsol'

# ══════════════════════════════════════════════════════════════════════
# BLOCO DE CACHE — aplicado quando chamado via app.py (Flask)
# Quando rodado standalone (PyCharm), o arquivo não existe e este
# bloco não faz nada.
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

_mix_params = _cache_get("mix_params") or {}
DIAS_PERIODO     = _mix_params.get("DIAS_PERIODO",     7)
BASE_DEMANDA_DIAS = _mix_params.get("BASE_DEMANDA_DIAS", 365)

# Lista de produtos
data1 = pd.read_excel('Simulador - Mix de cartões_3__2_.xlsx', sheet_name='Produtividade', skiprows=2)
data2 = pd.read_excel('Simulador - Mix de cartões_3__2_.xlsx', sheet_name='Total waste', skiprows=3)
data3 = pd.read_excel('Simulador - Mix de cartões_3__2_.xlsx', sheet_name='Desclassificados', skiprows=2)
data4 = pd.read_excel('Simulador - Mix de cartões_3__2_.xlsx', sheet_name='Reprocesso', skiprows=2)
data5 = pd.read_excel('Simulador - Mix de cartões_3__2_.xlsx', sheet_name='Demanda', skiprows=2)
data6 = pd.read_excel('Simulador - Mix de cartões_3__2_.xlsx', sheet_name='Produto por MP', skiprows=2)
data7 = pd.read_excel('Simulador - Mix de cartões_3__2_.xlsx', sheet_name='Custos', skiprows=2)

#Processo para fazer o ajuste de produtos entre a aba demanda e a aba reprocesso
demanda_ajustada = pd.DataFrame(data5)
demanda_ajustada['Produto Ajustado'] = demanda_ajustada['Produto']
reprocesso_df = pd.DataFrame(data4)
reprocesso_df = reprocesso_df[['Produto vendável','Produto base']]
dict_reprocesso = reprocesso_df.set_index('Produto vendável')['Produto base'].to_dict()
demanda_ajustada['Produto Ajustado'] = demanda_ajustada['Produto Ajustado'].apply(lambda x: dict_reprocesso.get(x,x))

# Converter em sets (removendo valores NaN)
set_produtos_data1 = set(data1['Produto'].dropna())
set_produtos_data2 = set(data2['Produto'].dropna())
set_produtos_data3 = set(data3['Produto original'].dropna())
set_produtos_data32 = set(data3['Desclassificado'].dropna())
#set_produtos_data4 = set(data4['Produto vendável'].dropna()) # Apagar
#set_produtos_data42 = set(data4['Produto base'].dropna()) # Apagar
set_produtos_data5 = set(demanda_ajustada['Produto Ajustado'].dropna()) # Lista de demanda ajustada
set_produtos_data6 = set(data6['Produto'].dropna())
set_produtos_data7 = set(data7['Produto'].dropna())

# União de todos os sets para obter lista completa de produtos únicos
lista_produtos = set_produtos_data1.union(
    set_produtos_data2,
    set_produtos_data3,
    set_produtos_data32,
#    set_produtos_data4,   #Apagar
#    set_produtos_data42,  # Apagar
    set_produtos_data5,   # Ajustado anteriormente
    set_produtos_data6,
    set_produtos_data7
)

#Carregar Produtividade, Taxa de Performance e Produtividade Bruta
data = pd.read_excel('Simulador - Mix de cartões_3__2_.xlsx',sheet_name='Produtividade',skiprows=2)
produtividade_df = pd.DataFrame(data)
produtividade = produtividade_df[['Produto','Máquina','Produtividade máxima (t/h)','Taxa PERF (Meta)']]
taxa_qual_por_maquina = {
    27: 1.0,   # MP27
    28: 1.0,   # MP28
    25: 1.0,   # MC25
    26: 1.0,   # MC26
}
produtividade['Taxa QUAL'] = produtividade['Máquina'].map(taxa_qual_por_maquina) # Adicionei a coluna Taxa QUAL para replicar a Taxa QUAL da

dict_produtividade = {}
dict_taxa_performance = {}
dict_taxa_qual = {}
dict_produtividade_bruta = {}

maquinas_ort = {27, 28, 25, 26}

for index,row in produtividade.iterrows():
    produto = row['Produto']
    maquina = int(row['Máquina'])
    if maquina not in maquinas_ort:
        continue
    produtividade_max = float(row['Produtividade máxima (t/h)'])
    taxa_performance = float(row['Taxa PERF (Meta)'])
    taxa_qual = float(row['Taxa QUAL'])
    dict_produtividade[(maquina,produto)] = produtividade_max
    dict_taxa_performance[(maquina,produto)] = taxa_performance
    dict_taxa_qual[(maquina,produto)] = taxa_qual
    #dict_produtividade_bruta[(maquina,produto)] = taxa_qual*produtividade_max
    dict_produtividade_bruta[(maquina, produto)] = taxa_qual * produtividade_max * taxa_performance

#Carregar Total Waste - Condicional Flag_CBA:

param_refugo_cba = {'MA':False,'OR':False}
maquina_unidade = {27: 'ORT', 28: 'ORT', 25: 'ORT', 26: 'ORT'}

# 1. Preparar o DataFrame 'Total Waste' para modificação
total_waste_df = pd.DataFrame(data2) # Usando data2 que já foi lido
total_waste = total_waste_df[['Produto','Máquina','IAC','Refugo MP','Rep Externo - Perdas','MR2 - Perdas','Sala - Perdas','Cortadeira - Perdas Gramatura','Cortadeira - Perdas Cortadeira','Estoque - Perdas Expedição','Estoque - Perdas Refugo']].copy()

# Criar a coluna 'Refugo_MP_CBA' inicializada com 'Refugo MP' original
total_waste['Refugo_MP_CBA'] = total_waste['Refugo MP']

# 2. Preparar o DataFrame 'Desclassificados' para modificação
desclassificados_df = pd.DataFrame(data3)
desclassificados_df = desclassificados_df[['Produto original','Desclassificado','Máquina','Taxa']].copy()
# Adiciona a coluna 'Unidade' para facilitar a verificação da flag
desclassificados_df['Unidade'] = desclassificados_df['Máquina'].map(maquina_unidade)

# 3. Aplicar a lógica de ajuste do "Refugo MP" e zerar taxas de CBA
for index, row in desclassificados_df.iterrows():
    produto_original = row['Produto original']
    produto_desclassificado = row['Desclassificado']
    maquina = int(row['Máquina'])
    if maquina not in maquinas_ort:
        continue
    unidade = row['Unidade']
    taxa_desclassificacao_cba = float(row['Taxa'])  # Assumindo que 'Taxa' já é float ou pode ser convertida

    # Condições para aplicar a regra do CBA como refugo:
    # - O produto desclassificado começa com "CBA"
    # - O produto original está na lista de produtos a serem otimizados
    # - A unidade da máquina tem a flag 'param_refugo_cba' como True
    if (str(produto_desclassificado).startswith("CBA") and
            produto_original in lista_produtos and
            param_refugo_cba.get(unidade, False) == True):

        # Encontra a linha correspondente no total_waste (para obter o Refugo MP original)
        mask_target_total_waste = (total_waste['Máquina'] == maquina) & \
                                  (total_waste['Produto'] == produto_original)

        if not total_waste.loc[mask_target_total_waste].empty:
            taxa_refugo_mp_atual = float(total_waste.loc[mask_target_total_waste, 'Refugo MP'].values[0])

            # Recalcula o Refugo MP ajustado multiplicativamente
            taxa_refugo_mp_cba_atualizada = 1 - (1 - taxa_refugo_mp_atual) * (1 - taxa_desclassificacao_cba)

            # Atualiza a coluna 'Refugo_MP_CBA' no total_waste
            total_waste.loc[mask_target_total_waste, 'Refugo_MP_CBA'] = taxa_refugo_mp_cba_atualizada

            # Zera a taxa de desclassificação NO desclassificados_df para evitar dupla contagem.
            desclassificados_df.loc[index, 'Taxa'] = 0.0

#Carregar Total Waste (IAC, Refugo MP, Rep Externo - Perdas, MR2 - Perdas, Sala - Perdas, Cortadeira - Perdas Gramatura, Cortadeira - Perdas Cortadeira, Estoque - Perdas Expedição, Estoque - Perdas Refugo)

dict_IAC = {}
dict_Refugo_MP = {}
dict_Rep_Externo = {}
dict_MR2 = {}
dict_Sala_Perdas = {}
dict_Cortadeira_Perda_Gramatura = {}
dict_Cortadeira_Perdas_Cortadeira = {}
dict_Estoque_Perdas = {}
dict_Estoque_Perdas_Refugo = {}

for index,row in total_waste.iterrows():
    produto = row['Produto']
    maquina = int(row['Máquina'])
    if maquina not in maquinas_ort:
        continue
    IAC = float(row['IAC'])
    Refugo_MP = float(row['Refugo_MP_CBA']) #Está pegando da coluna atualizada considerando se CBA é refugo ou não
    Rep_Externo = float(row['Rep Externo - Perdas'])
    MR2 = float(row['MR2 - Perdas'])
    Sala_Perdas = float(row['Sala - Perdas'])
    Cortadeira_Perda_Gramatura = float(row['Cortadeira - Perdas Gramatura'])
    Cortadeira_Perdas_Cortadeira = float(row['Cortadeira - Perdas Cortadeira'])
    Estoque_Perdas = float(row['Estoque - Perdas Expedição'])
    Estoque_Perdas_Refugo = float(row['Estoque - Perdas Refugo'])

    dict_IAC[(maquina, produto)] = IAC
    dict_Refugo_MP[(maquina, produto)] = Refugo_MP
    dict_Rep_Externo[(maquina, produto)] = Rep_Externo
    dict_MR2[(maquina, produto)] = MR2
    dict_Sala_Perdas[(maquina, produto)] = Sala_Perdas
    dict_Cortadeira_Perda_Gramatura[(maquina, produto)] = Cortadeira_Perda_Gramatura
    dict_Cortadeira_Perdas_Cortadeira[(maquina, produto)] = Cortadeira_Perdas_Cortadeira
    dict_Estoque_Perdas[(maquina, produto)] = Estoque_Perdas
    dict_Estoque_Perdas_Refugo[(maquina, produto)] = Estoque_Perdas_Refugo

#Considerar as perdas por IAC e Refugo
dict_refugo_ajustado = {}
for (maquina,produto) in dict_Refugo_MP.keys():
    refugo_ajustado = 1 - (1 - dict_Refugo_MP[(maquina, produto)]) * dict_IAC.get((maquina, produto), 1)
    dict_refugo_ajustado[(maquina, produto)] = refugo_ajustado

#Considerar as perdas pelos demais fatores da aba Total Waste
"""
dict_waste = {}
for (maquina, produto) in dict_Refugo_MP.keys():
    waste_total = (
        dict_Rep_Externo.get((maquina, produto), 0) +
        dict_MR2.get((maquina, produto), 0) +
        dict_Sala_Perdas.get((maquina, produto), 0) +
        dict_Cortadeira_Perda_Gramatura.get((maquina, produto), 0) +
        dict_Cortadeira_Perdas_Cortadeira.get((maquina, produto), 0) +
        dict_Estoque_Perdas.get((maquina, produto), 0) +
        dict_Estoque_Perdas_Refugo.get((maquina, produto), 0)
    )
    dict_waste[(maquina, produto)] = waste_total
"""
dict_waste = {}
for (maquina, produto) in dict_Refugo_MP.keys():
    waste_total = 1 - (
        (1 - dict_Rep_Externo.get((maquina, produto), 0)) *
        (1 - dict_MR2.get((maquina, produto), 0)) *
        (1 - dict_Sala_Perdas.get((maquina, produto), 0)) *
        (1 - dict_Cortadeira_Perda_Gramatura.get((maquina, produto), 0)) *
        (1 - dict_Cortadeira_Perdas_Cortadeira.get((maquina, produto), 0)) *
        (1 - dict_Estoque_Perdas.get((maquina, produto), 0)) *
        (1 - dict_Estoque_Perdas_Refugo.get((maquina, produto), 0))
    )
    dict_waste[(maquina, produto)] = waste_total

#Fazer a combinação de IAC e Refugo com as demais perdas
dict_total_waste = {}
for (maquina, produto) in dict_Refugo_MP.keys():
    refugo_ajustado = dict_refugo_ajustado.get((maquina, produto), 0)
    waste = dict_waste.get((maquina, produto), 0)
    total_waste = 1 - (1 - refugo_ajustado) * (1 - waste)
    dict_total_waste[(maquina, produto)] = total_waste

#Carregar Desclassificados (Produto original, Desclassificado, Máquina e Taxa)

desclassificados_para_dict = desclassificados_df[['Produto original','Desclassificado','Máquina','Taxa']]
soma_taxas = desclassificados_para_dict.groupby(['Máquina','Produto original'])['Taxa'].sum().reset_index()
soma_taxas['Desclassificado'] = soma_taxas['Produto original']
soma_taxas['Taxa'] = 1 - soma_taxas['Taxa']
desclassificados_complementar = desclassificados_para_dict[['Produto original','Desclassificado','Máquina','Taxa']].copy()
desclassificados_complementar = pd.concat([desclassificados_complementar, soma_taxas[['Máquina', 'Produto original', 'Desclassificado', 'Taxa']]])
dict_taxa_desclassificado = desclassificados_complementar.set_index(['Máquina', 'Produto original', 'Desclassificado'])['Taxa'].to_dict()

# Produtos sem nenhuma desclassificação cadastrada — taxa complementar = 1.0
for (m, p) in dict_produtividade_bruta.keys():
    if (m, p, p) not in dict_taxa_desclassificado:
        dict_taxa_desclassificado[(m, p, p)] = 1.0
#Carregar Informações da aba Máquinas (Unidade, Máquina, Tempo de carga (h), Taxa DISP, Taxa QUAL)

maquinas = pd.read_excel('Simulador - Mix de cartões_3__2_.xlsx', sheet_name='Máquinas', skiprows=2)
maquinas_df = pd.DataFrame(maquinas)
maquinas_df = maquinas_df[['Unidade','Máquina','Tempo de carga (h)','Taxa DISP','Taxa QUAL','Produção bruta máxima (t)']]

dict_tempo_carga = {}
dict_taxa_DISP = {}
dict_taxa_QUAL_MAQUINA = {}
dict_Prod_bruta_max = {}

for index,row in maquinas_df.iterrows():
    Unidade = row['Unidade']
    Maquina = int(row['Máquina'])
    Tempo_Carga = float(row['Tempo de carga (h)'])
    Taxa_Disp = float(row['Taxa DISP'])
    Taxa_QUAL = float(row['Taxa QUAL'])
    Prod_bruta_max = row['Produção bruta máxima (t)']
    if pd.isna(Prod_bruta_max):  # Verifica se o valor é NaN
        Prod_bruta_max = 1e9  # Atribui 1e9 para indicar "sem limite"
    else:
        Prod_bruta_max = float(Prod_bruta_max)  # Converte para float se não for NaN
    dict_tempo_carga[(Maquina)] = Tempo_Carga
    dict_taxa_DISP[(Maquina)] = Taxa_Disp
    dict_taxa_QUAL_MAQUINA[(Maquina)] = Taxa_QUAL
    dict_Prod_bruta_max[(Maquina)] = Prod_bruta_max

for m in dict_tempo_carga:
    dict_tempo_carga[m] = dict_tempo_carga[m] * (DIAS_PERIODO / 365)

for m in dict_Prod_bruta_max:
    if dict_Prod_bruta_max[m] < 1e8:  # não pró-ratear o 1e9 (sem limite)
        dict_Prod_bruta_max[m] = dict_Prod_bruta_max[m] * (DIAS_PERIODO / 365)

#Carregar Informações da aba Produto por MP:

produto_por_MP = pd.DataFrame(data6)
produto_por_MP = produto_por_MP[['Produto','MP27','MP28','MC25','MC26']]
colunas_maquinas = ['MP27', 'MP28', 'MC25', 'MC26']
produto_por_MP_long = produto_por_MP.melt(id_vars=['Produto'],value_vars=colunas_maquinas,var_name='maquinas_str',value_name='flag_producao')
produto_por_MP_long['Máquina'] = produto_por_MP_long['maquinas_str'].str.replace('MP','').str.replace('MC','').astype(int)
produto_por_MP_long['flag_producao']=produto_por_MP_long['flag_producao'].astype(int)
produto_por_MP_long = produto_por_MP_long.drop(columns=['maquinas_str'])
dict_prod_por_MP = produto_por_MP_long.set_index(['Máquina','Produto'])['flag_producao'].to_dict()

#Carregar informações da Lista de Materiais

data8 = pd.read_excel('Simulador - Mix de cartões_3__2_.xlsx', sheet_name='Lista de Materiais', skiprows=2)
Lista_Mat = pd.DataFrame(data8)
Lista_Mat = Lista_Mat[['Centro','Índice','Código','UMB','Material','Valor']]
Lista_Mat_Filtered = Lista_Mat[(~Lista_Mat['Código'].astype(str).str.isnumeric())&
Lista_Mat['Índice'].isin(['Esp','Específico','Qtd'])&
~Lista_Mat['Material'].isin([
    'AGUA-TRATADA','ACIDO-SULF','VAPOR-MEDIO',
    'ENERG-MEDIA','ENERG-TERM','VAPOR-LICOR','AGUA-DESMI',
    'ENERG-PROP','AGUA-CAP1','ETE',
    'VAPOR-BIOMASSA','VAPOR-OLEO','BIOMASSA-F','CAV-ENERGIA',
    'DISPER-MP12','DISPER-MP13','DISPER-MP16','DISPER-MP23'
])].copy()

for index,row in Lista_Mat_Filtered.iterrows():
    if row['UMB'] == 'KG':
        Lista_Mat_Filtered.at[index, 'Valor'] = row['Valor'] / 1000

def explodir_lista_materiais(df: pd.DataFrame) -> Dict[Tuple[str, str], Dict[str, float]]:
    """
    Replica EXATAMENTE a função do código original lista_de_materiais.txt
    """
    def _recursive_explode(material: str, bom_map: defaultdict, factor: float = 1.0,
                           _visited: frozenset = frozenset()) -> Dict[str, float]:
        """Função interna recursiva com detecção de ciclos"""
        totals = defaultdict(float)
        componentes = bom_map.get(material, [])
        for sub_codigo, qty_direct in componentes:
            total_qty_to_add = qty_direct * factor
            totals[sub_codigo] += total_qty_to_add
            if sub_codigo in bom_map and sub_codigo not in _visited:
                sub_totals = _recursive_explode(sub_codigo, bom_map, total_qty_to_add, _visited | {material})
                for chave_sub, valor_sub in sub_totals.items():
                    totals[chave_sub] += valor_sub
        return dict(totals)

    bom_achatada_global = {}
    for centro in df['Centro'].unique():
        df_centro = df[df['Centro'] == centro]
        bom_map = defaultdict(list)
        for _, row in df_centro.iterrows():
            bom_map[row["Material"]].append((row["Código"], row["Valor"]))
        produtos_raiz = set(df_centro["Material"])
        for produto_raiz in produtos_raiz:
            resultado_explosao = _recursive_explode(produto_raiz, bom_map, factor=1.0)
            if resultado_explosao:
                bom_achatada_global[(centro, produto_raiz)] = resultado_explosao
    return bom_achatada_global

# USO CORRETO:
# Esta é a linha que o otimizador executará (via lista_materiais.explodir_lista_materiais(self.lista_materiais))
dict_consumo_especifico = explodir_lista_materiais(Lista_Mat_Filtered)

#print(dict_consumo_especifico)

todas_fibras = set()
for (centro, produto), consumos in dict_consumo_especifico.items():
    todas_fibras.update(consumos.keys())

# 2. Função para obter consumo específico produto-fibra
def get_consumo_especifico(model, centro, produto, maq, fibra):
    """
    Replica a função consumo_especifico_produto_fibra do código original
    """
    # Reconstrói o nome do produto como aparece no Excel
    produto_ajustado = produto[:3] + str(maq).zfill(2) + produto[3:]
    chave_produto = (centro, produto_ajustado)
    return dict_consumo_especifico.get(chave_produto, {}).get(fibra, 0)

# 4. Função auxiliar para consumo de cavaco (se necessário)
def get_consumo_cavaco(model, centro, fibra, cavaco_madeira):
    """
    Replica a função consumo_especifico_cavaco do código original
    """
    chave_fibra = (centro, fibra)
    return dict_consumo_especifico.get(chave_fibra, {}).get(cavaco_madeira, 0)


#Constante de horas
horas_dia = 24

print(f"\n{'='*70}")
print(f"mix_attempt_30dias.py — Horizonte: {DIAS_PERIODO} dias")
print(f"{'='*70}")

# ── Diagnóstico de coeficientes de consumo específico ────────────────
print("\n" + "="*70)
print("DIAGNÓSTICO ORT — COEFICIENTES DE CONSUMO ESPECÍFICO (Lista de Materiais)")
print("="*70)

fibras_balanco_ORT = ['CKB-FC', 'CKB-FL']
coefs_ORT = []
for (centro, produto_lm), consumos in dict_consumo_especifico.items():
    for fibra, valor in consumos.items():
        if fibra in fibras_balanco_ORT and centro == 'ORT':
            coefs_ORT.append((produto_lm, fibra, valor))

if coefs_ORT:
    vals_ORT = [v for _, _, v in coefs_ORT if v > 0]
    print(f"\n  ORT — Digestores (CKB-FC, CKB-FL):")
    print(f"    Total de coeficientes não-zero: {len(vals_ORT)}")
    if vals_ORT:
        print(f"    Mínimo:  {min(vals_ORT):.6f} t/t")
        print(f"    Máximo:  {max(vals_ORT):.6f} t/t")
        print(f"    Média:   {sum(vals_ORT)/len(vals_ORT):.6f} t/t")
        top5 = sorted(coefs_ORT, key=lambda x: -x[2])[:5]
        print(f"    Top 5 maiores:")
        for prod, fib, val in top5:
            print(f"      {prod:<20} {fib:<15} {val:.6f}")

print("="*70)

#Carregamento dados de balanço de ORT:

#Tabela - Produção celulose e consumo fibras
data13 = pd.read_excel('Simulador - Mix de cartões_3__2_.xlsx', sheet_name='Balanço fábrica ORT', skiprows=3,usecols='B:E')
Prod_Cel_Fibras_ORT = pd.DataFrame(data13).dropna()
Prod_Cel_Fibras_ORT['Consumo_Anual_Fibras'] = Prod_Cel_Fibras_ORT['Consumo (t/dia)']*Prod_Cel_Fibras_ORT['Dias operação']
Prod_Cel_Fibras_ORT = Prod_Cel_Fibras_ORT.groupby('Fibra')['Consumo_Anual_Fibras'].sum().reset_index()

#Tabela Parâmetros adicionais
data14 = pd.read_excel('Simulador - Mix de cartões_3__2_.xlsx', sheet_name='Balanço fábrica ORT', skiprows=3,usecols='G:I')
Param_add_ORT = pd.DataFrame(data14).dropna()

data15 = pd.read_excel('Simulador - Mix de cartões_3__2_.xlsx', sheet_name='Balanço fábrica ORT', skiprows=3,usecols='K:P')
Capac_plantas_ORT = pd.DataFrame(data15).dropna()

data16 = pd.read_excel('Simulador - Mix de cartões_3__2_.xlsx', sheet_name='Balanço fábrica ORT', skiprows=3,usecols='R:U')
Fibras_Digestores_ORT = pd.DataFrame(data16).dropna()

#Dicionário para Produção de celulose e consumo de fibras - CONSUMO ANUAL:
dict_consumo_fibras_ORT = Prod_Cel_Fibras_ORT.set_index('Fibra')['Consumo_Anual_Fibras'].to_dict()

for fibra in dict_consumo_fibras_ORT:
    dict_consumo_fibras_ORT[fibra] = dict_consumo_fibras_ORT[fibra] * (DIAS_PERIODO / 365)

#Dicionário para Parâmetros adicionais:
dict_param_add_ORT = Param_add_ORT.set_index('Parâmetro')['Valor'].to_dict()

#Dicionário para Capacidade das Plantas:
dict_emissario = Capac_plantas_ORT.set_index('Área.1')['Capacidade MSR'].to_dict()
dict_capacmax_ort = Capac_plantas_ORT.set_index('Área.1')['Capacidade Máx'].to_dict()
dict_dias_operacao = Capac_plantas_ORT.set_index('Área.1')['Dias operação.1'].to_dict()
# Aplica overrides de capacidade do cache (substitui apenas as chaves presentes)
_cap_cache = _cache_get("capacidades")
if _cap_cache:
    dict_emissario.update(_cap_cache)

#Dicionário para Fibras e Digestores:
dict_rendimento_ort = Fibras_Digestores_ORT.set_index('Fibra.1')['Rendimento (%)'].to_dict()
dict_carga_alcalina_ort = Fibras_Digestores_ORT.set_index('Fibra.1')['Carga alcalina (%)'].to_dict()

#Lista de fibras - Acrescentado em SETS:
lista_fibras_ORT = list(Prod_Cel_Fibras_ORT['Fibra'].unique())
fibras_cdr_ORT = ['CKN-FC', 'CKN-FL', 'BCTMP']
for _f in fibras_cdr_ORT:
    if _f not in lista_fibras_ORT:
        lista_fibras_ORT.append(_f)

# Inverso: digestor → lista de fibras (apenas ORT)
dict_digestor = defaultdict(list)
for fibra in lista_fibras_ORT:
    dict_digestor[fibra].append(fibra)

#Lista de parâmetros adicionais - Acrescentando em SETS:
nomes_parametros_balanco_ORT = list(Param_add_ORT['Parâmetro'].unique())

#Lista de sets para Capacidade das Plantas:
nomes_area_capacPlantas = list(Capac_plantas_ORT['Área.1'].unique())

#Listas de fibras:
nomes_fibras_ort = Fibras_Digestores_ORT['Fibra.1'].unique()

#Carregar dados da aba Custos:
_custos_cache = _cache_get("custos")
if _custos_cache:
    dados_custos = pd.DataFrame(_custos_cache).rename(columns={"Custo Variavel": "Custo Variavel "}).dropna()
    dados_custos = dados_custos[['Produto','Máquina','Custo Variavel ']]
else:
    data17 = pd.read_excel('Simulador - Mix de cartões_3__2_.xlsx', sheet_name='Custos', skiprows=2)
    dados_custos = pd.DataFrame(data17)[['Produto','Máquina','Custo Variavel ']].dropna()

#Dicionário custos (apenas máquinas ORT):
dict_custos = {
    (produto, maquina): custo
    for (produto, maquina), custo in dados_custos.set_index(['Produto','Máquina'])['Custo Variavel '].to_dict().items()
    if maquina in maquinas_ort
}

#Carregar dados da aba Demanda:
_demanda_cache = _cache_get("demanda")
if _demanda_cache:
    dados_demanda = pd.DataFrame(_demanda_cache)[['Mercado','Produto','Quantidade (TO)','Preço']]
else:
    data18 = pd.read_excel('Simulador - Mix de cartões_3__2_.xlsx', sheet_name='Demanda', skiprows=2)
    dados_demanda = pd.DataFrame(data18)[['Mercado','Produto','Quantidade (TO)','Preço']]
data19 = pd.read_excel('Simulador - Mix de cartões_3__2_.xlsx', sheet_name='Reprocesso', skiprows=2)
dados_reprocesso = pd.DataFrame(data19)
dados_reprocesso = dados_reprocesso[['Produto vendável','Produto base']]
#dict_reprocesso = dados_reprocesso.set_index('Produto base')['Produto vendável'].to_dict()
dict_reprocesso = dados_reprocesso.set_index('Produto vendável')['Produto base'].to_dict()

for index, row in dados_demanda.iterrows():
    produto = row['Produto']
    if produto in dict_reprocesso:
        dados_demanda.at[index, 'Produto'] = dict_reprocesso[produto]

#Dicionáro

#dict_demanda = dados_demanda.groupby(['Produto', 'Mercado'])['Quantidade (TO)'].sum().to_dict()
dict_demanda = {(p, m): v * (DIAS_PERIODO / BASE_DEMANDA_DIAS) for (p, m), v in dados_demanda.groupby(['Produto', 'Mercado'])['Quantidade (TO)'].sum().to_dict().items()}

dict_preco = dados_demanda.set_index(['Produto', 'Mercado'])['Preço'].to_dict() #IMPORTANTE - NÃO PODE HAVER DUPLICATAS PARA O MESMO PRODUTO, MESMO MERCADO NA ABA DEMANDAS
#dict_preco = dados_demanda.groupby(['Produto', 'Mercado']).apply(
#    lambda x: (x['Preço'] * x['Quantidade (TO)']).sum() / x['Quantidade (TO)'].sum()
#).to_dict()

#Carregador dados da aba Parâmetros:

data20 = pd.read_excel('Simulador - Mix de cartões_3__2_.xlsx', sheet_name='Parâmetros', skiprows=2)
parametros1 = pd.DataFrame(data20)
parametros1 = parametros1[['Parâmetro','Valor']]

# Remove linhas com Parâmetro NaN ou que são cabeçalhos intermediários
parametros1 = parametros1.dropna(subset=['Parâmetro'])
parametros1 = parametros1[parametros1['Parâmetro'] != 'Configurações do modelo:']

#dicionário Parâmetros:
dict_param_modelo = parametros1.set_index('Parâmetro')['Valor'].to_dict()

# Parâmetros de remuneração e priorização — lidos do dict
flag_remuneracao_celulose = dict_param_modelo.get('Incluir remuneração da celulose', 0)

_param_cache = _cache_get("parametros") or {}
cambio = _param_cache.get('Cambio', dict_param_modelo.get('Cambio', 1))
custo_variavel_fibra_curta = _param_cache.get('Custo Variavel Celulose FC', dict_param_modelo.get('Custo Variavel Celulose FC', 0))
custo_variavel_fibra_longa = _param_cache.get('Custo Variavel Celulose FL', dict_param_modelo.get('Custo Variavel Celulose FL', 0))
preco_venda_fibra_curta = _param_cache.get('Preço de Cel. Merc. ME FC', dict_param_modelo.get('Preço de Cel. Merc. ME FC', 0)) * cambio
preco_venda_fibra_longa = _param_cache.get('Preço de Cel. Merc. ME FF', dict_param_modelo.get('Preço de Cel. Merc. ME FF', 0)) * cambio

dict_penalidade_ociosidade = {
    27: 0,
    28: 0,
    25: 0,
    26: 0,
}

lista_parametros = list(parametros1['Parâmetro'])

###### Cálculo do Consumo de Fibras (apenas ORT):

def calc_consumo_fibra(model, centro, fibra):
    if fibra not in lista_fibras_ORT:
        return 0
    consumo = sum(
        model.producao_bruta[p, m] * get_consumo_especifico(model, centro, p, m, fibra)
        for p in model.produtos
        for m in model.maquinas
    )
    consumo += dict_consumo_fibras_ORT.get(fibra, 0)
    return consumo

## Cálculo ORT - Caustificação

def calc_caustificacao(model):
    centro = 'ORT'
    licor_branco = 0
    dias = DIAS_PERIODO
    for fibra in lista_fibras_ORT:
        rendimento = dict_rendimento_ort.get(fibra, 1)
        carga_alcalina = dict_carga_alcalina_ort.get(fibra, 0)
        if 'CKB' in fibra:
            concentracao = dict_param_add_ORT.get('Concentração caust 1 (g/l NaOH)', 1) / 1000
            perda = dict_param_add_ORT.get('Perda de fibra branca (%)', 0)
            consumo_dia = calc_consumo_fibra(model, centro, fibra) / (1 - perda) / dias
        else:
            concentracao = dict_param_add_ORT.get('Concentração caust 2 (g/l NaOH)', 1) / 1000
            consumo_dia = calc_consumo_fibra(model, centro, fibra) / dias
        licor_branco += 0.9 * consumo_dia * carga_alcalina / rendimento / concentracao

    dias_deslig = DIAS_PERIODO
    deslig_euca = calc_consumo_fibra(model, 'ORT', 'CKB-FC') * dict_param_add_ORT.get('Deslignificação Euca (m3/tsa)', 0) / dias_deslig
    deslig_pinus = calc_consumo_fibra(model, 'ORT', 'CKB-FL') * dict_param_add_ORT.get('Deslignificação Pinus (m3/tsa)', 0) / dias_deslig
    licor_angatuba = dict_param_add_ORT.get('Licor branco Angatuba (m3 LB/d)', 0)

    return licor_branco + deslig_euca + deslig_pinus + licor_angatuba

## Cálculo ORT Efluentes

def calc_captacao(model):
    prod_papeis = sum(
        model.producao_bruta[p, m]
        for p in model.produtos
        for m in [27, 28]
    )
    prod_cel = sum(dict_consumo_fibras_ORT.get(f, 0) for f in lista_fibras_ORT)
    dias = DIAS_PERIODO
    captacao = dict_param_add_ORT.get('Captação de Água (m3/t)', 0)
    return captacao * (prod_papeis + prod_cel) / horas_dia / dias

def calc_emissario(model):
    prod_papeis = sum(
        model.producao_bruta[p, m]
        for p in model.produtos
        for m in [27, 28]
    )
    prod_cel = sum(dict_consumo_fibras_ORT.get(f, 0) for f in lista_fibras_ORT)
    dias = DIAS_PERIODO
    emissario = dict_param_add_ORT.get('Emissário (m3/t)', 0)
    return emissario * (prod_papeis + prod_cel) / horas_dia / dias

## Cálculo - ORT Produção por digestor

def calc_producao_digestor_ORT(model, digestor):
    centro = 'ORT'
    fibras_digestor = dict_digestor.get(digestor, [])
    dias = DIAS_PERIODO
    return sum(
        calc_consumo_fibra(model, centro, fibra) / dias
        for fibra in fibras_digestor
        if fibra in lista_fibras_ORT
    )

## Cálculo ORT — CDR (Caldeira de Recuperação)

def calc_tss_fibra_ORT(model, fibra):
    """TSS anual gerado pela fibra."""
    centro = 'ORT'
    coef_ss = dict_consumo_especifico.get((centro, fibra), {}).get('SOLIDO-SECO', 0)
    return -1 * coef_ss * calc_consumo_fibra(model, 'ORT', fibra)

def calc_cdr(model):
    fibras_cdr = ['CKB-FC', 'CKB-FL', 'CKN-FC', 'CKN-FL', 'BCTMP']
    dias = dict_dias_operacao.get('CDR', DIAS_PERIODO)
    total_tss = sum(calc_tss_fibra_ORT(model, fibra) for fibra in fibras_cdr)

    volume_angatuba = dict_param_add_ORT.get('Volume recebido de Licor Preto de Angatuba (m³/d)', 0)
    conc_angatuba   = dict_param_add_ORT.get('Concentração Licor Preto recebido de Angatuba (%)', 0)
    tss_angatuba_dia = volume_angatuba * conc_angatuba

    return (total_tss / dias) + tss_angatuba_dia

## Cálculo ORT — Evaporação

def calc_evaporacao_ORT(model):
    mapa_conc_digestor = {
        'CKB-FC': 'Concentração Licor Preto gerado no digestor 1 (%)',
        'CKB-FL': 'Concentração Licor Preto gerado no digestor 2 (%)',
        'CKN-FC': 'Concentração Licor Preto gerado no digestor 3 (%)',
        'CKN-FL': 'Concentração Licor Preto gerado no digestor 4 (%)',
        'BCTMP':  'Concentração Licor Preto gerado na BCTMP (%)',
    }

    conc_saida_evap = dict_param_add_ORT.get('Concentração Licor Preto na Saída Evap1', 1)

    volume_lp = 0
    tss_total  = 0

    for fibra, chave_conc in mapa_conc_digestor.items():
        tss_fibra = calc_tss_fibra_ORT(model, fibra)
        conc_digestor = dict_param_add_ORT.get(chave_conc, 1)
        volume_lp += tss_fibra / conc_digestor
        tss_total  += tss_fibra

    dias = dict_dias_operacao.get('CDR', DIAS_PERIODO)
    volume_angatuba  = dict_param_add_ORT.get('Volume recebido de Licor Preto de Angatuba (m³/d)', 0)
    conc_angatuba    = dict_param_add_ORT.get('Concentração Licor Preto recebido de Angatuba (%)', 0)
    tss_angatuba_dia = volume_angatuba * conc_angatuba
    volume_lp += volume_angatuba * dias
    tss_total  += tss_angatuba_dia * dias

    return (1 / (dias * 24)) * (volume_lp - tss_total / conc_saida_evap)

#MODEL

model = pyo.ConcreteModel()

#SETS

model.maquinas = pyo.Set(initialize={27,28,25,26})  ###Realmente utilizada
model.produtos = pyo.Set(initialize=lista_produtos)   ###Realmente utilizada
model.mercados = pyo.Set(initialize={'ME','MI','Transferência'}) ###Realmente utilizada
model.centros = pyo.Set(initialize={'ORT'})
model.fibras = pyo.Set(initialize=todas_fibras)
model.tipos_cavaco = pyo.Set(initialize=['CAV-EUCALIPTO', 'CAV-PINUS'])
model.fibras_ORT = pyo.Set(initialize=lista_fibras_ORT)
model.nomes_parametros_balanco_ORT = pyo.Set(initialize=nomes_parametros_balanco_ORT)
model.nomes_areas_capacPlantas = pyo.Set(initialize=nomes_area_capacPlantas)
model.fibras_ort = pyo.Set(initialize=nomes_fibras_ort)
model.parametros_modelo = pyo.Set(initialize=lista_parametros)

#PARAMETERS:

#Parâmetros de Produtividade:
model.Produtividade_max = pyo.Param(model.maquinas,model.produtos,initialize=dict_produtividade,default=0.0)
model.Taxa_PERF = pyo.Param(model.maquinas,model.produtos,initialize=dict_taxa_performance,default=0.0)
model.Produtividade_Bruta = pyo.Param(model.maquinas,model.produtos,initialize=dict_produtividade_bruta)
model.Refugo_Ajustado = pyo.Param(model.maquinas, model.produtos,initialize=lambda m, maq, prod: dict_refugo_ajustado.get((maq, prod), 0.0))
model.Waste = pyo.Param(model.maquinas, model.produtos,initialize=lambda m, maq, prod: dict_waste.get((maq, prod), 0.0))
model.Total_Waste = pyo.Param(model.maquinas, model.produtos,initialize=lambda m, maq, prod: dict_total_waste.get((maq, prod), 0.0))

#Parâmetros de Total Waste:
model.IAC = pyo.Param(model.maquinas,model.produtos, initialize=lambda m,maq,prod:dict_IAC.get((maq,prod),1))
model.Refugo_MP = pyo.Param(model.maquinas, model.produtos, initialize=lambda m, maq, prod: dict_Refugo_MP.get((maq, prod), 0.0))
model.Rep_Externo = pyo.Param(model.maquinas, model.produtos, initialize=lambda m, maq, prod: dict_Rep_Externo.get((maq, prod), 0.0))
model.MR2 = pyo.Param(model.maquinas, model.produtos, initialize=lambda m, maq, prod: dict_MR2.get((maq, prod), 0.0))
model.Sala_Perdas = pyo.Param(model.maquinas, model.produtos, initialize=lambda m, maq, prod: dict_Sala_Perdas.get((maq, prod), 0.0))
model.Cortadeira_Perda_Gramatura = pyo.Param(model.maquinas, model.produtos, initialize=lambda m, maq, prod: dict_Cortadeira_Perda_Gramatura.get((maq, prod), 0.0))
model.Cortadeira_Perdas_Cortadeira = pyo.Param(model.maquinas, model.produtos, initialize=lambda m, maq, prod: dict_Cortadeira_Perdas_Cortadeira.get((maq, prod), 0.0))
model.Estoque_Perdas = pyo.Param(model.maquinas, model.produtos, initialize=lambda m, maq, prod: dict_Estoque_Perdas.get((maq, prod), 0.0))
model.Estoque_Perdas_Refugo = pyo.Param(model.maquinas, model.produtos, initialize=lambda m, maq, prod: dict_Estoque_Perdas_Refugo.get((maq, prod), 0.0))

#Parâmetros de Desclassificados:
model.Taxa_Desclassificacao = pyo.Param(model.maquinas,model.produtos,model.produtos,initialize=lambda m,maq,prod,prod2: dict_taxa_desclassificado.get((maq,prod,prod2),0.0))

#Parâmetros de Máquinas:
model.Tempo_Carga = pyo.Param(model.maquinas,initialize=lambda m,maq:dict_tempo_carga.get((maq),0.0))
model.Taxa_DISP = pyo.Param(model.maquinas,initialize=lambda m,maq:dict_taxa_DISP.get((maq),0.0))
model.Taxa_QUAL_MAQUINA = pyo.Param(model.maquinas,initialize=lambda m,maq:dict_taxa_QUAL_MAQUINA.get((maq),0.0))
model.Prod_bruta_max = pyo.Param(model.maquinas,initialize=lambda m,maq:dict_Prod_bruta_max.get((maq),1e9))

#Parâmetros de Produto por Máquina:

model.Prod_por_Maq = pyo.Param(model.maquinas, model.produtos, initialize=lambda m, maq, prod: dict_prod_por_MP.get((maq, prod), 0.0))

# Parâmetros de Lista de Materiais:

model.Consumo_Especifico = pyo.Param(model.centros,model.produtos,model.maquinas,model.fibras,initialize=get_consumo_especifico,default=0.0)
model.Consumo_Cavaco = pyo.Param(model.centros,model.fibras,model.tipos_cavaco,initialize=get_consumo_cavaco,default=0.0)

#Parâmetros de Balanço de Fábrica ORT:

# Parâmetros da seção - Produção de celulose e consumo de fibras:
model.consumo_fibras_ORT = pyo.Param(model.fibras_ORT,initialize=lambda m,fibra: dict_consumo_fibras_ORT.get((fibra),0))
# Parâmetros adicionais ORT - mapeado por nome de parâmetro
model.Param_Balanco_ORT = pyo.Param(model.nomes_parametros_balanco_ORT, initialize=lambda m,nome_param: dict_param_add_ORT.get((nome_param),0))
#Parâmetros Capacidade das Plantas:
model.emissario = pyo.Param(model.nomes_areas_capacPlantas, initialize=lambda m,msr:dict_emissario.get(msr,0))
#Parâmetros Capacidade_Max das Plantas:
model.capac_max = pyo.Param(model.nomes_areas_capacPlantas, initialize=lambda m,capacmax:dict_capacmax_ort.get(capacmax,0))
#Parâmetros Dias Operação das Plantas:
model.dias_op = pyo.Param(model.nomes_areas_capacPlantas, initialize=lambda m,dias:dict_dias_operacao.get(dias,0))
#Parâmetros Fibras e digestores:
model.Rendimento_ORT = pyo.Param(model.fibras_ort,initialize=lambda m, fibra: dict_rendimento_ort.get(fibra, 0.0))
model.Carga_Alcalina_ORT = pyo.Param(model.fibras_ort,initialize=lambda m, fibra: dict_carga_alcalina_ort.get(fibra, 0.0))

#Parâmetros de Custos:
model.custos = pyo.Param(model.produtos,model.maquinas,initialize=lambda m,produto,maquina: dict_custos.get((produto,maquina),0.0))

#Parâmetros de demanda:

model.demanda = pyo.Param(model.produtos,model.mercados, initialize=lambda m,produtos,mercados: dict_demanda.get((produtos,mercados),0.0))
model.precos = pyo.Param(model.produtos,model.mercados, initialize=lambda m,produtos,mercados: dict_preco.get((produtos,mercados),0.0))

#Parâmetros de Parâmetros do modelo:

model.param_modelo = pyo.Param(model.parametros_modelo,initialize=lambda m,parametro: dict_param_modelo.get(parametro,0.0))


#DECISION VARIABLES

# VARIÁVEIS DE DECISÃO

# VARIÁVEIS DE DECISÃO — estrutura correta

# Produção bruta: física na máquina, SEM mercado
model.producao_bruta = pyo.Var(model.produtos, model.maquinas, domain=pyo.NonNegativeReals)

# Produção líquida: pool físico por máquina, SEM mercado
model.producao_liquida = pyo.Var(model.produtos, model.maquinas, domain=pyo.NonNegativeReals)

# Produção vendável: COM mercado — aqui ocorre a alocação
model.producao_vendavel = pyo.Var(model.produtos, model.maquinas, model.mercados, domain=pyo.NonNegativeReals)

# Demanda não atendida: COM mercado
model.demanda_nao_atendida = pyo.Var(model.produtos, model.mercados, domain=pyo.NonNegativeReals)

# Produção adicional: SEM mercado
model.producao_adicional = pyo.Var(model.produtos, model.maquinas, domain=pyo.NonNegativeReals)

# Capacidade ociosa: SEM mercado
model.capacidade_ociosa = pyo.Var(model.maquinas, domain=pyo.NonNegativeReals)

#Restrições

# Demanda — por mercado (igual ao que já tem)
def Constraint_demanda(model, p, mercado):
    return (sum(model.producao_vendavel[p, m, mercado] for m in model.maquinas)
            + model.demanda_nao_atendida[p, mercado] == model.demanda[p, mercado])
model.Constraint_demanda = pyo.Constraint(model.produtos, model.mercados, rule=Constraint_demanda)

# Produção vendável = pool líquido alocado por mercado — SEM mercado na líquida
def Constraint_prod_vendavel(model, p, m):
    taxa_waste = dict_waste.get((m, p), 0)
    return (sum(model.producao_vendavel[p, m, mercado] for mercado in model.mercados)
            == model.producao_liquida[p, m] * (1 - taxa_waste))
model.Constraint_prod_vendavel = pyo.Constraint(model.produtos, model.maquinas, rule=Constraint_prod_vendavel)

# Desclassificados — SEM mercado na bruta e adicional
def Constraint_desclassificados(model, p, m):
    taxa_refugo = dict_refugo_ajustado.get((m, p), 0)
    return (model.producao_liquida[p, m]
            + (1 - taxa_refugo) * (
                model.producao_adicional[p, m]
                - sum(model.producao_bruta[p_origem, m]
                      * dict_taxa_desclassificado.get((m, p_origem, p), 1 if p_origem == p else 0)
                      for p_origem in model.produtos)
            ) == 0)
model.Constraint_desclassificados = pyo.Constraint(model.produtos, model.maquinas, rule=Constraint_desclassificados)

# Produto por máquina — SEM mercado
def Constraint_maquinas(model, p, m):
    if dict_prod_por_MP.get((m, p), 0) == 0:
        return model.producao_bruta[p, m] == 0
    return pyo.Constraint.Skip
model.Constraint_maquinas = pyo.Constraint(model.produtos, model.maquinas, rule=Constraint_maquinas)

# Tempo disponível — SEM mercado
def Constraint_tempo_producao(model, m):
    tempo_usado = sum(
        model.producao_bruta[p, m] / dict_produtividade_bruta.get((m, p), 1)
        for p in model.produtos
        if dict_produtividade_bruta.get((m, p), 0) > 0
    )
    return tempo_usado + model.capacidade_ociosa[m] == model.Tempo_Carga[m] * model.Taxa_DISP[m]
model.Constraint_tempo_producao = pyo.Constraint(model.maquinas, rule=Constraint_tempo_producao)

# Produção máxima — SEM mercado
def Constraint_producao_maxima(model, m):
    return (sum(model.producao_bruta[p, m] for p in model.produtos)
            <= model.Prod_bruta_max[m])
model.Constraint_producao_maxima = pyo.Constraint(model.maquinas, rule=Constraint_producao_maxima)

# Balanço ORT — Caustificação
def Constraint_caustificacao(model):
    return calc_caustificacao(model) <= dict_emissario.get('Caustificação', 0)
model.Constraint_caustificacao = pyo.Constraint(rule=Constraint_caustificacao)

# Balanço ORT — Captação
def Constraint_captacao(model):
    return calc_captacao(model) <= dict_emissario.get('Outorga Captação', 0)
model.Constraint_captacao = pyo.Constraint(rule=Constraint_captacao)

# Balanço ORT — Emissário
def Constraint_emissario(model):
    return calc_emissario(model) <= dict_emissario.get('Outorga Emissario', 0)
model.Constraint_emissario = pyo.Constraint(rule=Constraint_emissario)

# Balanço ORT — Digestores
def Constraint_digestor_ORT(model, digestor):
    return calc_producao_digestor_ORT(model, digestor) <= dict_emissario.get(digestor, 1e9)
model.Constraint_digestor_ORT = pyo.Constraint(list(lista_fibras_ORT), rule=Constraint_digestor_ORT)

# Balanço ORT — CDR (Caldeira de Recuperação)
def Constraint_cdr(model):
    return calc_cdr(model) <= dict_emissario.get('CDR', 1e9)
model.Constraint_cdr = pyo.Constraint(rule=Constraint_cdr)

# Balanço ORT — Evaporação
def Constraint_evaporacao_ORT(model):
    return calc_evaporacao_ORT(model) <= dict_emissario.get('Evaporação', 1e9)
model.Constraint_evaporacao_ORT = pyo.Constraint(rule=Constraint_evaporacao_ORT)

# Função Objetivo — Max. Margem (Receita - Custo Variável)
def obj_max_margem(model):
    receita = sum(
        model.producao_vendavel[p, m, merc] * dict_preco.get((p, merc), 0)
        for p in model.produtos
        for m in model.maquinas
        for merc in model.mercados
    )
    custo = sum(
        model.producao_bruta[p, m] * (
                dict_custos.get((p, m), 0)
                + flag_remuneracao_celulose * (
                        get_consumo_especifico(model, 'ORT', p, m, 'CKB-FC') * (
                            preco_venda_fibra_curta - custo_variavel_fibra_curta)
                        + get_consumo_especifico(model, 'ORT', p, m, 'CKB-FL') * (
                                    preco_venda_fibra_longa - custo_variavel_fibra_longa)
                )
        )
        for p in model.produtos
        for m in model.maquinas
    )
    penalidade_ociosidade = sum(
        model.capacidade_ociosa[m] * dict_penalidade_ociosidade.get(m, 0)
        for m in model.maquinas
    )
    return receita - custo - penalidade_ociosidade
model.obj = pyo.Objective(rule=obj_max_margem, sense=pyo.maximize)

#"""
# Solver
solver = pyo.SolverFactory('glpk', executable=get_glpk_path())
results = solver.solve(model, tee=True)

# Verificação do status
if results.solver.termination_condition == pyo.TerminationCondition.optimal:
    print("\n✓ Solução ótima encontrada!")

    print("\n=== PRODUÇÃO BRUTA (> 0) ===")
    for p in model.produtos:
        for m in model.maquinas:
            val = pyo.value(model.producao_bruta[p, m])
            if val and val > 0:
                print(f"  {p} | MP{m}: {val:.1f} t")

    print("\n=== PRODUÇÃO VENDÁVEL (> 0) ===")
    for p in model.produtos:
        for m in model.maquinas:
            for merc in model.mercados:
                val = pyo.value(model.producao_vendavel[p, m, merc])
                if val and val > 0:
                    print(f"  {p} | MP{m} | {merc}: {val:.1f} t")

    print("\n=== DEMANDA NÃO ATENDIDA (> 0) ===")
    for p in model.produtos:
        for merc in model.mercados:
            val = pyo.value(model.demanda_nao_atendida[p, merc])
            if val and val > 0:
                print(f"  {p} | {merc}: {val:.1f} t")

    print("\n=== CAPACIDADE OCIOSA (> 0) ===")
    for m in model.maquinas:
        val = pyo.value(model.capacidade_ociosa[m])
        if val and val > 0:
            print(f"  MP{m}: {val:.1f} h")

    print("=== VALIDAÇÃO DAS RESTRIÇÕES ===")
    print("\nTempo usado vs disponível (h):")
    for m in model.maquinas:
        tempo_usado = sum(
            pyo.value(model.producao_bruta[p, m]) / dict_produtividade_bruta.get((m, p), 1)
            for p in model.produtos
            if dict_produtividade_bruta.get((m, p), 0) > 0
        )
        tempo_disp = dict_tempo_carga.get(m, 0) * dict_taxa_DISP.get(m, 1)
        print(f"  MP{m}: usado={tempo_usado:.1f} | disponível={tempo_disp:.1f} | ocioso={tempo_disp - tempo_usado:.1f}")

    print("\nDemanda atendida vs total:")
    total_demanda = 0
    total_atendida = 0
    total_nao_atendida = 0

    for (p, merc), qtd in dict_demanda.items():
        if qtd == 0:
            continue
        total_demanda += qtd
        atendida = sum(pyo.value(model.producao_vendavel[p, m, merc]) for m in model.maquinas)
        nao_atendida = pyo.value(model.demanda_nao_atendida[p, merc])
        total_atendida += atendida
        total_nao_atendida += nao_atendida

    print(f"  Total demanda:      {total_demanda:,.1f} t")
    print(f"  Total atendida:     {total_atendida:,.1f} t")
    print(f"  Total não atendida: {total_nao_atendida:,.1f} t")
    print(f"  Cobertura:          {100 * total_atendida / total_demanda:.1f}%")
#"""

########## Código

print("\n" + "="*110)
print("COMPARAÇÃO DEMANDA vs PRODUÇÃO VENDÁVEL POR PRODUTO E MERCADO")
print("="*110)
print(f"{'Produto':<20} {'Mercado':<15} {'Demanda':>12} {'Atendida':>12} {'Diferença':>12} {'Cobertura':>10}  {'Máquinas utilizadas'}")
print("-"*110)

linhas = []
for (p, merc), qtd in dict_demanda.items():
    if qtd == 0:
        continue
    atendida = sum(pyo.value(model.producao_vendavel[p, m, merc]) for m in model.maquinas)
    diferenca = qtd - atendida
    cobertura = 100 * atendida / qtd if qtd > 0 else 100

    # Máquinas que efetivamente produziram para esse produto/mercado
    maquinas_usadas = [
        f"MP{m}={pyo.value(model.producao_vendavel[p, m, merc]):,.0f}t"
        for m in model.maquinas
        if pyo.value(model.producao_vendavel[p, m, merc]) > 0.1
    ]
    maquinas_str = ", ".join(maquinas_usadas) if maquinas_usadas else "—"

    linhas.append((p, merc, qtd, atendida, diferenca, cobertura, maquinas_str))

linhas.sort(key=lambda x: -x[4])

total_dem = 0
total_at = 0
total_dif = 0

for p, merc, qtd, atendida, diferenca, cobertura, maquinas_str in linhas:
    total_dem += qtd
    total_at += atendida
    total_dif += diferenca
    flag = " ← sem máquina" if not any(dict_prod_por_MP.get((m, p), 0) == 1 for m in [27,28,25,26]) else ""
    print(f"  {p:<18} {merc:<15} {qtd:>12,.1f} {atendida:>12,.1f} {diferenca:>12,.1f} {cobertura:>9.1f}%  {maquinas_str}{flag}")

print("-"*110)
print(f"  {'TOTAL':<18} {'':15} {total_dem:>12,.1f} {total_at:>12,.1f} {total_dif:>12,.1f} {100*total_at/total_dem:>9.1f}%")

print("\n" + "="*70)
print("UTILIZAÇÃO DAS RESTRIÇÕES DE BALANÇO DE FÁBRICA")
print("="*70)

print(f"\n--- FÁBRICA ORT ---")

caust = pyo.value(calc_caustificacao(model))
caust_lim = dict_emissario.get('Caustificação', 0)
print(f"  Caustificação:     {caust:>10.2f} / {caust_lim:>10.2f}  ({100*caust/caust_lim:.1f}%)")

capt = pyo.value(calc_captacao(model))
capt_lim = dict_emissario.get('Outorga Captação', 0)
print(f"  Outorga Captação:  {capt:>10.2f} / {capt_lim:>10.2f}  ({100*capt/capt_lim:.1f}%)")

emis = pyo.value(calc_emissario(model))
emis_lim = dict_emissario.get('Outorga Emissario', 0)
print(f"  Outorga Emissário: {emis:>10.2f} / {emis_lim:>10.2f}  ({100*emis/emis_lim:.1f}%)")

print(f"\n  Digestores ORT:")
for fibra in lista_fibras_ORT:
    prod = pyo.value(calc_producao_digestor_ORT(model, fibra))
    lim = dict_emissario.get(fibra, 0)
    pct = 100*prod/lim if lim > 0 else 0
    print(f"    {fibra:<10}: {prod:>10.2f} / {lim:>10.2f}  ({pct:.1f}%)")

cdr_val = pyo.value(calc_cdr(model))
cdr_lim = dict_emissario.get('CDR', 0)
print(f"\n  CDR (Caldeira de Recuperação):")
print(f"    CDR: {cdr_val:>10.2f} / {cdr_lim:>10.2f}  ({100*cdr_val/cdr_lim:.1f}% )" if cdr_lim > 0 else f"    CDR: {cdr_val:>10.2f} / sem limite")

evap_ort_val = pyo.value(calc_evaporacao_ORT(model))
evap_ort_lim = dict_emissario.get('Evaporação', 0)
print(f"\n  Evaporação ORT:")
print(f"    Evap: {evap_ort_val:>10.2f} / {evap_ort_lim:>10.2f}  ({100*evap_ort_val/evap_ort_lim:.1f}%)" if evap_ort_lim > 0 else f"    Evap: {evap_ort_val:>10.2f} / sem limite")


import json

# ── Diagnóstico IA ─────────────────────────────────────────────────────────

_diag_r = {}

for _m in model.maquinas:
    _t_us = sum(
        (pyo.value(model.producao_bruta[_p, _m]) or 0) / dict_produtividade_bruta.get((_m, _p), 1)
        for _p in model.produtos if dict_produtividade_bruta.get((_m, _p), 0) > 0
    )
    _t_lim = dict_tempo_carga.get(_m, 0) * dict_taxa_DISP.get(_m, 1)
    _diag_r[f'Tempo MP{_m}'] = {'nome': f'Tempo MP{_m}', 'usado': round(_t_us, 1),
                                  'limite': round(_t_lim, 1), 'unidade': 'h', 'tipo': 'maquina'}

for _nm, _vl, _lm, _un in [
    ('Caustificação',    pyo.value(calc_caustificacao(model)),    dict_emissario.get('Caustificação', 0),     'm³/d'),
    ('Outorga Captação', pyo.value(calc_captacao(model)),         dict_emissario.get('Outorga Captação', 0),  'm³/h'),
    ('Outorga Emissário',pyo.value(calc_emissario(model)),        dict_emissario.get('Outorga Emissario', 0), 'm³/h'),
    ('CDR',              pyo.value(calc_cdr(model)),              dict_emissario.get('CDR', 0),               'tss/d'),
    ('Evaporação ORT',   pyo.value(calc_evaporacao_ORT(model)),   dict_emissario.get('Evaporação', 0),        't_H2O/h'),
]:
    _diag_r[_nm] = {'nome': _nm, 'usado': round(_vl or 0, 2), 'limite': round(_lm, 2), 'unidade': _un, 'tipo': 'ORT'}

for _f in lista_fibras_ORT:
    _vl2 = pyo.value(calc_producao_digestor_ORT(model, _f)) or 0
    _lm2 = dict_emissario.get(_f, 0)
    _diag_r[_f] = {'nome': _f, 'usado': round(_vl2, 2), 'limite': round(_lm2, 2), 'unidade': 't/d', 'tipo': 'ORT'}

def _diag_cls(r):
    if r['limite'] <= 0: return 'N/A', 0.0
    ratio = r['usado'] / r['limite']
    if ratio > 1.001:   return 'VIOLADA', round(100 * ratio, 1)
    elif ratio >= 0.95: return 'CRÍTICA', round(100 * ratio, 1)
    elif ratio >= 0.80: return 'ATIVA',   round(100 * ratio, 1)
    else:               return 'FOLGADA', round(100 * ratio, 1)

restricoes_status = {}
for _k, _rv in _diag_r.items():
    _st, _pt = _diag_cls(_rv)
    restricoes_status[_k] = {**_rv, 'status': _st, 'percentual': _pt,
                              'folga': round(_rv['limite'] - _rv['usado'], 2)}

restricoes_criticas = {k: v for k, v in restricoes_status.items()
                       if v['status'] in ('CRÍTICA', 'VIOLADA')}

alocacao_produtos = []
for _p in sorted(model.produtos):
    _dem = sum(dict_demanda.get((_p, _mc), 0) for _mc in model.mercados)
    if _dem == 0: continue
    _prod = sum((pyo.value(model.producao_vendavel[_p, _m, _mc]) or 0)
                for _m in model.maquinas for _mc in model.mercados)
    _mq_us = [f'MP{_m}' for _m in model.maquinas
               if sum((pyo.value(model.producao_vendavel[_p, _m, _mc]) or 0)
                      for _mc in model.mercados) > 0.1]
    _mq_dp = [f'MP{_m}' for _m in model.maquinas if dict_prod_por_MP.get((_m, _p), 0) == 1]
    _gap_p = round(_dem - _prod, 1)
    _rec_p = sum((pyo.value(model.producao_vendavel[_p, _m, _mc]) or 0) * dict_preco.get((_p, _mc), 0)
                 for _m in model.maquinas for _mc in model.mercados)
    _cst_p = sum((pyo.value(model.producao_bruta[_p, _m]) or 0) * dict_custos.get((_p, _m), 0)
                 for _m in model.maquinas)
    alocacao_produtos.append({
        'produto': _p, 'demanda': round(_dem, 1), 'producao': round(_prod, 1), 'gap': _gap_p,
        'maquinas_usadas': _mq_us, 'maquinas_disponiveis': _mq_dp,
        'margem': round(_rec_p - _cst_p, 0),
        'motivo': ('Sem máquina habilitada' if not _mq_dp
                   else 'Atendida' if _gap_p <= 0.1
                   else 'Gap: capacidade insuficiente')
    })

sensibilidade_restricoes = {}
for _k, _rv in restricoes_status.items():
    if _rv['status'] not in ('CRÍTICA', 'ATIVA'): continue
    _top = []
    if _rv['tipo'] == 'maquina':
        _mn = int(_k.replace('Tempo MP', ''))
        _cs = [(_p, round((pyo.value(model.producao_bruta[_p, _mn]) or 0)
                           / dict_produtividade_bruta.get((_mn, _p), 1), 1))
               for _p in model.produtos
               if dict_produtividade_bruta.get((_mn, _p), 0) > 0
               and (pyo.value(model.producao_bruta[_p, _mn]) or 0) > 0.1]
        _top = [{'produto': t[0], 'h': t[1]} for t in sorted(_cs, key=lambda x: -x[1])[:5]]
    sensibilidade_restricoes[_k] = {
        'status': _rv['status'], 'percentual': _rv['percentual'],
        'folga': _rv['folga'], 'unidade': _rv['unidade'],
        'top_consumidores': _top,
        'interpretacao': (f"{_k}: {_rv['percentual']}% do limite "
                          f"({_rv['usado']} / {_rv['limite']} {_rv['unidade']}). "
                          f"Folga: {_rv['folga']} {_rv['unidade']}.")
    }

_oc_h = {_m: (pyo.value(model.capacidade_ociosa[_m]) or 0) for _m in model.maquinas}
_gap_lst = sorted([r for r in alocacao_produtos if r['gap'] > 0.1], key=lambda x: -x['gap'])
trocas_possiveis = []
for _rw in _gap_lst[:10]:
    _p = _rw['produto']
    _cap_mq = {}
    for _m in model.maquinas:
        if dict_prod_por_MP.get((_m, _p), 0) == 0: continue
        _pb = dict_produtividade_bruta.get((_m, _p), 0)
        if _pb <= 0: continue
        _cap_mq[f'MP{_m}'] = {'ociosa_h': round(_oc_h[_m], 1), 'potencial_t': round(_oc_h[_m] * _pb, 1)}
    trocas_possiveis.append({**_rw, 'capacidade_por_maquina': _cap_mq})

_n_crit = len(restricoes_criticas)
_gt = sum(r['gap'] for r in _gap_lst)
_fo_v = pyo.value(model.obj)
_crit_str = ', '.join(f"{k}({v['percentual']}%)" for k, v in restricoes_criticas.items()) or 'nenhuma'
diagnostico_mix = (
    f"OTIMIZADOR DE MIX ORT — Horizonte: {DIAS_PERIODO} dias\n"
    f"FO: R$ {_fo_v:,.0f}\n"
    f"Restrições críticas/violadas ({_n_crit}): {_crit_str}\n"
    f"Demanda não atendida: {_gt:,.1f}t em {len(_gap_lst)} produtos\n"
    f"Máquinas: MP27 (ORT), MP28 (ORT), MC25 (ORT), MC26 (ORT)\n"
    f"Mercados: ME (Externo), MI (Interno), Transferência"
)

razao_alocacao = {}
for _p in sorted(model.produtos):
    _dem = sum(dict_demanda.get((_p, _mc), 0) for _mc in model.mercados)
    if _dem == 0:
        continue
    _mq_dp = [m for m in model.maquinas if dict_prod_por_MP.get((m, _p), 0) == 1]
    if not _mq_dp:
        continue

    _margem_por_maq = {}
    for _m in _mq_dp:
        _rec = sum(
            (pyo.value(model.producao_vendavel[_p, _m, _mc]) or 0) * dict_preco.get((_p, _mc), 0)
            for _mc in model.mercados
        )
        _cst = (pyo.value(model.producao_bruta[_p, _m]) or 0) * dict_custos.get((_p, _m), 0)
        _margem_por_maq[f'MP{_m}'] = round(_rec - _cst, 0)

    _mq_usadas = [
        f'MP{m}' for m in model.maquinas
        if sum((pyo.value(model.producao_vendavel[_p, m, _mc]) or 0)
               for _mc in model.mercados) > 0.1
    ]

    if len(_mq_dp) == 1:
        _razao = 'unica_maquina_habilitada'
    elif not _mq_usadas:
        _razao = 'nao_produzido'
    else:
        _melhor_maq = max(_margem_por_maq, key=_margem_por_maq.get)
        _razao = 'maior_margem' if _melhor_maq in _mq_usadas else 'restricao_de_balanco'

    _impacto_restricoes = {}
    for _m in _mq_dp:
        _coefs = {}
        _esp_ckb_fc = get_consumo_especifico(model, 'ORT', _p, _m, 'CKB-FC')
        _esp_ckb_fl = get_consumo_especifico(model, 'ORT', _p, _m, 'CKB-FL')
        _capt_rate  = dict_param_add_ORT.get('Captação de Água (m3/t)', 0)
        _emis_rate  = dict_param_add_ORT.get('Emissário (m3/t)', 0)
        if _esp_ckb_fc > 1e-9:
            _coefs['CKB_FC_t_t'] = round(_esp_ckb_fc, 6)
        if _esp_ckb_fl > 1e-9:
            _coefs['CKB_FL_t_t'] = round(_esp_ckb_fl, 6)
        if _capt_rate > 1e-9 and _m in [27, 28]:
            _coefs['Captacao_m3_t'] = round(_capt_rate, 4)
        if _emis_rate > 1e-9 and _m in [27, 28]:
            _coefs['Emissario_m3_t'] = round(_emis_rate, 4)
        if _coefs:
            _impacto_restricoes[f'MP{_m}'] = _coefs

    razao_alocacao[_p] = {
        'maquinas_habilitadas': [f'MP{m}' for m in _mq_dp],
        'maquinas_usadas':      _mq_usadas,
        'razao':                _razao,
        'margem_por_maquina':   _margem_por_maq,
        'impacto_restricoes':   _impacto_restricoes,
    }

diagnostico_ia = {
    'diagnostico_mix': diagnostico_mix,
    'restricoes_status': restricoes_status,
    'restricoes_criticas': restricoes_criticas,
    'alocacao_produtos': alocacao_produtos,
    'sensibilidade_restricoes': sensibilidade_restricoes,
    'trocas_possiveis': trocas_possiveis,
    'razao_alocacao': razao_alocacao,
}

# ── waste_data ─────────────────────────────────────────────────────────────
waste_data = []
for (maquina, produto) in sorted(dict_total_waste.keys()):
    if dict_prod_por_MP.get((maquina, produto), 0) == 0:
        continue

    _preco_wf = 0.0
    for _mc in ['ME', 'MI', 'Transferência']:
        _p_mc = dict_preco.get((produto, _mc), 0)
        if _p_mc > 0:
            _preco_wf = _p_mc
            break

    _custo_wf  = dict_custos.get((produto, maquina), 0.0)
    _prod_h_wf = dict_produtividade_bruta.get((maquina, produto), 0.0)

    _sIAC      = dict_IAC.get((maquina, produto), 1.0)
    _sRefugo   = 1.0 - dict_refugo_ajustado.get((maquina, produto), 0.0)
    _sRepExt   = _sRefugo   * (1 - dict_Rep_Externo.get((maquina, produto), 0.0))
    _sMR2      = _sRepExt   * (1 - dict_MR2.get((maquina, produto), 0.0))
    _sSala     = _sMR2      * (1 - dict_Sala_Perdas.get((maquina, produto), 0.0))
    _sCortGram = _sSala     * (1 - dict_Cortadeira_Perda_Gramatura.get((maquina, produto), 0.0))
    _sCortCort = _sCortGram * (1 - dict_Cortadeira_Perdas_Cortadeira.get((maquina, produto), 0.0))
    _sEstExp   = _sCortCort * (1 - dict_Estoque_Perdas.get((maquina, produto), 0.0))
    _sVend     = _sEstExp   * (1 - dict_Estoque_Perdas_Refugo.get((maquina, produto), 0.0))

    waste_data.append({
        "produto":             produto,
        "maquina":             maquina,
        "preco":               round(_preco_wf, 2),
        "custo_variavel":      round(_custo_wf, 2),
        "produtividade_bruta": round(_prod_h_wf, 4),
        "surv_IAC":            round(_sIAC,      6),
        "surv_Refugo":         round(_sRefugo,   6),
        "surv_RepExt":         round(_sRepExt,   6),
        "surv_MR2":            round(_sMR2,      6),
        "surv_Sala":           round(_sSala,     6),
        "surv_CortGram":       round(_sCortGram, 6),
        "surv_CortCort":       round(_sCortCort, 6),
        "surv_EstExp":         round(_sEstExp,   6),
        "surv_Vendavel":       round(_sVend,     6),
        "delta_IAC":           round(1.0       - _sIAC,      6),
        "delta_Refugo":        round(_sIAC     - _sRefugo,   6),
        "delta_RepExt":        round(_sRefugo  - _sRepExt,   6),
        "delta_MR2":           round(_sRepExt  - _sMR2,      6),
        "delta_Sala":          round(_sMR2     - _sSala,     6),
        "delta_CortGram":      round(_sSala    - _sCortGram, 6),
        "delta_CortCort":      round(_sCortGram - _sCortCort,6),
        "delta_EstExp":        round(_sCortCort - _sEstExp,  6),
        "delta_EstRef":        round(_sEstExp  - _sVend,     6),
    })

# ── bridge_data ────────────────────────────────────────────────────────────
_maquinas_lista = [27, 28, 25, 26]

# Capacidade instalada: tempo disponível × maior produtividade habilitada por máquina
_capac_instalada_br = 0.0
for _m_br in _maquinas_lista:
    _prods_hab_br = [p for p in lista_produtos if dict_produtividade_bruta.get((_m_br, p), 0) > 0]
    if not _prods_hab_br:
        continue
    _prod_max_br = max(dict_produtividade_bruta.get((_m_br, p), 0) for p in _prods_hab_br)
    _t_disp_br   = dict_tempo_carga.get(_m_br, 0) * dict_taxa_DISP.get(_m_br, 1)
    _capac_instalada_br += _t_disp_br * _prod_max_br

# Produção bruta real: soma do que o otimizador alocou
_bruta_real_br = sum(
    pyo.value(model.producao_bruta[p, m]) or 0
    for p in lista_produtos for m in _maquinas_lista
)

# Custo de portfólio: horas ociosas × produtividade média das habilitadas
# (igual ao modelo completo mix_attempt1_celulose.py)
_capac_ociosa_br = 0.0
for _m_br in _maquinas_lista:
    _h_oc_br = pyo.value(model.capacidade_ociosa[_m_br]) or 0
    if _h_oc_br <= 0:
        continue
    _prods_hab_br = [p for p in lista_produtos if dict_produtividade_bruta.get((_m_br, p), 0) > 0]
    if not _prods_hab_br:
        continue
    _prod_med_br = sum(dict_produtividade_bruta.get((_m_br, p), 0) for p in _prods_hab_br) / len(_prods_hab_br)
    _capac_ociosa_br += _h_oc_br * _prod_med_br

# Produção vendável real
_vendavel_real_br = sum(
    pyo.value(model.producao_vendavel[p, m, merc]) or 0
    for p in lista_produtos for m in _maquinas_lista for merc in ['ME', 'MI', 'Transferência']
)

# Perda total waste: diferença entre bruto e vendável
_perda_waste_br = max(_bruta_real_br - _vendavel_real_br, 0.0)

# Demanda total e gap: apenas produtos com produtividade cadastrada em ao menos
# uma máquina ORT (flag=1 + produtividade > 0) — exclui produtos de outras
# unidades que têm flag mas não têm coeficiente de produtividade no escopo ORT
_prods_ativos_ort = set(
    p for p in model.produtos
    if any(
        dict_prod_por_MP.get((m, p), 0) == 1
        and dict_produtividade_bruta.get((m, p), 0) > 0
        for m in _maquinas_lista
    )
)
_demanda_total_br = sum(
    v for (p, merc), v in dict_demanda.items()
    if p in _prods_ativos_ort
)

# Produção vendável real apenas dos produtos ativos na ORT
_vendavel_prods_ort_br = sum(
    pyo.value(model.producao_vendavel[p, m, merc]) or 0
    for p in _prods_ativos_ort
    for m in _maquinas_lista
    for merc in model.mercados
)

# Restrição de Fábrica = demanda não atendida de produtos ativos na ORT
# (mesmo conceito do modelo completo: gap_demanda_t vira a barra laranja no bridge)
_gap_demanda_br = sum(
    pyo.value(model.demanda_nao_atendida[p, merc]) or 0
    for p in _prods_ativos_ort
    for merc in model.mercados
)

_receita_total_br = sum(
    (pyo.value(model.producao_vendavel[p, m, merc]) or 0) * dict_preco.get((p, merc), 0)
    for p in model.produtos for m in model.maquinas for merc in model.mercados
)
_preco_medio_br = (_receita_total_br / _vendavel_real_br) if _vendavel_real_br > 0 else 0

_custo_total_br = sum(
    (pyo.value(model.producao_bruta[p, m]) or 0) * dict_custos.get((p, m), 0)
    for p in model.produtos for m in model.maquinas
)
_custo_medio_br = (_custo_total_br / _bruta_real_br) if _bruta_real_br > 0 else 0

bridge_data = {
    "capac_instalada_t":      round(_capac_instalada_br, 1),
    "capac_ociosa_t":         round(_capac_ociosa_br, 1),
    "producao_bruta_real_t":  round(_bruta_real_br, 1),
    "gap_demanda_t":          round(_gap_demanda_br, 1),
    "perda_waste_t":          round(_perda_waste_br, 1),
    "producao_vendavel_t":    round(_vendavel_real_br, 1),
    "demanda_total_t":        round(_demanda_total_br, 1),
    "pct_capac_ociosa":       round(100 * _capac_ociosa_br  / _capac_instalada_br, 2) if _capac_instalada_br > 0 else 0,
    "pct_gap_demanda":        round(100 * _gap_demanda_br   / _capac_instalada_br, 2) if _capac_instalada_br > 0 else 0,
    "pct_waste":              round(100 * _perda_waste_br   / _capac_instalada_br, 2) if _capac_instalada_br > 0 else 0,
    "pct_vendavel":           round(100 * _vendavel_real_br / _capac_instalada_br, 2) if _capac_instalada_br > 0 else 0,
    "preco_medio_ponderado":  round(_preco_medio_br, 2),
    "custo_medio_ponderado":  round(_custo_medio_br, 2),
    "margem_media_ponderada": round(_preco_medio_br - _custo_medio_br, 2),
}

# ── Coleta os resultados para o dashboard ──────────────────────────────────

dashboard_data = {

    "dias_periodo": DIAS_PERIODO,

    # KPI cards
    "fo_reais": pyo.value(model.obj),
    "producao_bruta_total": sum(
        pyo.value(model.producao_bruta[p, m])
        for p in model.produtos for m in model.maquinas
        if pyo.value(model.producao_bruta[p, m]) > 0
    ),
    "producao_vendavel_total": sum(
        pyo.value(model.producao_vendavel[p, m, merc])
        for p in model.produtos for m in model.maquinas for merc in model.mercados
        if pyo.value(model.producao_vendavel[p, m, merc]) > 0
    ),

    # Produção bruta por máquina
    "bruta_por_maquina": {
        str(m): sum(
            pyo.value(model.producao_bruta[p, m])
            for p in model.produtos
        )
        for m in model.maquinas
    },

    # Produção vendável por mercado
    "vendavel_por_mercado": {
        merc: sum(
            pyo.value(model.producao_vendavel[p, m, merc])
            for p in model.produtos for m in model.maquinas
        )
        for merc in model.mercados
    },

    # Cobertura de demanda por produto
    "cobertura_por_produto": [
        {
            "produto": p,
            "mercado": merc,
            "demanda": qtd,
            "atendida": sum(pyo.value(model.producao_vendavel[p, m, merc]) for m in model.maquinas),
            "nao_atendida": pyo.value(model.demanda_nao_atendida[p, merc]),
        }
        for (p, merc), qtd in dict_demanda.items()
        if qtd > 0
    ],

    # Capacidade por máquina
    "capacidade_por_maquina": {
        str(m): {
            "usado": sum(
                pyo.value(model.producao_bruta[p, m]) / dict_produtividade_bruta.get((m, p), 1)
                for p in model.produtos
                if dict_produtividade_bruta.get((m, p), 0) > 0
            ),
            "disponivel": dict_tempo_carga.get(m, 0) * dict_taxa_DISP.get(m, 1),
            "ocioso": pyo.value(model.capacidade_ociosa[m]),
        }
        for m in model.maquinas
    },

    # Balanço de fábrica ORT
    "balanco_ORT": [
        {"nome": "Caustificação",     "usado": pyo.value(calc_caustificacao(model)), "limite": dict_emissario.get("Caustificação", 0)},
        {"nome": "Outorga Captação",  "usado": pyo.value(calc_captacao(model)),      "limite": dict_emissario.get("Outorga Captação", 0)},
        {"nome": "Outorga Emissário", "usado": pyo.value(calc_emissario(model)),     "limite": dict_emissario.get("Outorga Emissario", 0)},
        {"nome": "CDR",               "usado": pyo.value(calc_cdr(model)),           "limite": dict_emissario.get("CDR", 0)},
        {"nome": "Evaporação ORT",    "usado": pyo.value(calc_evaporacao_ORT(model)),"limite": dict_emissario.get("Evaporação", 0)},
        *[
            {"nome": fibra, "usado": pyo.value(calc_producao_digestor_ORT(model, fibra)), "limite": dict_emissario.get(fibra, 0)}
            for fibra in lista_fibras_ORT
        ],
    ],

    # Produção vendável por máquina × produto
    "atendida_por_maquina_produto": {
        p: {
            str(m): round(sum(
                pyo.value(model.producao_vendavel[p, m, merc])
                for merc in model.mercados
            ), 1)
            for m in model.maquinas
        }
        for p in model.produtos
        if sum(
            pyo.value(model.producao_vendavel[p, m, merc])
            for m in model.maquinas
            for merc in model.mercados
        ) > 0.1
    },

    # Máquinas utilizadas por produto × mercado
    "maquinas_por_produto_mercado": [
        {
            "produto": p,
            "mercado": merc,
            "maquinas": {
                str(m): round(pyo.value(model.producao_vendavel[p, m, merc]), 1)
                for m in model.maquinas
                if pyo.value(model.producao_vendavel[p, m, merc]) > 0.1
            }
        }
        for (p, merc), qtd in dict_demanda.items()
        if qtd > 0
    ],

    # Margem de contribuição por máquina
    "margem_por_maquina": {
        str(m): round(
            sum(
                pyo.value(model.producao_vendavel[p, m, merc]) * dict_preco.get((p, merc), 0)
                for p in model.produtos
                for merc in model.mercados
            ) - sum(
                pyo.value(model.producao_bruta[p, m]) * (
                    dict_custos.get((p, m), 0)
                    + flag_remuneracao_celulose * (
                        get_consumo_especifico(model, 'ORT', p, m, 'CKB-FC') * (preco_venda_fibra_curta - custo_variavel_fibra_curta)
                        + get_consumo_especifico(model, 'ORT', p, m, 'CKB-FL') * (preco_venda_fibra_longa - custo_variavel_fibra_longa)
                    )
                )
                for p in model.produtos
            ),
            0
        )
        for m in model.maquinas
    },

    # Capacidade máxima bruta por máquina
    # = tempo disponível × maior produtividade habilitada (mesma base do bridge)
    "prod_bruta_max_por_maquina": {
        str(m): round(
            dict_tempo_carga.get(m, 0) * dict_taxa_DISP.get(m, 1) *
            max((dict_produtividade_bruta.get((m, p), 0)
                 for p in lista_produtos
                 if dict_produtividade_bruta.get((m, p), 0) > 0), default=0),
            1
        )
        for m in model.maquinas
    },
    **diagnostico_ia,

    "waste_data": waste_data,
    "bridge_data": bridge_data,
}

# Grava o JSON
with open("dashboard_data_ort.json", "w", encoding="utf-8") as f:
    json.dump(dashboard_data, f, ensure_ascii=False, indent=2)

print("\n✓ dashboard_data_ort.json gerado — abra o dashboard ORT no navegador")


