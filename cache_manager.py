"""
cache_manager.py
Gerencia o user_cache.json — overrides do usuário que substituem o Excel.
Nunca modifica o Excel original.
"""

import json
import os
from pathlib import Path

CACHE_PATH = Path("user_cache.json")

SECOES_VALIDAS = {
    "capacidades",   # dict {nome_area: valor_msr}
    "demanda",       # list [{Mercado, Produto, Quantidade (TO), Preço}]
    "custos",        # list [{Produto, Máquina, Custo Variavel}]
    "parametros",    # dict {Custo Variavel Celulose FC, ...}
    "mix_params",    # dict {DIAS_PERIODO, BASE_DEMANDA_DIAS}
    "seq_params",    # dict {DIAS_PERIODO, LOTE_T, ...}
    "sessao",        # dict {persistir: bool}
}


def _load_raw() -> dict:
    if CACHE_PATH.exists():
        try:
            with open(CACHE_PATH, encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}


def _save_raw(data: dict):
    with open(CACHE_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def load_cache() -> dict:
    return _load_raw()


def get_secao(secao: str) -> dict | list | None:
    return _load_raw().get(secao)


def set_secao(secao: str, valor):
    """Salva uma seção inteira do cache."""
    data = _load_raw()
    data[secao] = valor
    _save_raw(data)


def reset_secao(secao: str):
    """Apaga apenas uma seção do cache, mantendo as demais."""
    data = _load_raw()
    if secao in data:
        del data[secao]
        _save_raw(data)


def reset_tudo():
    """Apaga o cache inteiro."""
    if CACHE_PATH.exists():
        CACHE_PATH.unlink()


def deve_persistir() -> bool:
    data = _load_raw()
    return data.get("sessao", {}).get("persistir", False)


def limpar_se_nao_persistir():
    """
    Chamado ao iniciar o servidor.
    Apaga o cache apenas se:
      1. O usuário escolheu não persistir (persistir=False), E
      2. O processo atual NÃO é o reloader filho do Flask
         (evita dupla execução com debug=True).
    """
    import os
    # Flask debug reloader: o processo filho tem a variável WERKZEUG_RUN_MAIN=true
    # Só limpa no processo PAI (antes do reloader iniciar o filho)
    if os.environ.get("WERKZEUG_RUN_MAIN") == "true":
        return
    if not deve_persistir():
        reset_tudo()
