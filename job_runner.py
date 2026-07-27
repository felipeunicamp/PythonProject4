"""
job_runner.py
Dispara os scripts de otimização como subprocessos.
Captura apenas status (inicio, fim, erro) — não faz streaming de stdout.
"""

import subprocess
import sys
import threading
from datetime import datetime

# Estado global do job (single-user, local)
_job_state = {
    "rodando": False,
    "ultimo_job": None,       # "mix" | "sequenciamento"
    "ultimo_status": None,    # "ok" | "erro"
    "ultima_mensagem": "",
    "ultimo_fim": None,       # datetime string
    "mix_concluido": False,   # habilita botão de sequenciamento
    "seq_concluido": False,
}
_lock = threading.Lock()


def get_state() -> dict:
    with _lock:
        return dict(_job_state)


def _run(script_path: str, job_name: str):
    with _lock:
        _job_state["rodando"] = True
        _job_state["ultimo_job"] = job_name
        _job_state["ultimo_status"] = None
        _job_state["ultima_mensagem"] = f"{job_name} iniciado"

    try:
        result = subprocess.run(
            [sys.executable, script_path],
            capture_output=True,
            text=True,
            encoding="utf-8",
        )
        with _lock:
            if result.returncode == 0:
                _job_state["ultimo_status"] = "ok"
                _job_state["ultima_mensagem"] = f"{job_name} concluído com sucesso"
                if job_name == "mix":
                    _job_state["mix_concluido"] = True
                elif job_name == "sequenciamento":
                    _job_state["seq_concluido"] = True
            else:
                _job_state["ultimo_status"] = "erro"
                # Pega apenas as últimas 20 linhas do stderr para exibir
                stderr_lines = (result.stderr or "").strip().splitlines()
                resumo = "\n".join(stderr_lines[-20:])
                _job_state["ultima_mensagem"] = f"Erro em {job_name}:\n{resumo}"
    except Exception as e:
        with _lock:
            _job_state["ultimo_status"] = "erro"
            _job_state["ultima_mensagem"] = f"Exceção ao rodar {job_name}: {str(e)}"
    finally:
        with _lock:
            _job_state["rodando"] = False
            _job_state["ultimo_fim"] = datetime.now().strftime("%d/%m/%Y %H:%M:%S")


def rodar_mix():
    """Dispara mix_ort_30dias.py em thread separada."""
    with _lock:
        if _job_state["rodando"]:
            return False, "Já há um job em execução"
    t = threading.Thread(target=_run, args=("mix_ort_30dias.py", "mix"), daemon=True)
    t.start()
    return True, "Mix iniciado"


def rodar_sequenciamento():
    """Dispara sequenciamento_30dias_ort.py em thread separada."""
    with _lock:
        if _job_state["rodando"]:
            return False, "Já há um job em execução"
        if not _job_state["mix_concluido"]:
            return False, "Execute o Mix antes do Sequenciamento"
    t = threading.Thread(target=_run, args=("sequenciamento_30dias_ort.py", "sequenciamento"), daemon=True)
    t.start()
    return True, "Sequenciamento iniciado"
