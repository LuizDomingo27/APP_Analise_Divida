"""
Sincronização do banco SQLite com o GitHub.
---------------------------------------------------------------------------
Por que isso existe?
    No Streamlit Community Cloud o disco NÃO é persistente entre reinícios
    ("sonos") do app. O banco data/dividas.db é gravado em disco local e,
    quando o app reinicia, esse arquivo volta ao estado que está no
    repositório. Se uma planilha foi importada pelo app na nuvem e não foi
    commitada de volta, os dados se perdem.

    Este módulo resolve isso de forma automática usando a GitHub Contents
    API + um Personal Access Token (PAT). Ele não depende do `git` estar
    instalado na máquina — fala direto com a API REST do GitHub:

        - push_db():  envia o data/dividas.db atual para o repositório
                      (um commit), chamado após cada importação.
        - pull_db():  baixa o data/dividas.db do repositório para o disco
                      local, usado na inicialização para reidratar os dados.

Configuração (NUNCA coloque o token no código):
    Defina em .streamlit/secrets.toml (local) ou em Settings > Secrets
    (Streamlit Cloud):

        [github]
        token  = "ghp_xxxxxxxxxxxxxxxxxxxx"   # PAT com permissão de Contents (write)
        repo   = "LuizDomingo27/APP_Analise_Divida"
        branch = "main"                        # opcional, padrão "main"
        path   = "data/dividas.db"             # opcional, padrão "data/dividas.db"

    O PAT precisa de escopo:
        - Fine-grained: permissão "Contents: Read and write" no repositório.
        - Classic: escopo "repo".
"""

import base64
import os

import requests
import streamlit as st

from database.connection import DB_PATH

API_BASE = "https://api.github.com"
TIMEOUT = 30


# ---------------------------------------------------------------------------
# Configuração
# ---------------------------------------------------------------------------
def _config() -> dict | None:
    """
    Lê a configuração de sincronização de st.secrets["github"].
    Retorna None (sync desativado) se o token/repo não estiverem definidos —
    assim o app continua funcionando normalmente sem a integração.
    """
    try:
        gh = st.secrets["github"]
    except (KeyError, FileNotFoundError):
        return None

    token = gh.get("token")
    repo = gh.get("repo")
    if not token or not repo:
        return None

    return {
        "token": token,
        "repo": repo,
        "branch": gh.get("branch", "main"),
        "path": gh.get("path", "data/dividas.db"),
    }


def sync_disponivel() -> bool:
    """True se a integração com o GitHub está configurada."""
    return _config() is not None


def _headers(token: str) -> dict:
    return {
        "Authorization": f"Bearer {token}",
        "Accept": "application/vnd.github+json",
        "X-GitHub-Api-Version": "2022-11-28",
    }


def _contents_url(cfg: dict) -> str:
    return f"{API_BASE}/repos/{cfg['repo']}/contents/{cfg['path']}"


# ---------------------------------------------------------------------------
# Estado remoto
# ---------------------------------------------------------------------------
def _obter_arquivo_remoto(cfg: dict) -> dict | None:
    """
    Retorna {"sha": ..., "content": <bytes>} do arquivo no GitHub, ou None se
    ele ainda não existe no repositório (primeiro envio).
    """
    resp = requests.get(
        _contents_url(cfg),
        headers=_headers(cfg["token"]),
        params={"ref": cfg["branch"]},
        timeout=TIMEOUT,
    )
    if resp.status_code == 404:
        return None
    resp.raise_for_status()
    data = resp.json()
    conteudo = base64.b64decode(data["content"]) if data.get("content") else b""
    return {"sha": data["sha"], "content": conteudo}


# ---------------------------------------------------------------------------
# Push — envia o banco local para o GitHub
# ---------------------------------------------------------------------------
def push_db(mensagem: str | None = None) -> dict:
    """
    Envia o data/dividas.db local para o repositório como um novo commit.

    Retorna:
        {"status": "ok" | "skipped" | "disabled", "detalhe": str}

    - "disabled": integração não configurada (silencioso, não é erro).
    - "skipped":  o conteúdo remoto já é idêntico ao local (nada a fazer).
    - "ok":       commit criado/atualizado com sucesso.

    Levanta exceção (requests.HTTPError) se a API do GitHub recusar a
    operação (ex.: token inválido ou sem permissão) — quem chama decide
    como mostrar isso ao usuário.
    """
    cfg = _config()
    if cfg is None:
        return {"status": "disabled", "detalhe": "Integração com GitHub não configurada."}

    if not os.path.exists(DB_PATH):
        return {"status": "skipped", "detalhe": "Banco local ainda não existe."}

    with open(DB_PATH, "rb") as f:
        conteudo_local = f.read()

    remoto = _obter_arquivo_remoto(cfg)
    if remoto is not None and remoto["content"] == conteudo_local:
        return {"status": "skipped", "detalhe": "O banco no GitHub já está atualizado."}

    from datetime import datetime
    msg = mensagem or f"Atualiza banco de dívidas via app ({datetime.now():%Y-%m-%d %H:%M:%S})"

    payload = {
        "message": msg,
        "content": base64.b64encode(conteudo_local).decode("ascii"),
        "branch": cfg["branch"],
    }
    if remoto is not None:
        payload["sha"] = remoto["sha"]  # obrigatório para atualizar arquivo existente

    resp = requests.put(
        _contents_url(cfg),
        headers=_headers(cfg["token"]),
        json=payload,
        timeout=TIMEOUT,
    )
    resp.raise_for_status()
    return {"status": "ok", "detalhe": "Banco enviado ao GitHub com sucesso."}


# ---------------------------------------------------------------------------
# Pull — traz o banco do GitHub para o disco local
# ---------------------------------------------------------------------------
def pull_db(sobrescrever: bool = False) -> dict:
    """
    Baixa o data/dividas.db do repositório para o disco local.

    Args:
        sobrescrever: se False (padrão), só grava quando NÃO existe banco
            local — comportamento seguro para reidratar após o app dormir
            sem apagar dados recém-importados. Se True, sempre substitui o
            arquivo local pelo conteúdo remoto.

    Retorna:
        {"status": "ok" | "skipped" | "disabled" | "vazio", "detalhe": str}
    """
    cfg = _config()
    if cfg is None:
        return {"status": "disabled", "detalhe": "Integração com GitHub não configurada."}

    if not sobrescrever and os.path.exists(DB_PATH) and os.path.getsize(DB_PATH) > 0:
        return {"status": "skipped", "detalhe": "Banco local já presente; pull ignorado."}

    remoto = _obter_arquivo_remoto(cfg)
    if remoto is None or not remoto["content"]:
        return {"status": "vazio", "detalhe": "Não há banco no GitHub para baixar."}

    os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
    with open(DB_PATH, "wb") as f:
        f.write(remoto["content"])
    return {"status": "ok", "detalhe": "Banco restaurado a partir do GitHub."}
