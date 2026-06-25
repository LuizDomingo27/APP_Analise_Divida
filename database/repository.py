"""
Repositório de dados — Dívidas de Fornecedores
---------------------------------------------------------------------------
Camada que isola todo o acesso ao SQLite. O restante da aplicação (app.py)
não executa SQL diretamente, apenas chama as funções deste módulo.

Regra de negócio importante:
    Cada planilha enviada contém o HISTÓRICO COMPLETO das dívidas (não é
    incremental). Por isso, a cada nova importação, os dados antigos da
    tabela `dividas` são totalmente substituídos pelos novos — evitando
    duplicidade. Antes de apagar, um backup em CSV é salvo automaticamente
    em data/backups/, e a operação é feita dentro de uma transação (se
    algo falhar no meio do processo, nada é alterado).
"""

import os
import sqlite3
from datetime import datetime

import pandas as pd

from .connection import get_connection, BACKUP_DIR
from .schema import SCHEMA_SQL

# Colunas no "formato app" (iguais às usadas pelo app.py / planilha original)
COLUNAS_APP = [
    "Nome do cliente", "VALOR DIVIDA", "VALOR PAGO",
    "VALOR RESTANTE", "TERMINO DA DIVIDA", "QUITADO",
]

# Mapeamento entre colunas da planilha/app e colunas do banco
MAPA_APP_PARA_BANCO = {
    "Nome do cliente": "nome_cliente",
    "VALOR DIVIDA": "valor_divida",
    "VALOR PAGO": "valor_pago",
    "VALOR RESTANTE": "valor_restante",
    "TERMINO DA DIVIDA": "termino_divida",
    "QUITADO": "quitado",
}
MAPA_BANCO_PARA_APP = {v: k for k, v in MAPA_APP_PARA_BANCO.items()}


# ---------------------------------------------------------------------------
# Inicialização
# ---------------------------------------------------------------------------
def init_db() -> None:
    """Cria as tabelas do banco, caso ainda não existam. Idempotente."""
    with get_connection() as conn:
        conn.executescript(SCHEMA_SQL)
        conn.commit()


def banco_esta_vazio() -> bool:
    """Retorna True se a tabela `dividas` não possui nenhum registro."""
    with get_connection() as conn:
        cur = conn.execute("SELECT COUNT(*) FROM dividas;")
        total = cur.fetchone()[0]
    return total == 0


# ---------------------------------------------------------------------------
# Escrita — substituição completa dos dados
# ---------------------------------------------------------------------------
def _normalizar_quitado(serie: pd.Series) -> pd.Series:
    serie = serie.astype(str).str.strip().str.upper()
    return serie.replace({"NÃO": "NAO", "SIM": "SIM"})


def _serializar_termino(valor) -> str:
    if pd.isna(valor):
        return ""
    if isinstance(valor, (pd.Timestamp, datetime)):
        return valor.strftime("%Y-%m-%d")
    return str(valor)


def _fazer_backup_atual(conn: sqlite3.Connection) -> str | None:
    """
    Exporta os dados vigentes (antes de serem apagados) para um CSV em
    data/backups/. Retorna o nome do arquivo gerado, ou None se a tabela
    já estava vazia (nada para guardar, ex: primeira carga).
    """
    df_atual = pd.read_sql_query("SELECT * FROM dividas;", conn)
    if df_atual.empty:
        return None

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    nome_arquivo = f"dividas_backup_{timestamp}.csv"
    caminho = os.path.join(BACKUP_DIR, nome_arquivo)
    df_atual.to_csv(caminho, index=False, encoding="utf-8-sig")
    return nome_arquivo


def substituir_dados(df: pd.DataFrame, origem_arquivo: str) -> dict:
    """
    Substitui TODOS os dados vigentes da tabela `dividas` pelos dados do
    DataFrame informado, dentro de uma única transação:

        1. Faz backup (CSV) dos dados atuais.
        2. Registra uma nova "carga" (auditoria: data, origem, qtd. de
           registros, nome do backup).
        3. Apaga os dados antigos da tabela `dividas`.
        4. Insere os novos dados, vinculados à nova carga.

    Se qualquer etapa falhar, a transação é revertida (rollback) e o
    banco permanece exatamente como estava antes da chamada.

    Retorna um dicionário com informações da carga realizada.
    """
    faltantes = [c for c in COLUNAS_APP if c not in df.columns]
    if faltantes:
        raise ValueError(
            "Os dados não contêm as colunas esperadas: " + ", ".join(faltantes)
        )

    df_norm = df.copy()
    df_norm["QUITADO"] = _normalizar_quitado(df_norm["QUITADO"])
    df_norm["TERMINO DA DIVIDA"] = df_norm["TERMINO DA DIVIDA"].apply(_serializar_termino)
    for col in ["VALOR DIVIDA", "VALOR PAGO", "VALOR RESTANTE"]:
        df_norm[col] = pd.to_numeric(df_norm[col], errors="coerce").fillna(0.0)

    df_banco = df_norm.rename(columns=MAPA_APP_PARA_BANCO)[list(MAPA_BANCO_PARA_APP.keys())]

    with get_connection() as conn:
        try:
            backup_nome = _fazer_backup_atual(conn)

            cur = conn.execute(
                "INSERT INTO cargas (data_carga, origem_arquivo, qtd_registros, backup_arquivo) "
                "VALUES (?, ?, ?, ?);",
                (
                    datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    origem_arquivo,
                    len(df_banco),
                    backup_nome,
                ),
            )
            carga_id = cur.lastrowid

            conn.execute("DELETE FROM dividas;")

            df_banco["carga_id"] = carga_id
            df_banco.to_sql("dividas", conn, if_exists="append", index=False)

            conn.commit()
        except Exception:
            conn.rollback()
            raise

    return {
        "carga_id": carga_id,
        "qtd_registros": len(df_banco),
        "backup_arquivo": backup_nome,
    }


# ---------------------------------------------------------------------------
# Leitura
# ---------------------------------------------------------------------------
def obter_dividas() -> pd.DataFrame:
    """Retorna os dados vigentes da tabela `dividas`, já no formato do app."""
    with get_connection() as conn:
        df = pd.read_sql_query(
            "SELECT nome_cliente, valor_divida, valor_pago, valor_restante, "
            "termino_divida, quitado FROM dividas;",
            conn,
        )
    df = df.rename(columns=MAPA_BANCO_PARA_APP)
    if not df.empty:
        df["QUITADO"] = df["QUITADO"].replace({"NAO": "NÃO"})
        df["TERMINO DA DIVIDA"] = pd.to_datetime(
            df["TERMINO DA DIVIDA"], errors="coerce"
        )
    return df


def obter_historico_cargas() -> pd.DataFrame:
    """Retorna o histórico completo de importações (auditoria), mais recente primeiro."""
    with get_connection() as conn:
        df = pd.read_sql_query(
            "SELECT id, data_carga, origem_arquivo, qtd_registros, backup_arquivo "
            "FROM cargas ORDER BY id DESC;",
            conn,
        )
    return df


def obter_ultima_carga() -> dict | None:
    """Retorna informações da carga mais recente, ou None se nunca houve carga."""
    with get_connection() as conn:
        cur = conn.execute(
            "SELECT id, data_carga, origem_arquivo, qtd_registros "
            "FROM cargas ORDER BY id DESC LIMIT 1;"
        )
        row = cur.fetchone()
    if row is None:
        return None
    return {
        "id": row[0],
        "data_carga": row[1],
        "origem_arquivo": row[2],
        "qtd_registros": row[3],
    }
