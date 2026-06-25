"""
Conexão com o banco de dados SQLite.
---------------------------------------------------------------------------
Centraliza o caminho do arquivo .db e a criação de conexões, garantindo
configurações consistentes (foreign keys, modo WAL) em toda a aplicação.
"""

import os
import sqlite3
from contextlib import contextmanager

# ---------------------------------------------------------------------------
# Caminho do banco de dados
# ---------------------------------------------------------------------------
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_DIR = os.path.join(BASE_DIR, "data")
BACKUP_DIR = os.path.join(DATA_DIR, "backups")
DB_PATH = os.path.join(DATA_DIR, "dividas.db")

os.makedirs(DATA_DIR, exist_ok=True)
os.makedirs(BACKUP_DIR, exist_ok=True)


@contextmanager
def get_connection():
    """
    Fornece uma conexão SQLite configurada (foreign keys ativas, WAL),
    com fechamento garantido ao final do bloco `with`.
    """
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.execute("PRAGMA foreign_keys = ON;")
    conn.execute("PRAGMA journal_mode = WAL;")
    try:
        yield conn
    finally:
        conn.close()
