"""
Definição do esquema (schema) do banco de dados.
---------------------------------------------------------------------------
Duas tabelas:

- cargas: histórico/auditoria de cada importação de planilha realizada
  (quando, qual arquivo, quantos registros). Nunca é apagada.
- dividas: dados vigentes das dívidas, sempre referentes à última carga
  importada. É totalmente substituída a cada nova importação, pois cada
  planilha enviada já contém o histórico completo (evitando duplicidade).
"""

SCHEMA_SQL = """
CREATE TABLE IF NOT EXISTS cargas (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    data_carga      TEXT    NOT NULL,
    origem_arquivo  TEXT,
    qtd_registros   INTEGER NOT NULL DEFAULT 0,
    backup_arquivo  TEXT
);

CREATE TABLE IF NOT EXISTS dividas (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    nome_cliente    TEXT    NOT NULL,
    valor_divida    REAL    NOT NULL DEFAULT 0,
    valor_pago      REAL    NOT NULL DEFAULT 0,
    valor_restante  REAL    NOT NULL DEFAULT 0,
    termino_divida  TEXT,
    quitado         TEXT    NOT NULL CHECK (quitado IN ('SIM', 'NAO')),
    carga_id        INTEGER NOT NULL REFERENCES cargas(id) ON DELETE CASCADE
);

CREATE INDEX IF NOT EXISTS idx_dividas_carga_id ON dividas (carga_id);
CREATE INDEX IF NOT EXISTS idx_dividas_quitado ON dividas (quitado);
CREATE INDEX IF NOT EXISTS idx_dividas_cliente ON dividas (nome_cliente);
"""
