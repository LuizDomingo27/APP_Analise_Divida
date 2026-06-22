# 💧 Gestão de Dívidas de Fornecedores

Aplicação em **Python + Streamlit** para acompanhamento das dívidas com clientes/fornecedores,
desenvolvida a partir da planilha `Contas_a_receber_23_02_2026_-_JOSIVAN.xlsx` (aba `RESUMO`).

## O que o app mostra

- **Valor total em dívida** — soma de todas as dívidas (coluna `VALOR DIVIDA`)
- **Valor já pago** — soma dos valores já recebidos (coluna `VALOR PAGO`)
- **Valor restante a receber** — soma do saldo devedor (coluna `VALOR RESTANTE`)
- **Quantas dívidas já foram quitadas** — registros com `QUITADO = SIM`
- **Quantas dívidas ainda estão ativas** — registros com `QUITADO = NÃO`
- **% pago do total** — percentual já recuperado em relação ao total devido

Esses indicadores aparecem em **cards** no topo da página, seguidos de uma **tabela detalhada**
(com barra de progresso de pagamento e status visual por cliente), e por fim botões para
**exportar os dados filtrados em Excel (.xlsx) ou PDF**, mantendo o mesmo padrão visual
(verde neon água) em ambos os formatos.

## Como instalar e executar

1. Tenha o Python 3.10+ instalado.
2. Instale as dependências:
   ```bash
   pip install -r requirements.txt
   ```
3. Execute o app:
   ```bash
   streamlit run app.py
   ```
4. O navegador abrirá automaticamente em `http://localhost:8501`.

## Atualizando os dados

Por padrão, o app carrega o arquivo em `data/Contas_a_receber.xlsx` (aba `RESUMO`).
Para usar uma planilha mais recente, basta enviar o novo arquivo `.xlsx` pelo
campo de upload na barra lateral — desde que a aba `RESUMO` mantenha as colunas:

```
Nome do cliente | VALOR DIVIDA | VALOR PAGO | VALOR RESTANTE | TERMINO DA DIVIDA | QUITADO
```

A coluna `QUITADO` deve conter `SIM` (dívida paga/finalizada) ou `NÃO` (dívida ainda ativa).

## Estrutura do projeto

```
divida_app/
├── app.py                      # aplicação principal
├── requirements.txt            # dependências
├── .streamlit/config.toml      # tema Light + cor neon água
├── data/Contas_a_receber.xlsx  # base de dados padrão
└── README.md
```
