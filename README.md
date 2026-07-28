# Monetia Accounts — Conciliador de Conta Corrente

Aplicação web que analisa uma **conta corrente de um fornecedor** em Excel e
encontra as linhas cujo valor a **débito/crédito soma zero** — ou seja, uma
fatura que já tem o seu pagamento (incluindo notas de crédito). Marca essas
linhas como conciliadas e deixa visíveis as **linhas soltas**, que são as que
compõem o saldo da conta nesse momento.

Faz uma procura lógica de faturas, pagamentos e notas de crédito, e também
encontra **um pagamento que salda várias faturas** (ou várias faturas/pagamentos
que juntos dão zero).

## O que faz, em detalhe

- Procura as colunas pelo **nome**, não pela posição (a ordem pode variar):
  - `DEB_CRED` → importes com sinal (positivos e negativos).
  - `DATA_CERTA` → data usada para validar que um pagamento faz sentido para
    uma fatura.
- Casa as linhas em duas fases:
  1. **1:1** — uma fatura com o seu pagamento do mesmo importe. Se houver
     vários candidatos com o mesmo importe, escolhe o de **data mais próxima**.
  2. **1:várias / várias:1** — um pagamento que salda várias faturas (ou o
     contrário), procurando o subconjunto de linhas que soma zero.
- Tolerância de arredondamento configurável (por defeito **± 0,01 €**).
- Devolve o **mesmo Excel** com duas colunas novas à direita:
  - `CONCILIADO` → **VERDADEIRO / FALSO** (booleano; no Excel pode convertê-lo
    numa caixa de verificação clicável).
  - `INFORMAÇÃO` → `Referência 001`, `Referência 002`, … (a mesma referência
    para todas as linhas do grupo). Se a linha não casa com nada →
    `SEM IDENTIFICAR`.

## Como usar

### Opção A — App web (recomendado)

```bash
pip install -r requirements.txt
streamlit run app.py
```

Abre no navegador, carregue o Excel, veja o resultado e descarregue o ficheiro
conciliado.

### Opção B — Linha de comandos

```bash
pip install -r requirements.txt
python conciliar_cli.py a_minha_conta.xlsx
```

Gera `a_minha_conta_CONCILIADO.xlsx`.

## Ficheiro de demonstração

`demo_conta_corrente.xlsx` contém um exemplo (com as colunas propositadamente
desordenadas) para testar. Reproduz o caso real:

| Grupo | Linhas | Soma |
|-------|--------|------|
| Referência 001 | 992,63 + (−992,63) | 0 |
| Referência 002 | −2 108,04 + 1 079,02 + 1 029,02 | 0 |
| Referência 003 | −1 500,00 + 200,00 + 1 300,00 (com nota de crédito) | 0 |
| SEM IDENTIFICAR | 992,63 (fica solta) | — |

## Estrutura do projeto

```
reconciliation.py   # motor de conciliação (deteção de colunas + casação)
excel_io.py         # lê o Excel e escreve as colunas de saída
app.py              # interface web (Streamlit)
conciliar_cli.py    # uso por linha de comandos
generar_demo.py     # gera o Excel de demonstração
probar.py           # teste rápido sobre o demo
requirements.txt
```

## Notas

- As caixas de verificação **nativas** e clicáveis do Excel são uma função
  recente; por isso o `CONCILIADO` sai como valor **VERDADEIRO/FALSO**. Para as
  tornar clicáveis: selecione a coluna → **Inserir → Caixa de verificação**.
- Só participam na conciliação as linhas com importe diferente de zero.
