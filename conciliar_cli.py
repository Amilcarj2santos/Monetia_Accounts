"""
Uso por linha de comandos (sem interface web):

    python conciliar_cli.py entrada.xlsx [saida.xlsx]

Gera o Excel com as colunas CONCILIADO e INFORMAÇÃO.
"""
import sys
import excel_io
import reconciliation as rec


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    entrada = sys.argv[1]
    saida = sys.argv[2] if len(sys.argv) > 2 else \
        entrada.rsplit(".", 1)[0] + "_CONCILIADO.xlsx"

    with open(entrada, "rb") as fh:
        data = fh.read()

    out_bytes, resultado, info = excel_io.procesar_workbook(data, cfg=rec.ReconConfig())

    with open(saida, "wb") as fh:
        fh.write(out_bytes)

    r = resultado.resumen
    print(f"Colunas: importes={info['col_valor']} | data={info['col_fecha']}")
    print(f"Linhas: {r['lineas_total']} | conciliadas: {r['lineas_casadas']} | "
          f"soltas: {r['lineas_sueltas']} | grupos: {r['grupos']}")
    print(f"Saldo em aberto: {r['saldo_abierto_eur']} €")
    print(f"Guardado: {saida}")


if __name__ == "__main__":
    main()
