"""Prueba el motor sobre el Excel de demostracion y muestra el resultado."""
import excel_io
import reconciliation as rec

with open("demo_conta_corrente.xlsx", "rb") as fh:
    data = fh.read()

out_bytes, resultado, info = excel_io.procesar_workbook(data)

print("== Columnas detectadas ==")
print("  valor:", info["col_valor"], "| fecha:", info["col_fecha"],
      "| fila cabecera:", info["fila_cabecera"])
print("\n== Resumen ==")
for k, v in resultado.resumen.items():
    print(f"  {k}: {v}")

print("\n== Grupos casados ==")
for g in resultado.grupos:
    print(f"  {g['referencia']}: filas {g['rows']}  ({g['tipo']})")

print("\n== Asignacion por fila ==")
for row in sorted(resultado.asignacion):
    a = resultado.asignacion[row]
    print(f"  fila {row}: conciliado={a['conciliado']!s:5}  {a['referencia']}")

with open("demo_conta_corrente_RESULTADO.xlsx", "wb") as fh:
    fh.write(out_bytes)
print("\nGuardado: demo_conta_corrente_RESULTADO.xlsx")
