# Ordenar la hoja "Calculo RE545" — notas

**Estado: aplicado** (BITACORA_BESS 2026-11-03 (11)).

Pedido del usuario (2026-11-03): mismo trabajo que E Costos
(`notas_orden_ecostos.md`). Formato: fila 1 vacía, "Cuadro N°" en la 2,
nombres en la 3, margen en A, una columna entre cuadros, panel en C:H, B:H sin
comentario, unidades entre corchetes.

## Propuesta (en revisión con el usuario)

Orden del cálculo: info traída + reservas de Subastas (× FD × FMA) → energía
de reserva de cada cuarto (AU) → tope por central + ventana (Edisp_T, tabla
resumen) → C1 (curva CMg descendente y asignación) → C2 (inyección real
cronológica + última hora) → monto (ΣC1 − ΣC2) repartido según AU.

1. **Información y reservas:** Mes … Configuracion, Barra, Descarga, Carga,
   SoC, Ciclo de Carga, Ventana de valorización, CMg, CMg Promedio, Adj SSCC ·
   Reservas CPF/CSF ± · FD CPF/CSF ± · FMA CPF/CSF ± · SUMA Reservas×FD×FMA
   (AU) · EiniT (U) · EalmT (V).
2. **Componente 1:** ranking cmg (S), Orden (BI), Periodo (BJ), Curva Cmg
   Decendente (BM), Edisp_Asig (BN), Total C1 (BO).
3. **Componente 2:** BS, BT, BQ, BU, BV, BW, BX, BZ, CA, Total C2 (CC).
4. **Monto a compensar** (CE), verde, totales del mes en la fila 1 sobre C1,
   C2 y Monto.
5. **Resumen por central y ventana** (AW:BG, otro largo), al final.

Quitar de la hoja (se siguen calculando):
- CTF ×6 (AE, AH, AK, AN, AQ, AT): verificar antes que su aporte a AU es 0.
- M, N, O (no las usa nada).
- Propuestas, pendientes de respuesta: BK "Curva Cmg Decendente promedio
  horario" (solo se muestra desde la observación 4) y BY "Energía ya
  Asignada ultima hora" (no la usa ninguna columna).

Fórmula de Excel (misma fila): AU, BO = BN×BM, BQ = BS+BN, BV, BW = I+J,
CA = BU+BZ, CC = CA×BM. Comentario: lo demás (mira otras filas o la tabla
resumen) y lo traído de otras hojas.

## Dudas resueltas (usuario: "1. quitalas, 2. si, 3. si")

1. BK y BY salen de la hoja.
2. Cuadro 5 al final; su primera columna se llama "Central".
3. Renombradas: "Reserva con FD y FMA [kWh]" (AU), "Energía asignada [kWh]"
   (BN), "Total C1" / "Total C2", prefijos "Reserva / FD / FMA" en los tres
   bloques (antes por encabezado de grupo).
