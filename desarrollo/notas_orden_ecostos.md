# Ordenar la hoja "Calculo E Costos" — notas

Pedido del usuario (2026-11-03): ir anotando y programar todo junto.
**Estado: aplicado** (ver BITACORA_BESS 2026-11-03 (8)).

## Aplicado

1. Fila 1 vacía; "Cuadro N° …" en la fila 2; nombres de columna en la 3.
2. Sin encabezados de grupo (se quitaron "Dia", "Nombre", "BESS" y el resto:
   el título del cuadro hace ese papel; prorratas y FD llevan el prefijo en el
   nombre: "Prorrata CPF(-)", "FD CPF(-)", …).
3. Sin las 4 columnas CTF (siempre 0).
4. Sin las 5 del método anterior que no llegan al monto: SoC sobre el minimo,
   Energía SSCC (-)/(+) por remunerar (sin FD), Energía descargada / cargada.
   Se siguen calculando; solo no se escriben.
5. Tres cuadros, en el orden del cálculo:
   - **Cuadro 1 — Información y energía con FD-prorrata** (B:Y): Mes … CMg,
     Adj SSCC, Prorratas, FD, Descarga / Carga ajustada FD-prorrata.
   - **Cuadro 2 — Componente 1** (AA:AM): ranking cmg, bloque ordenado,
     curvas monótonas, energía ajustada acumulada, asignación, valorización,
     Total C1.
   - **Cuadro 3 — Componente 2 y monto a compensar** (AO:AR): Valorizacion
     Descarga / Carga, Total C2, Monto a compensar.
6. Comentario con el origen (hoja, cuadro, columna y condición) en el nombre
   de cada columna traída de otra hoja, y en Asignación Descarga / Carga, que
   usan Pmax de Centrales sin que esté como columna.
7. Panel inmovilizado en I4: quedan fijas las columnas hasta la H
   (Dia … Configuracion; B "Mes" también, por estar a la izquierda) y las
   filas 1–3.

## Segunda vuelta (aplicado, BITACORA 2026-11-03 (9))

- B:H sin comentario.
- Fórmula de Excel en lo que se calcula con la misma fila: Descarga / Carga
  ajustada FD-prorrata, Valorizacion Descarga / Carga y Total C2 (cuadro 3),
  Total C1. Comentario en lo que mira otras filas: ranking cmg, Bloque
  ordenado, curvas monótonas, energía por remunerar, Valorización
  Descarga / Carga, Monto.
- AG:AH → "Energía descargada / cargada por remunerar [kWh]" (detalle en
  comentario).
- Monto a compensar en su propio cuadro (N° 4), en verde.
- Total del mes en la fila 1 sobre Total C1, Total C2 y Monto ("Total: …").

## Dudas resueltas

- Umbral de "SoC sobre el minimo": la columna sale de la hoja (no llega al
  monto), así que la duda queda sin efecto.
