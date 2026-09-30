# Trazabilidad e ingeniería inversa — `11_PAGOS_BESS_2607_Definitivo.xlsm`

> **Objetivo:** documentar el entregable final del proceso de pagos/balance asociado a BESS y SSCC, partiendo desde el `.xlsm` final y avanzando aguas arriba hasta encontrar cada origen primario. Este archivo está pensado como documento vivo: cada nueva fuente descubierta debe agregarse sin borrar la evidencia ya registrada.

## 1. Alcance de esta primera revisión

- Se inspeccionó la estructura interna del `.xlsm`, sus 20 hojas, fórmulas almacenadas, valores constantes, nombres definidos, tablas dinámicas y proyecto VBA.
- Se extrajo el código fuente VBA directamente desde `xl/vbaProject.bin`. El código completo se incluye al final de este documento.
- Se clasifican las celdas en **fórmula**, **valor constante escrito por macro (evidencia VBA)**, **valor constante probablemente manual** y **valor constante de origen aún no identificado**.
- **Limitación importante:** un `.xlsm` final no conserva historial de edición. Por sí solo no puede demostrar si un valor constante fue pegado a mano o escrito por una macro en una ejecución pasada. La clasificación “macro” solo se afirma cuando el VBA contiene instrucciones que escriben explícitamente ese rango/hoja. Si no existe esa evidencia, se marca como “manual/probable” o “origen pendiente”.

## 2. Mapa inicial de trazabilidad (de final hacia origen)

```text
11_PAGOS_BESS_2607_Definitivo.xlsm
├─ Resumen / PRORRATA_RETIROS / Compensacion total / CSV SAE_202607.csv  ← salidas
├─ Calculo E Costos
│  ├─ Medidores ← [carpeta del libro]/Medidas/Medidas_SAE.xlsx :: hoja Medidas A:I
│  ├─ CMg ← [carpeta del libro]/Cmg/cmg.xlsx :: hoja CMg A:I
│  ├─ FD ← [carpeta]/SSCC_Desempeño/SSCC_Desempeño_* :: CPF Horario + CSF Horario
│  ├─ Subastas ← 3_REMUNERACIÓN_SUBASTAS_E_ID_* :: hoja DB (leída por ADO)
│  ├─ Prorrata SSCC ← ORIGEN PENDIENTE
│  ├─ Diccionario ← mapeo constante; probable mantenimiento manual
│  └─ Resumen ← parámetros BESS + resultados de otras hojas
├─ Calculo RE545
│  ├─ Medidores
│  ├─ CMg
│  └─ fórmulas propias de metodología RE545
└─ Resumen Ofertas SSCC ← archivo más reciente cuyo nombre contiene OfertasSSCC (todas sus hojas A:I)
   └─ Medidores W:Y → macro de ofertas → Medidores AB:AE → macro de ventana/oferta completa
```

### Archivos externos ya identificados con evidencia directa en VBA

| Fuente externa | Ubicación relativa | Hoja / rango leído | Macro | Destino |
|---|---|---|---|---|
| `Medidas_SAE.xlsx` | `./Medidas/Medidas_SAE.xlsx` | `Medidas!A:I`, desde fila 2 | `Cargar_Medidas_SAE_En_Medidores` | `Medidores!A3:I...` |
| `cmg.xlsx` | `./Cmg/cmg.xlsx` | `CMg!A:I`, desde fila 2 | `Cargar_CMg_Desde_Archivo` | `CMg!A2:I...` |
| `SSCC_Desempeño_*` | `./SSCC_Desempeño/` | hojas `CPF Horario` y `CSF Horario` | `Cargar_SSCC_Desempeno_En_FD` | `FD` + fórmulas plantilla |
| `3_REMUNERACIÓN_SUBASTAS_E_ID_*` | misma carpeta del `.xlsm` | `DB!B3:Y1048576` mediante ADO/SQL | `Cargar_Remuneracion_Subastas_Rapido` | `Subastas` |
| `*OfertasSSCC*` | misma carpeta del `.xlsm` | todas las hojas, `A:I` | `Generar_Resumen_Ofertas_SSCC` | `Resumen Ofertas SSCC` y luego `Medidores!W:Y` |

## 3. Inventario de hojas y clasificación de la información

| # | Hoja | Dimensión observada | Fórmulas | Valores constantes | Rol | Escritura por macro identificada | Observación |
|---:|---|---:|---:|---:|---|---|---|
| 1 | `Instructivo` | `E19` | 0 | 8 | Entrada/parametrización | Sin macro escritora detectada | Celdas de año/mes, supuestos e instrucciones; probable edición manual. |
| 2 | `Control de Cambios` | `D6` | 0 | 10 | Documentación | Sin macro escritora detectada | Registro documental; valores constantes. |
| 3 | `Medidores` | `AE26786` | 187,797 | 269,953 | Entrada + cálculo auxiliar | D, H, I y fórmulas | A:I se cargan desde Medidas_SAE.xlsx; W:Y y AB:AE se escriben por macros; varias columnas intermedias son fórmulas. J queda como valor constante sin macro escritora identificada. |
| 4 | `CMg` | `I26785` | 0 | 241,065 | Entrada materializada | E_lee_cmg_origen | A:I se reemplaza desde Cmg/cmg.xlsx, hoja CMg. |
| 5 | `FD` | `AE92171` | 80,352 | 107,171 | Entrada + fórmulas plantilla | F_Leer_FD | Carga hojas CPF Horario / CSF Horario de SSCC_Desempeño_* y extiende fórmulas de plantilla. |
| 6 | `Prorrata SSCC` | `D3835` | 0 | 15,321 | Entrada materializada | No se detectó macro escritora | Tabla de configuración/hora-mes con prorratas CPF/CSF. Debe rastrearse su origen en la siguiente iteración. |
| 7 | `Subastas` | `W32080` | 16,774 | 98,659 | Entrada + fórmulas plantilla | G_Lee_Subastas | Lee DB de 3_REMUNERACIÓN_SUBASTAS_E_ID_* mediante ADO y repone fórmulas M:N; resumen U:W calculado con fórmulas. |
| 8 | `Calculo E Costos` | `BP26787` | 26,875 | 1,207,438 | Cálculo / salida | J_Calculo_Ecostos | A:G, I:J, K/L/T llegan desde Medidores según macro B; Q:R desde CMg por macro A; gran parte de L:AZ es escrita por macro J; existen fórmulas auxiliares y resúmenes a la derecha. |
| 9 | `Hoja1` | `C779` | 0 | 1,580 | Auxiliar / salida | No identificado / fórmulas / pivote | Salida de tabla dinámica / resumen; no debe tratarse como dato origen. |
| 10 | `Gráfico Energía` | `AH99` | 16 | 935 | Auxiliar / salida | No identificado / fórmulas / pivote | Hoja auxiliar de gráfico con datos y unas pocas fórmulas. |
| 11 | `Gráfico Valorizado` | `C124` | 0 | 368 | Auxiliar / salida | No identificado / fórmulas / pivote | Salida de tabla dinámica / resumen; no debe tratarse como dato origen. |
| 12 | `Calculo RE545` | `CE26787` | 1,100,445 | 456,291 | Cálculo / salida | No identificado / fórmulas / pivote | A:G, I:J, K/L/T llegan desde Medidores; Q:R desde CMg; cálculo principal de metodología RE545 se mantiene mayoritariamente en fórmulas. |
| 13 | `PRORRATA_RETIROS` | `P238812` | 483,652 | 716,518 | Cálculo / salida | No identificado / fórmulas / pivote | Gran tabla de cálculo con fórmulas y constantes; alimenta Resumen y exportación de pagos. Origen aguas arriba aún por detallar. |
| 14 | `Compensacion total` | `R380` | 3,411 | 431 | Cálculo / salida | No identificado / fórmulas / pivote | Consolidación calculada que alimenta Resumen. |
| 15 | `Diccionario` | `J11` | 0 | 41 | Auxiliar / salida | No identificado / fórmulas / pivote | Mapeos de nombres entre fuentes; constantes, sin macro escritora detectada. Probable mantenimiento manual. |
| 16 | `Resumen` | `Q102` | 384 | 277 | Cálculo / salida | No identificado / fórmulas / pivote | Contiene parámetros BESS, mapeos/empresas y resumen monetario; mezcla constantes y fórmulas. |
| 17 | `Aportes SSCC` | `G6` | 0 | 14 | Auxiliar / salida | No identificado / fórmulas / pivote | Tabla pequeña de aportes SSCC; constantes, origen no identificado aún. |
| 18 | `Verificacion_CSV` | `H15` | 0 | 33 | Auxiliar / salida | C_Exporta_pagos | Es limpiada y escrita por macro C durante la generación del CSV SAE_AAAAMM.csv. |
| 19 | `Resumen Ofertas SSCC` | `H198` | 0 | 1,584 | Auxiliar / salida | H_Leer_Ofertas | Es recreada por macro H desde archivos cuyo nombre contiene OfertasSSCC. |
| 20 | `Log Calculo E Costos` | `J36` | 0 | 360 | Auxiliar / salida | No identificado / fórmulas / pivote | Es escrita por macro J y documenta la lógica de cada columna calculada. |

## 4. Clasificación de “pegado a mano” vs macro vs fórmula

### 4.1 Valores con evidencia fuerte de haber sido escritos por macros

- `Medidores!A3:I...`: la macro **D_Lee_medidas_origen** borra el rango anterior y pega los datos de `Medidas_SAE.xlsx`.
- `CMg!A2:I...`: la macro **E_lee_cmg_origen** borra y reemplaza el bloque con `cmg.xlsx`. Por eso, aunque en el archivo final sean constantes, **no deben catalogarse como pegados manualmente**.
- `FD`: la macro **F_Leer_FD** trae datos desde `SSCC_Desempeño_*` y además conserva/replica fórmulas plantilla (`A:C`, `K:M`, `Q:S`, `AC:AE`).
- `Subastas`: la macro **G_Lee_Subastas** escribe datos leídos por ADO y vuelve a extender fórmulas `M:N`.
- `Resumen Ofertas SSCC`: la macro **H_Leer_Ofertas** recrea el resumen desde archivos `OfertasSSCC`.
- `Medidores!W:Y`: la macro **H_Leer_Ofertas** escribe nombre/día/oferta completa.
- `Medidores!AB:AE`: la macro **I_Ofertas_ventana** escribe central, ventana T, oferta agregada y bandera de completitud.
- `Calculo E Costos!L:O, R:U, W:Y, AB:AF, AG:AX, AZ`: la macro **J_Calculo_Ecostos** calcula estos bloques en memoria y los escribe como **valores**, no como fórmulas. Esto explica por qué muchos resultados del cálculo aparecen como constantes en el `.xlsm`.
- `Calculo E Costos!Q:R` y `Calculo RE545!Q:R`: la macro **A_Carga_Cmg_a_Destino** asigna CMg desde la hoja `CMg`.
- `Calculo E Costos` y `Calculo RE545` reciben además datos desde `Medidores` mediante **B_medidores_a_calculos**.
- `Verificacion_CSV`: la macro **C_Exporta_pagos** la prepara/limpia y escribe alertas y totales antes de guardar `SAE_AAAAMM.csv`.
- `Log Calculo E Costos`: la macro **J_Calculo_Ecostos** registra la lógica y la ejecución del cálculo.

### 4.2 Valores constantes que parecen manuales o parametrización

- `Instructivo!C3` (año = 2026), `Instructivo!C4` (mes = 7) y los supuestos del instructivo: **alta probabilidad de edición manual/parametrización**.
- `Resumen!B:J` contiene parámetros de BESS (Pmax, horas, capacidad, energía mínima, barra, umbral de ciclo, ciclos diarios, eficiencia). No se encontró una macro que los cargue: **probable tabla maestra mantenida manualmente**, hasta que se identifique otra fuente.
- `Diccionario`: equivalencias de nombres entre FD/Subastas/Ofertas. No se detectó macro escritora: **probable mantenimiento manual**.
- `Prorrata SSCC` y `Aportes SSCC`: son constantes en el archivo final y no se identificó macro de carga en este proyecto VBA. **No se debe afirmar todavía que son manuales**; su origen queda pendiente.
- `Medidores!J` (encabezado “Valor Obtenido de Scada” / carga %) contiene valores constantes y no fue encontrada una macro escritora directa entre los módulos del libro. **Origen pendiente**; es un candidato prioritario para rastrear.


## 4.3 Estado de entradas no cargadas por macros — validación con usuario (2026-09-09)

Esta sección registra lo confirmado durante la reconstrucción del proceso. La columna **Estado** distingue conocimiento confirmado por el usuario de hipótesis todavía pendientes.

| ID | Entrada | Estado actual | Origen / forma de mantenimiento |
|---:|---|---|---|
| 1 | `Medidores!J` — valor SCADA / carga % | **Pendiente** | El origen aún no está identificado; se revisará posteriormente. |
| 2 | `Prorrata SSCC` | **Pendiente** | El origen aún no está identificado; se revisará posteriormente. |
| 3 | `PRORRATA_RETIROS` — bloque de entradas materializadas | **Confirmado parcialmente** | El usuario dispone del archivo/entrada de origen. A la fecha, esa información se **pega manualmente** en el libro. Falta documentar el archivo exacto, rango y procedimiento cuando se analice dicha entrada. |
| 4 | `Resumen!B:J` — parámetros de modelado BESS | **Entendimiento operativo, por confirmar** | Corresponderían a parámetros de modelado de los BESS que se incorporan al balance a medida que entran nuevos BESS al sistema. No se ha confirmado todavía una fuente maestra externa. |
| 5 | `Diccionario` | **Confirmado operacionalmente** | Se construye/mantiene a medida que llegan nuevos BESS. Debe tratarse como dato maestro evolutivo del proceso. |
| 6 | `Aportes SSCC` | **Pendiente** | Origen y forma de mantenimiento todavía no identificados. |
| 7 | `Instructivo` — año, mes y parámetros del período | **Confirmado** | Se completa **manualmente** para cada ejecución/balance. |

> **Criterio documental:** cuando se obtenga evidencia adicional, no borrar estos estados históricos. Actualizar el estado y agregar la fuente concreta, ruta, hoja/rango y procedimiento de generación.

## 5. Fórmulas del libro

> Para rangos con miles o millones de fórmulas repetidas se registra la **fórmula modelo** y el rango al que se extiende, en vez de copiar una versión por cada fila. Esto conserva la lógica sin inflar el documento con cientos de miles de fórmulas equivalentes. En fórmulas compartidas de Excel, el XML guarda la fórmula maestra y referencias compartidas para las demás celdas.

### 5.1 `Medidores`

| Rango / celda | Fórmula modelo |
|---|---|
| `K3:K26786` | `L3 (familia compartida por bloques; K refleja ciclo/ventana según plantilla)` |
| `L4:L26786` | `=IF(G4<>G3,0,IF(C4=C3,0,IF(C4=$S$1,1,0))+L3)` |
| `N3:N26786` | `=B3&"&"&E3` |
| `O3:O26786` | `=1*(J3>6%)` |
| `R3:R26786` | `=VLOOKUP(B3&G3,V:Y,4,FALSE)` |
| `S3:S26786` | `=IF(L3=L2,S2,IF(R3=1,1,2))` |
| `T3:T26786` | `=1-_xlfn.XLOOKUP(1,(AB$2:AB$10000=G3)*(AC$2:AC$10000=L3),AE$2:AE$10000,"")` |
| `V3:V312` | `=X3&_xlfn.XLOOKUP(W3,Diccionario!F:F,Diccionario!E:E,_xlfn.XLOOKUP(W3,Diccionario!G:G,Diccionario!E:E,0,0),0)` |

### 5.2 `FD`

| Rango / celda | Fórmula modelo |
|---|---|
| `A12:A6707` | `=B12&F12` |
| `B12:B6707` | `=(DAY(D12)-1)*24+E12+1+IF(DAY(D12)>100,1,0)` |
| `C12:C6707` | `=DAY(D12)` |
| `K12:K6707` | `=J12` |
| `L12:L6707` | `=K12 (la macro F preserva/propaga plantilla)` |
| `M12:M6707` | `=(DAY(D12)-1)*24+E12+1+IF(DAY(D12)>100,1,0)` |
| `Q12:Q6707` | `=R12&V12` |
| `R12:R6707` | `=(DAY(D12)-1)*24+E12+1+IF(DAY(D12)>100,1,0)` |
| `S12:S6707` | `=DAY(T12)` |
| `AC12:AC6707` | `=AA12` |
| `AD12:AD6707` | `=AC12` |
| `AE12:AE6707` | `=(DAY(D12)-1)*24+E12+1+IF(DAY(D12)>100,1,0)` |

### 5.3 `Subastas`

| Rango / celda | Fórmula modelo |
|---|---|
| `B1` | `=IF(AND(C1="CSF",D1="SUBIDA"),"CSF(+)",IF(AND(C1="CSF",D1="BAJADA"),"CSF(-)",IF(AND(C1="CTF",D1="SUBIDA"),"CTF(+)",IF(AND(C1="CTF",D1="BAJADA"),"CTF(-)","REVISAR"))))` |
| `M3:M6997` | `=K3&H3&I3` |
| `N3:N6997` | `=IFERROR(_xlfn.XLOOKUP(1,('Calculo E Costos'!$D$2:$D$50000=J3)*('Calculo E Costos'!$G$2:$G$50000=K3),'Calculo E Costos'!$P$2:$P$50000,""),"")` |
| `U3:U700` | `=S3&"&"&T3` |
| `V3:V700` | `=COUNTIFS($N:$N,$T3,$D:$D,V$2,K:K,S3)` |
| `W3:W700` | `=COUNTIFS($N:$N,$T3,$D:$D,W$2,K:K,S3)` |

### 5.4 `Calculo E Costos`

- Fórmulas detectadas: **26,875**.
- Valores constantes detectados: **1,207,438**.
- Regiones de fórmulas detectadas:

| Región | Definiciones explícitas | Primera fórmula almacenada | Última fórmula explícita almacenada |
|---|---:|---|---|
| `H4:H26787` | 26,784 | `VLOOKUP(G4,Resumen!B:G,6,FALSE)` | `VLOOKUP(G26787,Resumen!B:G,6,FALSE)` |
| `U1` | 1 | `SUM(U4:U1048576)` | `SUM(U4:U1048576)` |
| `AX1` | 1 | `SUM(AX4:AX1048576)` | `SUM(AX4:AX1048576)` |
| `AY1` | 1 | `SUM(AZ4:AZ1048576)` | `SUM(AZ4:AZ1048576)` |
| `BF4:BF11` | 3 | `Subastas!T3` | `BF5+1` |
| `BG4:BG11` | 3 | `BE4&"&"&BF4` | `BE7&"&"&BF7` |
| `BH4:BH11` | 1 | `COUNTIFS($AE:$AE,">"&0,$X:$X,$BF4)` | `COUNTIFS($AE:$AE,">"&0,$X:$X,$BF4)` |
| `BI4:BI11` | 1 | `COUNTIFS($AF:$AF,"<"&0,$X:$X,$BF4)` | `COUNTIFS($AF:$AF,"<"&0,$X:$X,$BF4)` |
| `BJ4:BJ11` | 1 | `COUNTIFS($I:$I,">"&0,$P:$P,$BF4)` | `COUNTIFS($I:$I,">"&0,$P:$P,$BF4)` |
| `BK4:BK11` | 1 | `COUNTIFS($J:$J,"<"&0,$P:$P,$BF4)` | `COUNTIFS($J:$J,"<"&0,$P:$P,$BF4)` |
| `BL4:BL11` | 8 | `Subastas!W3*4` | `Subastas!W10*4` |
| `BM4:BM11` | 8 | `Subastas!V3*4` | `Subastas!V10*4` |
| `BN4:BN11` | 2 | `SUMIF($P$4:$P$2979,BF4,$S$4:$S$2979)` | `SUMIF($P$4:$P$2979,BF5,$S$4:$S$2979)` |
| `BO4:BO11` | 2 | `SUMIF($P$4:$P$2979,BF4,$T$4:$T$2979)` | `SUMIF($P$4:$P$2979,BF5,$T$4:$T$2979)` |
| `BP4:BP11` | 2 | `BO4+BN4` | `BO5+BN5` |

### 5.5 `Gráfico Energía`

- Fórmulas detectadas: **16**.
- Valores constantes detectados: **935**.
- Regiones de fórmulas detectadas:

| Región | Definiciones explícitas | Primera fórmula almacenada | Última fórmula explícita almacenada |
|---|---:|---|---|
| `S36:S43` | 3 | `AVERAGE(R36:R43)` | `S37` |
| `S84:S91` | 3 | `AVERAGE(R84:R91)` | `S85` |

### 5.6 `Calculo RE545`

- Fórmulas detectadas: **1,100,445**.
- Valores constantes detectados: **456,291**.
- Regiones de fórmulas detectadas:

| Región | Definiciones explícitas | Primera fórmula almacenada | Última fórmula explícita almacenada |
|---|---:|---|---|
| `H1` | 1 | `SUM(I4:I11539)` | `SUM(I4:I11539)` |
| `H4:H26787` | 26,784 | `VLOOKUP(G4,Resumen!B:G,6,FALSE)` | `VLOOKUP(G26787,Resumen!B:G,6,FALSE)` |
| `J1` | 1 | `SUM(J4:J11539)` | `SUM(J4:J11539)` |
| `L4:L26787` | 26,784 | `1*(OR(COUNTIFS(Subastas!$K:$K,G4,Subastas!$G:$G,A4,Subastas!$H:$H,B4,Subastas!$I:$I,C4,Subastas!$D:$D,"BAJADA"),COUNTIFS(Subastas!$K:$K,G4,Subastas!$G:$G,A4,Subastas!$H:$H,B4,Subastas!$I:$I,C4,Subastas!$D:$D,"SUBIDA")))` | `1*(OR(COUNTIFS(Subastas!$K:$K,G26787,Subastas!$G:$G,A26787,Subastas!$H:$H,B26787,Subastas!$I:$I,C26787,Subastas!$D:$D,"BAJADA"),COUNTIFS(Subastas!$K:$K,G26787,Subastas!$G:$G,A26787,Subastas!$H:$H,B26787,Subastas!$I:$I,C26787,Subastas!$D:$D,"SUBIDA")))` |
| `M4:M26787` | 26,784 | `1*(K4>Resumen!$H$8)` | `1*(K26787>Resumen!$H$8)` |
| `N4:N26787` | 421 | `SUMIFS($I:$I,$G:$G,$G4,$P:$P,$P4,L:L,1)-SUMIFS($I:$I,$G:$G,$G4,$P:$P,$P4,$F:$F,"<"&F4,L:L,1)` | `SUMIFS($I:$I,$G:$G,$G26763,$P:$P,$P26763,L:L,1)-SUMIFS($I:$I,$G:$G,$G26763,$P:$P,$P26763,$F:$F,"<"&F26763,L:L,1)` |
| `O4:O26787` | 422 | `-SUMIFS($J:$J,$G:$G,$G4,$P:$P,$P4,L:L,1)+SUMIFS($J:$J,$G:$G,$G4,$P:$P,$P4,$F:$F,"<"&F4,L:L,1)` | `-SUMIFS($J:$J,$G:$G,$G26763,$P:$P,$P26763,L:L,1)+SUMIFS($J:$J,$G:$G,$G26763,$P:$P,$P26763,$F:$F,"<"&F26763,L:L,1)` |
| `R1` | 1 | `SUM(CC4:CC20000)` | `SUM(CC4:CC20000)` |
| `S4:S26787` | 420 | `(COUNTIFS($T:$T, T4, $R:$R, ">" & R4,G:G,G4) + COUNTIFS($T:$T, T4, $R:$R, R4, $C:$C, ">" & C4,G:G,G4) )/4+ 1` | `(COUNTIFS($T:$T, T26757, $R:$R, ">" & R26757,G:G,G26757) + COUNTIFS($T:$T, T26757, $R:$R, R26757, $C:$C, ">" & C26757,G:G,G26757) )/4+ 1` |
| `U4:U26787` | 26,784 | `K4*VLOOKUP(G4,Resumen!$B$8:$J$26,4,0)*1000` | `K26787*VLOOKUP(G26787,Resumen!$B$8:$J$26,4,0)*1000` |
| `V4:V26787` | 26,784 | `-SUMIFS(J:J,T:T,T4,G:G,G4)*VLOOKUP(G4,Resumen!$B$8:$J$26,9,0)` | `-SUMIFS(J:J,T:T,T26787,G:G,G26787)*VLOOKUP(G26787,Resumen!$B$8:$J$26,9,0)` |
| `W4:W12000` | 11,997 | `SUMIFS(Subastas!$O:$O,Subastas!$K:$K,'Calculo RE545'!$G4,Subastas!$J:$J,'Calculo RE545'!$D4,Subastas!$B:$B,'Calculo RE545'!AC$3)` | `SUMIFS(Subastas!$O:$O,Subastas!$K:$K,'Calculo RE545'!$G12000,Subastas!$J:$J,'Calculo RE545'!$D12000,Subastas!$B:$B,'Calculo RE545'!AC$3)` |
| `AC12001:AC26787` | 14,787 | `SUMIFS(Subastas!$O:$O,Subastas!$K:$K,'Calculo RE545'!$G12001,Subastas!$J:$J,'Calculo RE545'!$D12001,Subastas!$B:$B,'Calculo RE545'!AC$3)` | `SUMIFS(Subastas!$O:$O,Subastas!$K:$K,'Calculo RE545'!$G26787,Subastas!$J:$J,'Calculo RE545'!$D26787,Subastas!$B:$B,'Calculo RE545'!AC$3)` |
| `AD4:AD26787` | 26,784 | `SUMIFS(Subastas!$O:$O,Subastas!$K:$K,'Calculo RE545'!$G4,Subastas!$J:$J,'Calculo RE545'!$D4,Subastas!$B:$B,'Calculo RE545'!AD$3)` | `SUMIFS(Subastas!$O:$O,Subastas!$K:$K,'Calculo RE545'!$G26787,Subastas!$J:$J,'Calculo RE545'!$D26787,Subastas!$B:$B,'Calculo RE545'!AD$3)` |
| `AE4:AE26787` | 26,784 | `SUMIFS(Subastas!$O:$O,Subastas!$K:$K,'Calculo RE545'!$G4,Subastas!$J:$J,'Calculo RE545'!$D4,Subastas!$B:$B,'Calculo RE545'!AE$3)` | `SUMIFS(Subastas!$O:$O,Subastas!$K:$K,'Calculo RE545'!$G26787,Subastas!$J:$J,'Calculo RE545'!$D26787,Subastas!$B:$B,'Calculo RE545'!AE$3)` |
| `AF4:AF26787` | 26,784 | `SUMIFS(Subastas!$O:$O,Subastas!$K:$K,'Calculo RE545'!$G4,Subastas!$J:$J,'Calculo RE545'!$D4,Subastas!$B:$B,'Calculo RE545'!AF$3)` | `SUMIFS(Subastas!$O:$O,Subastas!$K:$K,'Calculo RE545'!$G26787,Subastas!$J:$J,'Calculo RE545'!$D26787,Subastas!$B:$B,'Calculo RE545'!AF$3)` |
| `AG4:AG26787` | 26,784 | `SUMIFS(Subastas!$O:$O,Subastas!$K:$K,'Calculo RE545'!$G4,Subastas!$J:$J,'Calculo RE545'!$D4,Subastas!$B:$B,'Calculo RE545'!AG$3)` | `SUMIFS(Subastas!$O:$O,Subastas!$K:$K,'Calculo RE545'!$G26787,Subastas!$J:$J,'Calculo RE545'!$D26787,Subastas!$B:$B,'Calculo RE545'!AG$3)` |
| `AH4:AH26787` | 26,784 | `SUMIFS(Subastas!$O:$O,Subastas!$K:$K,'Calculo RE545'!$G4,Subastas!$J:$J,'Calculo RE545'!$D4,Subastas!$B:$B,'Calculo RE545'!AH$3)` | `SUMIFS(Subastas!$O:$O,Subastas!$K:$K,'Calculo RE545'!$G26787,Subastas!$J:$J,'Calculo RE545'!$D26787,Subastas!$B:$B,'Calculo RE545'!AH$3)` |
| `AI4:AI26787` | 26,784 | `SUMIFS(Subastas!$P:$P,Subastas!$K:$K,'Calculo RE545'!$G4,Subastas!$J:$J,'Calculo RE545'!$D4,Subastas!$B:$B,'Calculo RE545'!AI$3)` | `SUMIFS(Subastas!$P:$P,Subastas!$K:$K,'Calculo RE545'!$G26787,Subastas!$J:$J,'Calculo RE545'!$D26787,Subastas!$B:$B,'Calculo RE545'!AI$3)` |
| `AJ4:AJ26787` | 26,784 | `SUMIFS(Subastas!$P:$P,Subastas!$K:$K,'Calculo RE545'!$G4,Subastas!$J:$J,'Calculo RE545'!$D4,Subastas!$B:$B,'Calculo RE545'!AJ$3)` | `SUMIFS(Subastas!$P:$P,Subastas!$K:$K,'Calculo RE545'!$G26787,Subastas!$J:$J,'Calculo RE545'!$D26787,Subastas!$B:$B,'Calculo RE545'!AJ$3)` |
| `AK4:AK26787` | 26,784 | `SUMIFS(Subastas!$P:$P,Subastas!$K:$K,'Calculo RE545'!$G4,Subastas!$J:$J,'Calculo RE545'!$D4,Subastas!$B:$B,'Calculo RE545'!AK$3)` | `SUMIFS(Subastas!$P:$P,Subastas!$K:$K,'Calculo RE545'!$G26787,Subastas!$J:$J,'Calculo RE545'!$D26787,Subastas!$B:$B,'Calculo RE545'!AK$3)` |
| `AL4:AL26787` | 26,784 | `SUMIFS(Subastas!$P:$P,Subastas!$K:$K,'Calculo RE545'!$G4,Subastas!$J:$J,'Calculo RE545'!$D4,Subastas!$B:$B,'Calculo RE545'!AL$3)` | `SUMIFS(Subastas!$P:$P,Subastas!$K:$K,'Calculo RE545'!$G26787,Subastas!$J:$J,'Calculo RE545'!$D26787,Subastas!$B:$B,'Calculo RE545'!AL$3)` |
| `AM4:AM26787` | 26,784 | `SUMIFS(Subastas!$P:$P,Subastas!$K:$K,'Calculo RE545'!$G4,Subastas!$J:$J,'Calculo RE545'!$D4,Subastas!$B:$B,'Calculo RE545'!AM$3)` | `SUMIFS(Subastas!$P:$P,Subastas!$K:$K,'Calculo RE545'!$G26787,Subastas!$J:$J,'Calculo RE545'!$D26787,Subastas!$B:$B,'Calculo RE545'!AM$3)` |
| `AN4:AN26787` | 26,784 | `SUMIFS(Subastas!$P:$P,Subastas!$K:$K,'Calculo RE545'!$G4,Subastas!$J:$J,'Calculo RE545'!$D4,Subastas!$B:$B,'Calculo RE545'!AN$3)` | `SUMIFS(Subastas!$P:$P,Subastas!$K:$K,'Calculo RE545'!$G26787,Subastas!$J:$J,'Calculo RE545'!$D26787,Subastas!$B:$B,'Calculo RE545'!AN$3)` |
| `AO4:AO26787` | 26,784 | `SUMIFS(Subastas!$Q:$Q,Subastas!$K:$K,'Calculo RE545'!$G4,Subastas!$J:$J,'Calculo RE545'!$D4,Subastas!$B:$B,'Calculo RE545'!AO$3)` | `SUMIFS(Subastas!$Q:$Q,Subastas!$K:$K,'Calculo RE545'!$G26787,Subastas!$J:$J,'Calculo RE545'!$D26787,Subastas!$B:$B,'Calculo RE545'!AO$3)` |
| `AP4:AP26787` | 26,784 | `SUMIFS(Subastas!$Q:$Q,Subastas!$K:$K,'Calculo RE545'!$G4,Subastas!$J:$J,'Calculo RE545'!$D4,Subastas!$B:$B,'Calculo RE545'!AP$3)` | `SUMIFS(Subastas!$Q:$Q,Subastas!$K:$K,'Calculo RE545'!$G26787,Subastas!$J:$J,'Calculo RE545'!$D26787,Subastas!$B:$B,'Calculo RE545'!AP$3)` |
| `AQ4:AQ26787` | 26,784 | `SUMIFS(Subastas!$Q:$Q,Subastas!$K:$K,'Calculo RE545'!$G4,Subastas!$J:$J,'Calculo RE545'!$D4,Subastas!$B:$B,'Calculo RE545'!AQ$3)` | `SUMIFS(Subastas!$Q:$Q,Subastas!$K:$K,'Calculo RE545'!$G26787,Subastas!$J:$J,'Calculo RE545'!$D26787,Subastas!$B:$B,'Calculo RE545'!AQ$3)` |
| `AU4:AU26787` | 420 | `SUMPRODUCT($AC4:$AH4,$AI4:$AN4,$AO4:$AT4)/4*1000` | `SUMPRODUCT($AC26757:$AH26757,$AI26757:$AN26757,$AO26757:$AT26757)/4*1000` |
| `AZ4:AZ291` | 288 | `IFERROR(_xlfn.SINGLE(INDEX($U$4:$U$30000,_xlfn.AGGREGATE(15,6,(ROW($U$4:$U$30000)-ROW($U$4)+1)/(($G$4:$G$30000=$AW4)*($T$4:$T$30000=$AX4)),1))),"")` | `IFERROR(_xlfn.SINGLE(INDEX($U$4:$U$30000,_xlfn.AGGREGATE(15,6,(ROW($U$4:$U$30000)-ROW($U$4)+1)/(($G$4:$G$30000=$AW291)*($T$4:$T$30000=$AX291)),1))),"")` |
| `BA4:BA291` | 288 | `IFERROR(_xlfn.SINGLE(INDEX($V$4:$V$30000,_xlfn.AGGREGATE(15,6,(ROW($V$4:$V$30000)-ROW($V$4)+1)/(($G$4:$G$30000=$AW4)*($T$4:$T$30000=$AX4)),1))),"")` | `IFERROR(_xlfn.SINGLE(INDEX($V$4:$V$30000,_xlfn.AGGREGATE(15,6,(ROW($V$4:$V$30000)-ROW($V$4)+1)/(($G$4:$G$30000=$AW291)*($T$4:$T$30000=$AX291)),1))),"")` |
| `BB4:BB291` | 8 | `SUMIFS(AU:AU,T:T,AX4,G:G,AW4)` | `SUMIFS(AU:AU,T:T,AX282,G:G,AW282)` |
| `BC4:BC291` | 288 | `MIN(MAX(MIN(AZ4+BA4,VLOOKUP(AW4,Resumen!$B$8:$J$26,4,0)*1000),BA4),BB4)*AY4*(AX4<>31)*1` | `MIN(MAX(MIN(AZ291+BA291,VLOOKUP(AW291,Resumen!$B$8:$J$26,4,0)*1000),BA291),BB291)*AY291*(AX291<>31)*1` |
| `BD4:BD291` | 8 | `SUMIFS($BN:$BN,$BR:$BR,$AX4,$G:$G,$AW4)-BC4` | `SUMIFS($BN:$BN,$BR:$BR,$AX282,$G:$G,$AW282)-BC282` |
| `BE4:BE291` | 8 | `SUMIFS($BU:$BU,$BR:$BR,$AX4,$G:$G,$AW4)-BC4+BF4` | `SUMIFS($BU:$BU,$BR:$BR,$AX282,$G:$G,$AW282)-BC282+BF282` |
| `BF4:BF291` | 8 | `SUMIFS(BV:BV,BR:BR,AX4,G:G,AW4)` | `SUMIFS(BV:BV,BR:BR,AX282,G:G,AW282)` |
| `BG4:BG291` | 288 | `AND(AZ4+BA4-SUMIFS(BF:BF,AX:AX,AX4-1,AW:AW,AW4)>VLOOKUP(AW4,Resumen!$B$8:$J$26,4,0)*1000,BF3<>0)*1` | `AND(AZ291+BA291-SUMIFS(BF:BF,AX:AX,AX291-1,AW:AW,AW291)>VLOOKUP(AW291,Resumen!$B$8:$J$26,4,0)*1000,BF290<>0)*1` |
| `BH8:BH99` | 4 | `+IF(T8=T4,BI4+1,1)` | `+IF(T89=T85,BI85+1,1)` |
| `BI100:BI26787` | 417 | `+IF(T153=T149,BI149+1,1)` | `+IF(T26777=T26773,BI26773+1,1)` |
| `BJ8:BJ26787` | 421 | `BJ4` | `BJ26773` |
| `BK4:BK26787` | 420 | `SUMIFS(R:R,G:G,G4,S:S,BI4,T:T,T4,E:E,BJ4)` | `SUMIFS(R:R,G:G,G26757,S:S,BI26757,T:T,T26757,E:E,BJ26757)` |
| `BL2` | 1 | `SUM(BO4:BO1048576)` | `SUM(BO4:BO1048576)` |
| `BL4:BL26787` | 420 | `SUMIFS($Q:$Q,$S:$S,BI4,E:E,BJ4,G:G,G4,T:T,T4)` | `SUMIFS($Q:$Q,$S:$S,BI26757,E:E,BJ26757,G:G,G26757,T:T,T26757)` |
| `BN19921:BN19923` | 1 | `BN19921*BM19921` | `BN19921*BM19921` |
| `BO4:BO19920` | 313 | `BN4*BM4` | `BN19909*BM19909` |
| `BO19924:BO26787` | 107 | `BN19973*BM19973` | `BN26757*BM26757` |
| `BQ4:BQ26787` | 421 | `BS4+BN4` | `BS26777+BN26777` |
| `BR4:BR26787` | 421 | `T4` | `T26777` |
| `BT2` | 1 | `SUM(CE4:CE30000)` | `SUM(CE4:CE30000)` |
| `BT4:BT26787` | 26,784 | `SUMIFS($BU5:$BU$30000,$BR5:$BR$30000,BR4,$G5:$G$30000,G4)` | `SUMIFS($BU26788:$BU$30000,$BR26788:$BR$30000,BR26787,$G26788:$G$30000,G26787)` |
| `BU4:BU26787` | 427 | `IF(BS4=0,0,MAX(0,MIN(BQ4,SUMIFS($BC:$BC,AX:AX,BR4,AW:AW,G4)-BT4-SUMIFS(BV:BV,BR:BR,BR4,G:G,G4))))` | `IF(BS26777=0,0,MAX(0,MIN(BQ26777,SUMIFS($BC:$BC,AX:AX,BR26777,AW:AW,G26777)-BT26777-SUMIFS(BV:BV,BR:BR,BR26777,G:G,G26777))))` |
| `BV4:BV26787` | 26,784 | `IF(AND(C4=Medidores!$S$1-1,AU4<>0),AU4,0)` | `IF(AND(C26787=Medidores!$S$1-1,AU26787<>0),AU26787,0)` |
| `BW4:BW26787` | 421 | `I4+J4` | `I26777+J26777` |
| `BX4:BX26787` | 421 | `SUMIFS($BF:$BF,$AX:$AX,BR4,AW:AW,G4)` | `SUMIFS($BF:$BF,$AX:$AX,BR26777,AW:AW,G26777)` |
| `BY4:BY26787` | 26,784 | `IF(BW4<0,BX4,SUMIFS($BZ3:$BZ$4,$BR3:$BR$4,BR4,$G3:$G$4,G4))` | `IF(BW26787<0,BX26787,SUMIFS($BZ$4:$BZ26786,$BR$4:$BR26786,BR26787,$G$4:$G26786,G26787))` |
| `BZ4:BZ26787` | 421 | `IF(BW4<0,0,8)*SUMIFS(BG:BG,AW:AW,G4,AX:AX,BR4)` | `IF(BW26777<0,0,8)*SUMIFS(BG:BG,AW:AW,G26777,AX:AX,BR26777)` |
| `CA4:CA26787` | 428 | `BU4+BZ4` | `BU26777+BZ26777` |
| `CB4:CB26787` | 421 | `CA4*BM4` | `CA26777*BM26777` |
| `CE4:CE26787` | 420 | `IFERROR(MAX(SUMIFS(BO:BO,BR:BR,BR4,G:G,G4)-SUMIFS(CC:CC,BR:BR,BR4,G:G,G4),0)*AU4/SUMIFS(AU:AU,BR:BR,BR4,G:G,G4),0)` | `IFERROR(MAX(SUMIFS(BO:BO,BR:BR,BR26777,G:G,G26777)-SUMIFS(CC:CC,BR:BR,BR26777,G:G,G26777),0)*AU26777/SUMIFS(AU:AU,BR:BR,BR26777,G:G,G26777),0)` |

### 5.7 `PRORRATA_RETIROS`

- Fórmulas detectadas: **483,652**.
- Valores constantes detectados: **716,518**.
- Regiones de fórmulas detectadas:

| Región | Definiciones explícitas | Primera fórmula almacenada | Última fórmula explícita almacenada |
|---|---:|---|---|
| `C1` | 1 | `FD!$E$9` | `FD!$E$9` |
| `E7` | 1 | `SUM(F9:F1048576)` | `SUM(F9:F1048576)` |
| `E9:E238812` | 3,734 | `SUMIF(I:I,B9,K:K)` | `SUMIF(I:I,B238796,K:K)` |
| `F9:F238812` | 3,734 | `E9*D9` | `E238796*D238796` |
| `G7` | 1 | `SUM(K9:K1048576)` | `SUM(K9:K1048576)` |
| `H10:H2888` | 47 | `I9+1` | `I2826+1` |
| `I2889:I2984` | 2 | `I2890+1` | `I2954+1` |
| `J9:J2984` | 2,976 | `SUMIF('Calculo E Costos'!Y:Y,PRORRATA_RETIROS!I9,'Calculo E Costos'!AZ:AZ)+SUMIF('Calculo RE545'!F:F,PRORRATA_RETIROS!I9,'Calculo RE545'!CE:CE)` | `SUMIF('Calculo E Costos'!Y:Y,PRORRATA_RETIROS!I2984,'Calculo E Costos'!AZ:AZ)+SUMIF('Calculo RE545'!F:F,PRORRATA_RETIROS!I2984,'Calculo RE545'!CE:CE)` |
| `L7` | 1 | `SUM(O9:O1048576)` | `SUM(O9:O1048576)` |
| `O5` | 1 | `SUM(D9:D1048576)` | `SUM(D9:D1048576)` |
| `O9:O96` | 3 | `SUMIF(C:C,N9,F:F)` | `SUMIF(C:C,N74,F:F)` |

### 5.8 `Compensacion total`

- Fórmulas detectadas: **3,411**.
- Valores constantes detectados: **431**.
- Regiones de fórmulas detectadas:

| Región | Definiciones explícitas | Primera fórmula almacenada | Última fórmula explícita almacenada |
|---|---:|---|---|
| `B3:B380` | 27 | `Medidores!A3` | `B340` |
| `C4:C44` | 2 | `C3+1` | `C4+1` |
| `C46:C86` | 2 | `C45+1` | `C46+1` |
| `C88:C128` | 2 | `C87+1` | `C88+1` |
| `C130:C170` | 2 | `C129+1` | `C130+1` |
| `C172:C212` | 2 | `C171+1` | `C172+1` |
| `C214:C254` | 2 | `C213+1` | `C214+1` |
| `C256:C296` | 2 | `C255+1` | `C256+1` |
| `C298:C338` | 2 | `C297+1` | `C298+1` |
| `C340:C380` | 2 | `C339+1` | `C340+1` |
| `D3:D380` | 378 | `SUMIFS('Calculo E Costos'!AX:AX,'Calculo E Costos'!$G:$G,'Compensacion total'!$A3,'Calculo E Costos'!$X:$X,'Compensacion total'!$C3)` | `SUMIFS('Calculo E Costos'!AX:AX,'Calculo E Costos'!$G:$G,'Compensacion total'!$A380,'Calculo E Costos'!$X:$X,'Compensacion total'!$C380)` |
| `E3:E380` | 378 | `SUMIFS('Calculo E Costos'!U:U,'Calculo E Costos'!$G:$G,'Compensacion total'!$A3,'Calculo E Costos'!$X:$X,'Compensacion total'!$C3)` | `SUMIFS('Calculo E Costos'!U:U,'Calculo E Costos'!$G:$G,'Compensacion total'!$A380,'Calculo E Costos'!$X:$X,'Compensacion total'!$C380)` |
| `F3:F380` | 7 | `IF(D3>E3,D3-E3,0)` | `IF(D324>E324,D324-E324,0)` |
| `G4:G44` | 2 | `G3+1` | `G4+1` |
| `G46:G86` | 2 | `G45+1` | `G68+1` |
| `G88:G128` | 2 | `G87+1` | `G88+1` |
| `G130:G170` | 2 | `G129+1` | `G130+1` |
| `G172:G212` | 2 | `G171+1` | `G172+1` |
| `G214:G254` | 2 | `G213+1` | `G214+1` |
| `G256:G296` | 2 | `G255+1` | `G256+1` |
| `G298:G338` | 2 | `G297+1` | `G298+1` |
| `G340:G380` | 2 | `G339+1` | `G340+1` |
| `H3:H380` | 378 | `SUMIFS('Calculo RE545'!BO:BO,'Calculo RE545'!$BR:$BR,G3,'Calculo RE545'!G:G,'Compensacion total'!A3)` | `SUMIFS('Calculo RE545'!BO:BO,'Calculo RE545'!$BR:$BR,G380,'Calculo RE545'!G:G,'Compensacion total'!A380)` |
| `I3:I380` | 378 | `SUMIFS('Calculo RE545'!CC:CC,'Calculo RE545'!$BR:$BR,G3,'Calculo RE545'!G:G,'Compensacion total'!A3)` | `SUMIFS('Calculo RE545'!CC:CC,'Calculo RE545'!$BR:$BR,G380,'Calculo RE545'!G:G,'Compensacion total'!A380)` |
| `J3:J380` | 7 | `IF(H3>I3,H3-I3,0)` | `IF(H324>I324,H324-I324,0)` |
| `N3:N11` | 2 | `+SUMIF(A:A,M3,F:F)` | `+SUMIF(A:A,M4,F:F)` |
| `O3:O11` | 2 | `SUMIF(A:A,M3,J:J)` | `SUMIF(A:A,M4,J:J)` |
| `P3:P11` | 2 | `+O3+N3` | `+O4+N4` |

### 5.9 `Resumen`

- Fórmulas detectadas: **384**.
- Valores constantes detectados: **277**.
- Regiones de fórmulas detectadas:

| Región | Definiciones explícitas | Primera fórmula almacenada | Última fórmula explícita almacenada |
|---|---:|---|---|
| `B3` | 1 | `PRORRATA_RETIROS!B2` | `PRORRATA_RETIROS!B2` |
| `B17` | 0 | `` | `` |
| `D13:D14` | 2 | `E13/C13` | `E14/C14` |
| `E8` | 1 | `C8*D8` | `C8*D8` |
| `E15:E16` | 2 | `D15*C15` | `D16*C16` |
| `L7` | 1 | `SUM(O9:O1000)` | `SUM(O9:O1000)` |
| `L10:L16` | 1 | `L9+1` | `L9+1` |
| `L18:L102` | 1 | `L73+1` | `L73+1` |
| `O9:O102` | 94 | `SUMIF('Compensacion total'!L:L,Resumen!N9,'Compensacion total'!P:P)` | `SUMIF('Compensacion total'!L:L,Resumen!N102,'Compensacion total'!P:P)` |
| `P7` | 1 | `SUM(P9:P1000)` | `SUM(P9:P1000)` |
| `P9:P102` | 94 | `SUMIF(PRORRATA_RETIROS!N:N,Resumen!N9,PRORRATA_RETIROS!O:O)` | `SUMIF(PRORRATA_RETIROS!N:N,N102,PRORRATA_RETIROS!O:O)` |
| `Q7` | 1 | `SUM(Q9:Q1000)` | `SUM(Q9:Q1000)` |
| `Q9:Q102` | 6 | `O9-P9` | `O102-P102` |

## 6. Lógica documentada por la propia macro `J_Calculo_Ecostos`

La hoja `Log Calculo E Costos` es especialmente valiosa porque el propio proceso dejó una descripción de cada columna calculada. Se transcribe a continuación como evidencia de trazabilidad:

| Columna | Qué se calcula | Origen | Proceso / destino |
|---|---|---|---|
| `L` | Marca con 1 si la combinación G-A-B-C de la fila existe en Subastas como registro BAJADA o SUBIDA; si no existe, deja 0. | Calculo E Costos!G4:G26787, A4:C26787; Subastas!D1:K6997 (D=tipo; G,H,I,K forman la clave equivalente). | Crea la clave G\|A\|B\|C; busca K\|G\|H\|I solo en filas con D=BAJADA o SUBIDA. Destino: Calculo E Costos!L4:L26787. |
| `M` | Compara K de cada fila con el umbral único de Resumen!H8: deja 1 cuando K>H8 y 0 en caso contrario. | Calculo E Costos!K4:K26787; Resumen!H8. | Lectura del umbral una sola vez y comparación fila a fila en memoria. Destino: Calculo E Costos!M4:M26787. |
| `N` | Para cada grupo con igual G y P, suma I de las filas con L=1 cuyo F es mayor o igual al F de la fila evaluada. | Calculo E Costos!G4:G26787, P4:P26787, F4:F26787, I4:I26787 y L4:L26787 calculada previamente. | Agrupa por G-P, ordena F de mayor a menor y acumula I por bloques de igual F; los empates reciben el mismo acumulado. Destino: Calculo E Costos!N4:N26787. |
| `O` | Para cada grupo con igual G y P, calcula el negativo de la suma J de las filas con L=1 cuyo F es mayor o igual al F evaluado. | Calculo E Costos!G4:G26787, P4:P26787, F4:F26787, J4:J26787 y L4:L26787 calculada previamente. | Agrupa por G-P, ordena F de mayor a menor, acumula J por bloques de igual F y cambia el signo del acumulado. Destino: Calculo E Costos!O4:O26787. |
| `R` | Asigna una posición dentro de cada grupo G-P, ordenando primero Q descendente y luego F descendente; iguales Q y F comparten posición. | Calculo E Costos!G4:G26787, P4:P26787, Q4:Q26787 y F4:F26787. | Ordenamiento en memoria por Q? y F?; la posición corresponde al inicio del bloque empatado. Destino: Calculo E Costos!R4:R26787. |
| `S` | Multiplica el valor I de cada fila por su valor Q. | Calculo E Costos!I4:I26787 y Q4:Q26787. | Operación fila a fila: S=I*Q. Destino: Calculo E Costos!S4:S26787. |
| `T` | Multiplica el valor J de cada fila por su valor Q. | Calculo E Costos!J4:J26787 y Q4:Q26787. | Operación fila a fila: T=J*Q. Destino: Calculo E Costos!T4:T26787. |
| `U` | Calcula únicamente el costo de filas reconocidas en L: multiplica L por la suma de S y T. | Calculo E Costos!L4:L26787, S4:S26787 y T4:T26787; S y T provienen de I, J y Q. | Operación equivalente a U=L*(S+T)=L*((I*Q)+(J*Q)). Destino: Calculo E Costos!U4:U26787. |
| `W` | Genera un contador secuencial por bloques consecutivos de P: W4 toma C4; desde la fila 5 aumenta W anterior en 1 si P actual=P anterior y reinicia en 1 cuando cambia. | Calculo E Costos!C4 y P4:P26787; utiliza además el valor W calculado para la fila anterior. | Recorrido estrictamente secuencial para respetar la dependencia entre filas. Destino: Calculo E Costos!W4:W26787. |
| `X` | Copia exactamente P en X, sin transformar el valor. | Calculo E Costos!P4:P26787. | Asignación directa X=P. Destino: Calculo E Costos!X4:X26787. |
| `Y` | Dentro de cada grupo G-P, coloca primero los F de filas con L=1 e I distinto de 0, ordenados por Q descendente; después coloca los F restantes. | Calculo E Costos!G4:G26787, P4:P26787, L4:L26787, I4:I26787, Q4:Q26787 y F4:F26787. | Separa filas calificadas/no calificadas; ordena las calificadas por Q? y fila original?, y distribuye los F sobre las posiciones del grupo. Destino: Calculo E Costos!Y4:Y26787. |
| `AB` | Dentro de cada grupo G-P, coloca los Q de filas con L=1 e I distinto de 0, ordenados de mayor a menor; las posiciones restantes quedan vacías. | Calculo E Costos!G4:G26787, P4:P26787, L4:L26787, I4:I26787 y Q4:Q26787. | Usa el mismo orden aplicado para Y (Q? y fila original?). Destino: Calculo E Costos!AB4:AB26787. |
| `AC` | Dentro de cada grupo G-P, coloca primero los F de filas con L=1 y J distinto de 0, ordenados por Q ascendente; después coloca los F restantes. | Calculo E Costos!G4:G26787, P4:P26787, L4:L26787, J4:J26787, Q4:Q26787 y F4:F26787. | Separa filas calificadas/no calificadas; ordena las calificadas por Q? y fila original?, y distribuye los F sobre las posiciones del grupo. Destino: Calculo E Costos!AC4:AC26787. |
| `AD` | Dentro de cada grupo G-P, coloca los Q de filas con L=1 y J distinto de 0, ordenados de menor a mayor; las posiciones restantes quedan vacías. | Calculo E Costos!G4:G26787, P4:P26787, L4:L26787, J4:J26787 y Q4:Q26787. | Usa el mismo orden aplicado para AC (Q? y fila original?). Destino: Calculo E Costos!AD4:AD26787. |
| `AE` | Distribuye el máximo N de cada grupo G-P según el número de bloque W y el factor asociado a G en Resumen B:C. | Calculo E Costos!G4:G26787, P4:P26787, N4:N26787 y W4:W26787; Resumen!B8:C16 (B=G, C=factor). | Calcula bloques=4*MAX(N)/factor/1000; asigna factor/4*1000 a bloques completos y la proporción al bloque fraccionario. Destino: Calculo E Costos!AE4:AE26787. |
| `AF` | Distribuye con signo negativo el máximo O de cada grupo G-P según W y el factor asociado a G en Resumen B:C. | Calculo E Costos!G4:G26787, P4:P26787, O4:O26787 y W4:W26787; Resumen!B8:C16 (B=G, C=factor). | Aplica la misma asignación por bloques de AE y cambia el signo del resultado. Destino: Calculo E Costos!AF4:AF26787. |
| `AG` | Suma la columna C de Prorrata SSCC para todas las filas donde A coincide con G y B coincide con D de la fila destino. | Calculo E Costos!G4:G26787 y D4:D26787; Prorrata SSCC!A1:D3835 (A=G, B=D, C=valor a sumar). | Diccionario de sumas por clave G-D; cuando no hay coincidencia deja 0. Destino: Calculo E Costos!AG4:AG26787. |
| `AH` | Suma la columna D de Prorrata SSCC para todas las filas donde A coincide con G y B coincide con D de la fila destino. | Calculo E Costos!G4:G26787 y D4:D26787; Prorrata SSCC!A1:D3835 (A=G, B=D, D=valor a sumar). | Diccionario de sumas por clave G-D; cuando no hay coincidencia deja 0. Destino: Calculo E Costos!AH4:AH26787. |
| `AI` | Asigna el valor constante 0 en todas las filas procesadas. | No utiliza rangos de origen. | Asignación directa. Destino: Calculo E Costos!AI4:AI26787. |
| `AJ` | Repite en AJ la misma suma calculada para AG: columna C de Prorrata SSCC por coincidencia G-D. | Calculo E Costos!G4:G26787 y D4:D26787; Prorrata SSCC!A1:D3835. | Reutiliza en memoria el resultado AG de la misma fila. Destino: Calculo E Costos!AJ4:AJ26787. |
| `AK` | Repite en AK la misma suma calculada para AH: columna D de Prorrata SSCC por coincidencia G-D. | Calculo E Costos!G4:G26787 y D4:D26787; Prorrata SSCC!A1:D3835. | Reutiliza en memoria el resultado AH de la misma fila. Destino: Calculo E Costos!AK4:AK26787. |
| `AL` | Asigna el valor constante 0 en todas las filas procesadas. | No utiliza rangos de origen. | Asignación directa. Destino: Calculo E Costos!AL4:AL26787. |
| `AM` | Obtiene el código equivalente de G en Diccionario A:B; calcula el bloque de Y y busca en FD la clave bloque+código para devolver FD columna AD. | Calculo E Costos!G4:G26787 y Y4:Y26787; Diccionario!A1:B11; FD!Q1:AD6707 (Q=clave, AD=resultado). | Bloque=ENTERO((Y-1)/4)+1; clave=texto del bloque concatenado con Diccionario!B. Primera coincidencia; si falta devuelve #N/A. Destino: Calculo E Costos!AM4:AM26787. |
| `AN` | Obtiene el código equivalente de G en Diccionario A:B; calcula el bloque de Y y busca en FD la clave bloque+código para devolver FD columna L. | Calculo E Costos!G4:G26787 y Y4:Y26787; Diccionario!A1:B11; FD!A1:L6707 (A=clave, L=resultado). | Bloque=ENTERO((Y-1)/4)+1; clave=texto del bloque concatenado con Diccionario!B. Primera coincidencia; si falta devuelve #N/A. Destino: Calculo E Costos!AN4:AN26787. |
| `AO` | Asigna el valor constante 0 en todas las filas procesadas. | No utiliza rangos de origen. | Asignación directa. Destino: Calculo E Costos!AO4:AO26787. |
| `AP` | Obtiene el código equivalente de G en Diccionario A:B; calcula el bloque de Y y busca en FD la clave bloque+código para devolver FD columna AC. | Calculo E Costos!G4:G26787 y Y4:Y26787; Diccionario!A1:B11; FD!Q1:AD6707 (Q=clave, AC=resultado). | Bloque=ENTERO((Y-1)/4)+1; clave=texto del bloque concatenado con Diccionario!B. Primera coincidencia; si falta devuelve #N/A. Destino: Calculo E Costos!AP4:AP26787. |
| `AQ` | Obtiene el código equivalente de G en Diccionario A:B; calcula el bloque de AC y busca en FD la clave bloque+código para devolver FD columna K. | Calculo E Costos!G4:G26787 y AC4:AC26787; Diccionario!A1:B11; FD!A1:L6707 (A=clave, K=resultado). | Bloque=ENTERO((AC-1)/4)+1; clave=texto del bloque concatenado con Diccionario!B. Primera coincidencia; si falta devuelve #N/A. Destino: Calculo E Costos!AQ4:AQ26787. |
| `AR` | Asigna el valor constante 0 en todas las filas procesadas. | No utiliza rangos de origen. | Asignación directa. Destino: Calculo E Costos!AR4:AR26787. |
| `AS` | Calcula AE por un factor formado con AG, AH, AM y AN. Si AG+AH es mayor que 0 usa AM*AG+AN*AH; si no, usa factor 1. | Calculo E Costos!AG4:AH26787, AM4:AN26787 y AE4:AE26787. | Por fila: AS=AE*(AM*AG+AN*AH) cuando AG+AH>0; de lo contrario AS=AE. Los errores de AE, AM o AN se propagan. Destino: Calculo E Costos!AS4:AS26787. |
| `AT` | Calcula AF por un factor formado con AJ, AK, AP y AQ. Si AJ+AK es mayor que 0 usa AP*AJ+AQ*AK; si no, usa factor 1. | Calculo E Costos!AJ4:AK26787, AP4:AQ26787 y AF4:AF26787. | Por fila: AT=AF*(AP*AJ+AQ*AK) cuando AJ+AK>0; de lo contrario AT=AF. Los errores de AF, AP o AQ se propagan. Destino: Calculo E Costos!AT4:AT26787. |
| `AU` | Para cada grupo G-P obtiene el promedio de AB solo en filas con AE numérico y distinto de 0; lo multiplica por AE si la suma total de I para ese P es mayor que 10. | Calculo E Costos!G4:G26787, P4:P26787, AB4:AB26787, AE4:AE26787 e I4:I26787. | El promedio se calcula por G-P, pero la condición SUMA(I)>10 se calcula por P considerando todas las G. Si algún AE del grupo es error, AU queda 0. Destino: Calculo E Costos!AU4:AU26787. |
| `AV` | Para cada grupo G-P obtiene el promedio de AD solo en filas con AF numérico y distinto de 0; lo multiplica por AF si la suma total de J para ese P es menor que -10. | Calculo E Costos!G4:G26787, P4:P26787, AD4:AD26787, AF4:AF26787 y J4:J26787. | El promedio se calcula por G-P, pero la condición SUMA(J)<-10 se calcula por P considerando todas las G. Si algún AF del grupo es error, AV queda 0. Destino: Calculo E Costos!AV4:AV26787. |
| `AW` | Calcula una diferencia valorizada usando promedios de AB y AD limitados por umbrales de subida y bajada definidos en Subastas U:W. | Calculo E Costos!G4:G26787, P4:P26787, W4:W26787, AB4:AB26787, AD4:AD26787, AE4:AF26787, AS4:AT26787; Subastas!U1:W700 (U=clave G&P, V=umbral subida, W=umbral bajada). | Por G-P promedia AB donde W<=umbralBajada*4 y AD donde W<=umbralSubida*4; luego AW=(AE-AS)*promAB-(AF-AT)*promAD. Destino: Calculo E Costos!AW4:AW26787. |
| `AX` | Combina los tres componentes económicos calculados previamente. | Calculo E Costos!AU4:AU26787, AV4:AV26787 y AW4:AW26787. | Operación fila a fila: AX=AU+AV-AW. Destino: Calculo E Costos!AX4:AX26787. |
| `AZ` | Para cada grupo G-P calcula la diferencia entre la suma de AX y la suma de U, la divide por la cantidad de filas y no permite resultados negativos. | Calculo E Costos!G4:G26787, P4:P26787, AX4:AX26787 y U4:U26787. | Por grupo: AZ=MAX(0;(SUMA(AX)-SUMA(U))/CONTAR(filas del grupo)); el mismo valor se escribe en todas las filas del G-P. Destino: Calculo E Costos!AZ4:AZ26787. |

## 7. Lectura hoja por hoja — qué sabemos y qué falta rastrear

### `Instructivo`

Parametrización del período y supuestos. `año` es un nombre definido que apunta a `Instructivo!C3`. No depende de fórmulas. Primer nivel manual.

### `Control de Cambios`

Documentación de cambios metodológicos. No interviene como fuente numérica directa salvo evidencia futura.

### `Medidores`

Hoja puente principal. A:I viene de `Medidas_SAE.xlsx`; J es una entrada constante cuyo origen aún no aparece en el VBA; K,L,N,O,R,S,T,V son cálculos; W:Y provienen de OfertasSSCC; AB:AE provienen del resumen de ventana/oferta. Luego macro B distribuye sus datos hacia ambas hojas de cálculo.

### `CMg`

Copia materializada de `cmg.xlsx`. Sus valores no son “manuales” aunque carezcan de fórmulas.

### `FD`

Capa de desempeño de SSCC. Mezcla datos importados de CPF/CSF Horario con fórmulas plantilla que generan claves, hora-mes y factores derivados.

### `Prorrata SSCC`

Entrada utilizada por macro J para AG/AH/AJ/AK. No existe macro de carga en este libro: localizar el Excel/CSV/proceso que genera esta tabla es una prioridad.

### `Subastas`

Capa importada desde archivo 3_REMUNERACIÓN_SUBASTAS_E_ID_*. La macro lee `DB!B:Y`, filtra/transforma y conserva fórmulas M:N; U:W resume conteos por configuración/ciclo y dirección.

### `Calculo E Costos`

Cálculo principal de Estudio de Costos. Importante: macro J reemplaza gran parte de antiguas fórmulas por valores calculados en memoria. Para reproducirlo fuera de Excel, el VBA J es una especificación casi completa.

### `Calculo RE545`

Cálculo alternativo/metodología RE545. Mantiene un volumen muy alto de fórmulas en hoja; requiere una segunda pasada específica para documentar por bloque funcional todas sus columnas.

### `PRORRATA_RETIROS`

Capa de cálculo masiva; alimenta `Resumen` (columna P vía SUMIF) y la exportación CSV. Necesita rastreo aguas arriba dedicado.

### `Compensacion total`

Consolida compensaciones y alimenta `Resumen!O` mediante SUMIF.

### `Resumen`

Combina parámetros BESS y pagos. Es una salida legible para empresa/configuración y a la vez contiene parámetros de entrada usados por macro J.

### `Diccionario`

Mapa de equivalencias entre nombres de distintas fuentes. Crítico para reproducibilidad; debe versionarse como dato maestro.

### `Verificacion_CSV`

Salida de control creada por macro C. No es fuente de cálculo; sirve para validación de exportación.

### `Resumen Ofertas SSCC`

Salida intermedia recreable desde archivos OfertasSSCC; no debe mantenerse manualmente.

### `Log Calculo E Costos`

Evidencia de ejecución y documentación funcional de la macro J; conservarla en cada versión ayuda a auditar cambios.

## 8. Prioridad propuesta para continuar la ingeniería inversa

1. **Medidores!J / dato SCADA o SoC**: identificar archivo/proceso que lo genera, porque macro B lo usa para distribuir energía y el archivo final no revela su carga.
2. **Prorrata SSCC**: identificar de dónde salen `Configuración`, `Hora_mes`, `CPF`, `CSF`; macro J depende directamente de esta tabla.
3. **PRORRATA_RETIROS**: descomponer sus fórmulas por bloques y rastrear todas las tablas origen, porque termina participando en pagos del `Resumen` y CSV.
4. **Calculo RE545**: documentar por columna la metodología y dependencias, similar al log ya existente para `Calculo E Costos`.
5. **Resumen / parámetros BESS / Diccionario / Aportes SSCC**: verificar si son maestros manuales o copias de otra fuente oficial.
6. Al recibir cada archivo origen, repetir la misma metodología: separar fórmula / valor materializado / macro, documentar ruta, hoja, rango, filtros y clave de unión, y enlazarlo en este documento.

## 9. Convención para futuras actualizaciones de este documento

Para cada nueva fuente descubierta agregar un registro con: **ID de fuente**, archivo, versión/fecha, ubicación, hoja/tabla, rango, campos, filtro, clave de unión, transformación, macro/fórmula que la consume, destino, responsable/origen oficial, y nivel de certeza (`confirmado`, `inferido`, `pendiente`). Nunca reemplazar una inferencia silenciosamente: cambiar su estado y dejar la evidencia que permitió confirmarla.

## 10. Código VBA completo extraído del `.xlsm`

> Los siguientes módulos fueron extraídos del proyecto VBA embebido. Los módulos de hoja y `ThisWorkbook` solo contienen atributos y no presentan procedimientos ejecutables. Los módulos funcionales se incluyen completos para que el cálculo pueda reconstruirse aunque se trabaje en otro chat/IA.

### Módulo `A_Carga_Cmg_a_Destino`

```vb
Attribute VB_Name = "A_Carga_Cmg_a_Destino"
Option Explicit

Sub Asignar_CMg_a_Calculos_Turbo()

    Dim wb As Workbook
    Dim wsCMg As Worksheet
    Dim dictCMg As Object
    
    Dim lastCMg As Long
    Dim arrCMg As Variant
    Dim i As Long
    
    Dim clave As String
    Dim barra As String
    Dim cuartoHora As String
    
    Dim t0 As Double
    Dim prevCalc As XlCalculation
    Dim prevScreenUpdating As Boolean
    Dim prevEnableEvents As Boolean
    Dim prevDisplayAlerts As Boolean
    
    On Error GoTo ErrHandler
    
    t0 = Timer
    
    Set wb = ThisWorkbook
    Set wsCMg = wb.Worksheets("CMg")
    Set dictCMg = CreateObject("Scripting.Dictionary")
    
    '====================================================
    ' MODO TURBO
    '====================================================
    With Application
        prevCalc = .Calculation
        prevScreenUpdating = .ScreenUpdating
        prevEnableEvents = .EnableEvents
        prevDisplayAlerts = .DisplayAlerts
        
        .ScreenUpdating = False
        .EnableEvents = False
        .DisplayAlerts = False
        .Calculation = xlCalculationManual
        .StatusBar = "Cargando datos CMg..."
    End With
    
    '====================================================
    ' 1) Cargar hoja CMg a diccionario
    '
    ' Hoja CMg:
    ' D = BARRA
    ' F = valor a asignar en Q
    ' H = Cuarto de Hora
    ' I = valor promedio horario a asignar en R
    '====================================================
    
    lastCMg = wsCMg.Cells(wsCMg.Rows.Count, "D").End(xlUp).Row
    
    If lastCMg < 2 Then
        MsgBox "La hoja CMg no tiene datos desde la fila 2.", vbExclamation
        GoTo Salida
    End If
    
    ' Rango D:I
    ' En el array:
    ' D = columna 1
    ' F = columna 3
    ' H = columna 5
    ' I = columna 6
    arrCMg = wsCMg.Range("D2:I" & lastCMg).Value
    
    For i = 1 To UBound(arrCMg, 1)
        
        barra = Trim(CStr(arrCMg(i, 1)))       ' CMg!D
        cuartoHora = NormalizaCuarto(arrCMg(i, 5)) ' CMg!H
        
        If barra <> "" And cuartoHora <> "" Then
            
            clave = UCase$(barra) & "|" & cuartoHora
            
            If Not dictCMg.Exists(clave) Then
                dictCMg.Add clave, Array(arrCMg(i, 3), arrCMg(i, 6)) ' F, I
            End If
            
        End If
        
    Next i
    
    '====================================================
    ' 2) Completar Calculo E Costos
    '
    ' Destino:
    ' F = Cuarto de Hora
    ' H = BARRA
    '
    ' Match:
    ' CMg!D = Destino!H
    ' CMg!H = Destino!F
    '
    ' Asigna:
    ' Q = CMg!F
    '====================================================
    
    Application.StatusBar = "Completando hoja Calculo E Costos..."
    
    CompletarDestinoTurbo _
        wsDestino:=wb.Worksheets("Calculo E Costos"), _
        dictCMg:=dictCMg, _
        escribirR:=False
    
    '====================================================
    ' 3) Completar Calculo RE545
    '
    ' Asigna:
    ' Q = CMg!F
    ' R = CMg!I
    '====================================================
    
    Application.StatusBar = "Completando hoja Calculo RE545..."
    
    CompletarDestinoTurbo _
        wsDestino:=wb.Worksheets("Calculo RE545"), _
        dictCMg:=dictCMg, _
        escribirR:=True
    
Salida:

    With Application
        .Calculation = prevCalc
        .ScreenUpdating = prevScreenUpdating
        .EnableEvents = prevEnableEvents
        .DisplayAlerts = prevDisplayAlerts
        .StatusBar = False
    End With
    
    MsgBox "Asignación CMg terminada en modo turbo." & vbCrLf & _
           "Registros CMg cargados: " & dictCMg.Count & vbCrLf & _
           "Tiempo: " & Format(Timer - t0, "0.0") & " segundos", vbInformation
    
    Exit Sub

ErrHandler:

    With Application
        .Calculation = prevCalc
        .ScreenUpdating = prevScreenUpdating
        .EnableEvents = prevEnableEvents
        .DisplayAlerts = prevDisplayAlerts
        .StatusBar = False
    End With
    
    MsgBox "Error en Asignar_CMg_a_Calculos_Turbo:" & vbCrLf & _
           Err.Number & " - " & Err.Description, vbCritical

End Sub


Private Sub CompletarDestinoTurbo( _
    ByVal wsDestino As Worksheet, _
    ByVal dictCMg As Object, _
    ByVal escribirR As Boolean)

    Dim lastDestinoF As Long
    Dim lastDestinoH As Long
    Dim lastDestino As Long
    
    Dim arrDestino As Variant
    Dim arrQ As Variant
    Dim arrR As Variant
    
    Dim i As Long
    Dim n As Long
    
    Dim barra As String
    Dim cuartoHora As String
    Dim clave As String
    
    Dim contadorMatch As Long
    Dim contadorSinMatch As Long
    
    ' Última fila considerando F y H
    lastDestinoF = wsDestino.Cells(wsDestino.Rows.Count, "F").End(xlUp).Row
    lastDestinoH = wsDestino.Cells(wsDestino.Rows.Count, "H").End(xlUp).Row
    lastDestino = Application.WorksheetFunction.Max(lastDestinoF, lastDestinoH)
    
    If lastDestino < 4 Then Exit Sub
    
    ' Leer de F a H desde fila 4
    ' En el array:
    ' F = columna 1
    ' G = columna 2
    ' H = columna 3
    arrDestino = wsDestino.Range("F4:H" & lastDestino).Value
    
    n = UBound(arrDestino, 1)
    
    ReDim arrQ(1 To n, 1 To 1)
    
    If escribirR Then
        ReDim arrR(1 To n, 1 To 1)
    End If
    
    For i = 1 To n
        
        cuartoHora = NormalizaCuarto(arrDestino(i, 1)) ' Destino!F
        barra = Trim(CStr(arrDestino(i, 3)))           ' Destino!H
        
        If cuartoHora <> "" And barra <> "" Then
            
            clave = UCase$(barra) & "|" & cuartoHora
            
            If dictCMg.Exists(clave) Then
                
                arrQ(i, 1) = dictCMg(clave)(0) ' CMg!F
                
                If escribirR Then
                    arrR(i, 1) = dictCMg(clave)(1) ' CMg!I
                End If
                
                contadorMatch = contadorMatch + 1
                
            Else
                
                arrQ(i, 1) = vbNullString
                
                If escribirR Then
                    arrR(i, 1) = vbNullString
                End If
                
                contadorSinMatch = contadorSinMatch + 1
                
            End If
            
        Else
            
            arrQ(i, 1) = vbNullString
            
            If escribirR Then
                arrR(i, 1) = vbNullString
            End If
            
        End If
        
    Next i
    
    ' Pegar resultados de una sola vez
    wsDestino.Range("Q4:Q" & lastDestino).Value = arrQ
    
    If escribirR Then
        wsDestino.Range("R4:R" & lastDestino).Value = arrR
    End If

End Sub


Private Function NormalizaCuarto(ByVal valor As Variant) As String

    If IsError(valor) Then
        NormalizaCuarto = ""
    ElseIf IsNumeric(valor) Then
        NormalizaCuarto = CStr(CLng(valor))
    Else
        NormalizaCuarto = Trim(CStr(valor))
    End If

End Function


```

### Módulo `B_medidores_a_calculos`

```vb
Attribute VB_Name = "B_medidores_a_calculos"
Option Explicit

Sub Traspasar_Medidores_A_Calculos_Rapido()

    Const FILA_ORIGEN As Long = 3
    Const FILA_DESTINO As Long = 4
    
    Const TOLERANCIA As Double = 0.000001
    Const TOLERANCIA_TOTAL As Double = 0.001
    
    Const HOJA_ERRORES As String = "Errores Traspaso Energia"

    Dim wb As Workbook
    
    Dim wsM As Worksheet
    Dim wsD As Worksheet
    Dim wsEC As Worksheet
    Dim wsRE As Worksheet
    Dim wsErr As Worksheet
    
    Dim nombreHoja As Variant

    Dim lastM As Long
    Dim lastDest As Long
    Dim lastClear As Long
    
    Dim n As Long
    Dim i As Long
    Dim j As Long

    '=====================================================
    ' ARRAYS ORIGEN
    '=====================================================
    Dim arrAG As Variant
    Dim arrAG_Dest As Variant
    
    Dim arrI As Variant
    Dim arrJ As Variant
    Dim arrK As Variant
    Dim arrL As Variant
    Dim arrT As Variant

    '=====================================================
    ' ARRAYS ENERGÍA A TRASPASAR
    '=====================================================
    Dim arrIJ_RE545() As Variant
    Dim arrIJ_ECostos() As Variant

    '=====================================================
    ' ARRAYS PARA VALIDACIÓN
    '=====================================================
    Dim arrValEC As Variant
    Dim arrValRE As Variant
    
    Dim arrErrores() As Variant

    '=====================================================
    ' VARIABLES DE TRABAJO
    '=====================================================
    Dim v As Variant
    Dim vT As Variant
    
    Dim energiaI As Double
    Dim energiaJ As Double
    
    Dim energiaOrigen As Double
    
    Dim energiaEC As Double
    Dim energiaRE As Double
    
    Dim esperadoEC As Double
    Dim esperadoRE As Double
    
    Dim diferenciaEC As Double
    Dim diferenciaRE As Double
    
    Dim totalOrigen As Double
    Dim totalEC As Double
    Dim totalRE As Double
    Dim diferenciaTotal As Double
    
    Dim cantidadErrores As Long
    Dim filaError As Long
    
    Dim destinoEsperado As String
    Dim observacion As String

    '=====================================================
    ' CONFIGURACIÓN EXCEL
    '=====================================================
    Dim calcMode As XlCalculation
    Dim colorVerde As Long
    Dim configuracionModificada As Boolean

    On Error GoTo ErrHandler

    Set wb = ThisWorkbook
    
    Set wsM = wb.Worksheets("Medidores")
    Set wsEC = wb.Worksheets("Calculo E Costos")
    Set wsRE = wb.Worksheets("Calculo RE545")

    colorVerde = RGB(226, 239, 218)

    '=====================================================
    ' MÁXIMA VELOCIDAD
    '=====================================================
    With Application
        
        .ScreenUpdating = False
        .EnableEvents = False
        
        calcMode = .Calculation
        .Calculation = xlCalculationManual
        
        .DisplayAlerts = False
        
        .StatusBar = "Preparando traspaso de medidores..."
        
    End With

    configuracionModificada = True

    '=====================================================
    ' DETERMINAR ÚLTIMA FILA DE MEDIDORES
    '
    ' No se utiliza T porque podría tener fórmulas
    ' más abajo que los datos reales.
    '=====================================================
    lastM = UltimaFilaEnColumnas( _
        wsM, _
        Array( _
            "A", "B", "C", "D", "E", "F", "G", _
            "I", "J", "K", "L" _
        ) _
    )

    If lastM < FILA_ORIGEN Then
        
        MsgBox _
            "No hay datos para traspasar en la hoja Medidores.", _
            vbExclamation
        
        GoTo Salida
        
    End If

    n = lastM - FILA_ORIGEN + 1

    '=====================================================
    ' CARGAR TODO EL ORIGEN EN MEMORIA
    '=====================================================
    Application.StatusBar = _
        "Leyendo " & Format$(n, "#,##0") & " registros..."

    arrAG = wsM.Range( _
        "A" & FILA_ORIGEN & ":G" & lastM _
    ).Value2

    arrI = wsM.Range( _
        "I" & FILA_ORIGEN & ":I" & lastM _
    ).Value2

    arrJ = wsM.Range( _
        "J" & FILA_ORIGEN & ":J" & lastM _
    ).Value2

    arrK = wsM.Range( _
        "K" & FILA_ORIGEN & ":K" & lastM _
    ).Value2

    arrL = wsM.Range( _
        "L" & FILA_ORIGEN & ":L" & lastM _
    ).Value2

    arrT = wsM.Range( _
        "T" & FILA_ORIGEN & ":T" & lastM _
    ).Value2

    '=====================================================
    ' PREPARAR COLUMNAS A:G
    '
    ' Medidores D -> Destino E
    ' Medidores E -> Destino D
    '=====================================================
    ReDim arrAG_Dest(1 To n, 1 To 7)

    For i = 1 To n

        For j = 1 To 7
            arrAG_Dest(i, j) = arrAG(i, j)
        Next j

        arrAG_Dest(i, 4) = arrAG(i, 5)
        arrAG_Dest(i, 5) = arrAG(i, 4)

    Next i

    '=====================================================
    ' PREPARAR ENERGÍA
    '=====================================================
    ReDim arrIJ_RE545(1 To n, 1 To 2)
    ReDim arrIJ_ECostos(1 To n, 1 To 2)

    Application.StatusBar = "Distribuyendo energía en memoria..."

    For i = 1 To n

        v = arrI(i, 1)
        vT = arrT(i, 1)

        energiaI = 0#
        energiaJ = 0#

        '=================================================
        ' SEPARAR ENERGÍA SEGÚN SIGNO
        '
        ' Positivo -> I
        ' Negativo -> J
        '=================================================
        If Not IsError(v) Then
            
            If IsNumeric(v) And Len(Trim$(CStr(v))) > 0 Then

                If CDbl(v) > 0 Then
                    
                    energiaI = CDbl(v)
                    energiaJ = 0#

                ElseIf CDbl(v) < 0 Then
                    
                    energiaI = 0#
                    energiaJ = CDbl(v)

                Else
                    
                    energiaI = 0#
                    energiaJ = 0#

                End If

            End If
            
        End If

        '=================================================
        ' DISTRIBUIR SEGÚN T
        '
        ' T = 1     -> Calculo E Costos
        ' T <> 1    -> Calculo RE545
        ' T vacío   -> Calculo RE545
        '=================================================
        If Not IsError(vT) Then
            
            If IsNumeric(vT) And _
               Len(Trim$(CStr(vT))) > 0 Then

                If CDbl(vT) = 1 Then

                    arrIJ_ECostos(i, 1) = energiaI
                    arrIJ_ECostos(i, 2) = energiaJ

                    arrIJ_RE545(i, 1) = 0#
                    arrIJ_RE545(i, 2) = 0#

                Else

                    arrIJ_ECostos(i, 1) = 0#
                    arrIJ_ECostos(i, 2) = 0#

                    arrIJ_RE545(i, 1) = energiaI
                    arrIJ_RE545(i, 2) = energiaJ

                End If

            Else

                arrIJ_ECostos(i, 1) = 0#
                arrIJ_ECostos(i, 2) = 0#

                arrIJ_RE545(i, 1) = energiaI
                arrIJ_RE545(i, 2) = energiaJ

            End If
            
        Else

            arrIJ_ECostos(i, 1) = 0#
            arrIJ_ECostos(i, 2) = 0#

            arrIJ_RE545(i, 1) = energiaI
            arrIJ_RE545(i, 2) = energiaJ

        End If

    Next i

    '=====================================================
    ' TRASPASAR A AMBAS HOJAS
    '=====================================================
    For Each nombreHoja In Array( _
        "Calculo E Costos", _
        "Calculo RE545" _
    )

        Application.StatusBar = _
            "Traspasando a " & CStr(nombreHoja) & "..."

        Set wsD = wb.Worksheets(CStr(nombreHoja))

        '=================================================
        ' DETERMINAR RANGO A LIMPIAR
        '=================================================
        lastDest = UltimaFilaEnColumnas( _
            wsD, _
            Array( _
                "A", "B", "C", "D", "E", "F", "G", _
                "I", "J", "K", "P", "T" _
            ) _
        )

        lastClear = Application.Max( _
            lastDest, _
            FILA_DESTINO + n - 1 _
        )

        '=================================================
        ' LIMPIAR DATOS ANTERIORES
        '=================================================
        If lastClear >= FILA_DESTINO Then

            ' No tocar H
            wsD.Range( _
                "A" & FILA_DESTINO & ":G" & lastClear _
            ).ClearContents

            wsD.Range( _
                "I" & FILA_DESTINO & ":K" & lastClear _
            ).ClearContents

            wsD.Range( _
                "P" & FILA_DESTINO & ":P" & lastClear _
            ).ClearContents

            ' Quitar color anterior
            wsD.Range( _
                "A" & FILA_DESTINO & ":G" & lastClear _
            ).Interior.Pattern = xlNone

            wsD.Range( _
                "I" & FILA_DESTINO & ":K" & lastClear _
            ).Interior.Pattern = xlNone

            wsD.Range( _
                "P" & FILA_DESTINO & ":P" & lastClear _
            ).Interior.Pattern = xlNone

            ' T solo se usa en Calculo RE545
            If CStr(nombreHoja) = "Calculo RE545" Then

                wsD.Range( _
                    "T" & FILA_DESTINO & ":T" & lastClear _
                ).ClearContents

                wsD.Range( _
                    "T" & FILA_DESTINO & ":T" & lastClear _
                ).Interior.Pattern = xlNone

            End If

        End If

        '=================================================
        ' PEGAR A:G EN UN SOLO BLOQUE
        '=================================================
        wsD.Range( _
            "A" & FILA_DESTINO _
        ).Resize(n, 7).Value2 = arrAG_Dest

        '=================================================
        ' PEGAR ENERGÍA EN UN SOLO BLOQUE
        '=================================================
        If CStr(nombreHoja) = "Calculo E Costos" Then

            wsD.Range( _
                "I" & FILA_DESTINO _
            ).Resize(n, 2).Value2 = arrIJ_ECostos

        Else

            wsD.Range( _
                "I" & FILA_DESTINO _
            ).Resize(n, 2).Value2 = arrIJ_RE545

        End If

        '=================================================
        ' OTRAS COLUMNAS
        '=================================================

        ' Medidores J -> Destino K
        wsD.Range( _
            "K" & FILA_DESTINO _
        ).Resize(n, 1).Value2 = arrJ

        ' Medidores K -> Destino P
        wsD.Range( _
            "P" & FILA_DESTINO _
        ).Resize(n, 1).Value2 = arrK

        ' Medidores L -> RE545 T
        If CStr(nombreHoja) = "Calculo RE545" Then

            wsD.Range( _
                "T" & FILA_DESTINO _
            ).Resize(n, 1).Value2 = arrL

        End If

        '=================================================
        ' PINTAR COLUMNAS TRASPASADAS
        '=================================================
        wsD.Range( _
            "A" & FILA_DESTINO & _
            ":G" & FILA_DESTINO + n - 1 _
        ).Interior.Color = colorVerde

        wsD.Range( _
            "I" & FILA_DESTINO & _
            ":K" & FILA_DESTINO + n - 1 _
        ).Interior.Color = colorVerde

        wsD.Range( _
            "P" & FILA_DESTINO & _
            ":P" & FILA_DESTINO + n - 1 _
        ).Interior.Color = colorVerde

        If CStr(nombreHoja) = "Calculo RE545" Then

            wsD.Range( _
                "T" & FILA_DESTINO & _
                ":T" & FILA_DESTINO + n - 1 _
            ).Interior.Color = colorVerde

        End If

    Next nombreHoja

    '*************************************************************
    '*************************************************************
    '
    ' VALIDACIÓN DE ENERGÍA EFECTIVAMENTE ESCRITA
    '
    '*************************************************************
    '*************************************************************

    Application.StatusBar = _
        "Validando energía traspasada..."

    '=====================================================
    ' LEER LAS DOS HOJAS DESTINO EN BLOQUES
    '
    ' Esta es la parte importante para velocidad:
    ' NO se leen celdas una por una.
    '=====================================================
    arrValEC = wsEC.Range( _
        "I" & FILA_DESTINO & _
        ":J" & FILA_DESTINO + n - 1 _
    ).Value2

    arrValRE = wsRE.Range( _
        "I" & FILA_DESTINO & _
        ":J" & FILA_DESTINO + n - 1 _
    ).Value2

    '=====================================================
    ' PRIMERA PASADA:
    '
    ' - Contar errores
    ' - Calcular totales
    '
    ' No se crea todavía ningún array grande de errores.
    '=====================================================
    cantidadErrores = 0

    totalOrigen = 0#
    totalEC = 0#
    totalRE = 0#

    For i = 1 To n

        '=================================================
        ' ENERGÍA ORIGINAL
        '=================================================
        energiaOrigen = 0#

        v = arrI(i, 1)

        If Not IsError(v) Then
            
            If IsNumeric(v) And _
               Len(Trim$(CStr(v))) > 0 Then
                
                energiaOrigen = CDbl(v)
                
            End If
            
        End If

        '=================================================
        ' ENERGÍA REAL EN CALCULO E COSTOS
        '=================================================
        energiaEC = 0#

        If Not IsError(arrValEC(i, 1)) Then
            If IsNumeric(arrValEC(i, 1)) Then
                energiaEC = energiaEC + CDbl(arrValEC(i, 1))
            End If
        End If

        If Not IsError(arrValEC(i, 2)) Then
            If IsNumeric(arrValEC(i, 2)) Then
                energiaEC = energiaEC + CDbl(arrValEC(i, 2))
            End If
        End If

        '=================================================
        ' ENERGÍA REAL EN CALCULO RE545
        '=================================================
        energiaRE = 0#

        If Not IsError(arrValRE(i, 1)) Then
            If IsNumeric(arrValRE(i, 1)) Then
                energiaRE = energiaRE + CDbl(arrValRE(i, 1))
            End If
        End If

        If Not IsError(arrValRE(i, 2)) Then
            If IsNumeric(arrValRE(i, 2)) Then
                energiaRE = energiaRE + CDbl(arrValRE(i, 2))
            End If
        End If

        '=================================================
        ' DESTINO ESPERADO SEGÚN T
        '=================================================
        esperadoEC = 0#
        esperadoRE = 0#

        vT = arrT(i, 1)

        If Not IsError(vT) Then
            
            If IsNumeric(vT) And _
               Len(Trim$(CStr(vT))) > 0 Then

                If CDbl(vT) = 1 Then
                    
                    esperadoEC = energiaOrigen
                    esperadoRE = 0#
                    
                Else
                    
                    esperadoEC = 0#
                    esperadoRE = energiaOrigen
                    
                End If

            Else
                
                esperadoEC = 0#
                esperadoRE = energiaOrigen
                
            End If
            
        Else
            
            esperadoEC = 0#
            esperadoRE = energiaOrigen
            
        End If

        '=================================================
        ' DIFERENCIAS
        '=================================================
        diferenciaEC = energiaEC - esperadoEC
        diferenciaRE = energiaRE - esperadoRE

        If Abs(diferenciaEC) > TOLERANCIA Or _
           Abs(diferenciaRE) > TOLERANCIA Then
            
            cantidadErrores = cantidadErrores + 1
            
        End If

        '=================================================
        ' TOTALES
        '=================================================
        totalOrigen = totalOrigen + energiaOrigen
        totalEC = totalEC + energiaEC
        totalRE = totalRE + energiaRE

    Next i

    diferenciaTotal = _
        (totalEC + totalRE) - totalOrigen

    '=====================================================
    ' SI HAY ERRORES:
    ' HACER SEGUNDA PASADA EN MEMORIA
    '
    ' Solo ahora se reserva memoria para exactamente
    ' la cantidad de errores encontrada.
    '=====================================================
    If cantidadErrores > 0 Then

        Application.StatusBar = _
            "Preparando detalle de " & _
            Format$(cantidadErrores, "#,##0") & _
            " diferencias..."

        ReDim arrErrores( _
            1 To cantidadErrores, _
            1 To 11 _
        )

        filaError = 0

        For i = 1 To n

            '---------------------------------------------
            ' ORIGEN
            '---------------------------------------------
            energiaOrigen = 0#

            v = arrI(i, 1)

            If Not IsError(v) Then
                
                If IsNumeric(v) And _
                   Len(Trim$(CStr(v))) > 0 Then
                    
                    energiaOrigen = CDbl(v)
                    
                End If
                
            End If

            '---------------------------------------------
            ' E COSTOS REAL
            '---------------------------------------------
            energiaEC = 0#

            If Not IsError(arrValEC(i, 1)) Then
                If IsNumeric(arrValEC(i, 1)) Then
                    energiaEC = energiaEC + _
                                CDbl(arrValEC(i, 1))
                End If
            End If

            If Not IsError(arrValEC(i, 2)) Then
                If IsNumeric(arrValEC(i, 2)) Then
                    energiaEC = energiaEC + _
                                CDbl(arrValEC(i, 2))
                End If
            End If

            '---------------------------------------------
            ' RE545 REAL
            '---------------------------------------------
            energiaRE = 0#

            If Not IsError(arrValRE(i, 1)) Then
                If IsNumeric(arrValRE(i, 1)) Then
                    energiaRE = energiaRE + _
                                CDbl(arrValRE(i, 1))
                End If
            End If

            If Not IsError(arrValRE(i, 2)) Then
                If IsNumeric(arrValRE(i, 2)) Then
                    energiaRE = energiaRE + _
                                CDbl(arrValRE(i, 2))
                End If
            End If

            '---------------------------------------------
            ' ESPERADO
            '---------------------------------------------
            esperadoEC = 0#
            esperadoRE = 0#
            destinoEsperado = "Calculo RE545"

            vT = arrT(i, 1)

            If Not IsError(vT) Then
                
                If IsNumeric(vT) And _
                   Len(Trim$(CStr(vT))) > 0 Then

                    If CDbl(vT) = 1 Then
                        
                        esperadoEC = energiaOrigen
                        esperadoRE = 0#
                        destinoEsperado = "Calculo E Costos"
                        
                    Else
                        
                        esperadoEC = 0#
                        esperadoRE = energiaOrigen
                        destinoEsperado = "Calculo RE545"
                        
                    End If

                Else
                    
                    esperadoRE = energiaOrigen
                    destinoEsperado = "Calculo RE545"
                    
                End If
                
            Else
                
                esperadoRE = energiaOrigen
                destinoEsperado = "Calculo RE545"
                
            End If

            diferenciaEC = energiaEC - esperadoEC
            diferenciaRE = energiaRE - esperadoRE

            '---------------------------------------------
            ' REGISTRAR SOLO SI EXISTE ERROR
            '---------------------------------------------
            If Abs(diferenciaEC) > TOLERANCIA Or _
               Abs(diferenciaRE) > TOLERANCIA Then

                filaError = filaError + 1

                ' Detectar tipo de problema
                If Abs(energiaEC) > TOLERANCIA And _
                   Abs(energiaRE) > TOLERANCIA And _
                   Abs(energiaOrigen) > TOLERANCIA Then

                    observacion = _
                        "Energía presente en ambas hojas"

                ElseIf destinoEsperado = "Calculo E Costos" And _
                       Abs(energiaRE) > TOLERANCIA Then

                    observacion = _
                        "Energía indebida en Calculo RE545"

                ElseIf destinoEsperado = "Calculo RE545" And _
                       Abs(energiaEC) > TOLERANCIA Then

                    observacion = _
                        "Energía indebida en Calculo E Costos"

                ElseIf Abs(diferenciaEC) > TOLERANCIA And _
                       Abs(diferenciaRE) > TOLERANCIA Then

                    observacion = _
                        "Diferencia en ambas hojas"

                ElseIf Abs(diferenciaEC) > TOLERANCIA Then

                    observacion = _
                        "Diferencia en Calculo E Costos"

                Else

                    observacion = _
                        "Diferencia en Calculo RE545"

                End If

                arrErrores(filaError, 1) = _
                    FILA_ORIGEN + i - 1

                ' Central = Medidores columna G
                arrErrores(filaError, 2) = _
                    arrAG(i, 7)

                arrErrores(filaError, 3) = _
                    energiaOrigen

                arrErrores(filaError, 4) = _
                    arrT(i, 1)

                arrErrores(filaError, 5) = _
                    esperadoEC

                arrErrores(filaError, 6) = _
                    energiaEC

                arrErrores(filaError, 7) = _
                    esperadoRE

                arrErrores(filaError, 8) = _
                    energiaRE

                arrErrores(filaError, 9) = _
                    diferenciaEC

                arrErrores(filaError, 10) = _
                    diferenciaRE

                arrErrores(filaError, 11) = _
                    observacion

            End If

        Next i

        '=================================================
        ' OBTENER / CREAR HOJA DE ERRORES
        '=================================================
        Set wsErr = ObtenerHojaSiExiste( _
            wb, _
            HOJA_ERRORES _
        )

        If wsErr Is Nothing Then

            Set wsErr = wb.Worksheets.Add( _
                After:=wb.Worksheets(wb.Worksheets.Count) _
            )

            wsErr.Name = HOJA_ERRORES

        Else

            wsErr.UsedRange.ClearContents

        End If

        '=================================================
        ' ENCABEZADOS
        '=================================================
        wsErr.Range("A1:K1").Value = Array( _
            "Fila Medidores", _
            "Central", _
            "Energia Origen", _
            "T", _
            "Esperado E Costos", _
            "Real E Costos", _
            "Esperado RE545", _
            "Real RE545", _
            "Dif E Costos", _
            "Dif RE545", _
            "Observacion" _
        )

        wsErr.Range("A1:K1").Font.Bold = True

        '=================================================
        ' PEGAR TODOS LOS ERRORES DE UNA SOLA VEZ
        '=================================================
        wsErr.Range("A2").Resize( _
            cantidadErrores, _
            11 _
        ).Value2 = arrErrores

        wsErr.Range( _
            "C2:J" & cantidadErrores + 1 _
        ).NumberFormat = "#,##0.000000"

    Else

        '=================================================
        ' SI LA VALIDACIÓN ES CORRECTA:
        ' LIMPIAR EVENTUALES ERRORES DE EJECUCIÓN ANTERIOR
        '=================================================
        Set wsErr = ObtenerHojaSiExiste( _
            wb, _
            HOJA_ERRORES _
        )

        If Not wsErr Is Nothing Then

            wsErr.UsedRange.ClearContents

            wsErr.Range("A1").Value = _
                "Sin diferencias en la última validación."

        End If

    End If

    '=====================================================
    ' RESULTADO FINAL
    '=====================================================
    Application.StatusBar = False

    If cantidadErrores = 0 And _
       Abs(diferenciaTotal) <= TOLERANCIA_TOTAL Then

        MsgBox _
            "Traspaso y validación terminados correctamente." & _
            vbCrLf & vbCrLf & _
            "Registros procesados: " & _
            Format$(n, "#,##0") & vbCrLf & vbCrLf & _
            "Energía Medidores: " & _
            Format$(totalOrigen, "#,##0.000000") & vbCrLf & _
            "Energía Calculo E Costos: " & _
            Format$(totalEC, "#,##0.000000") & vbCrLf & _
            "Energía Calculo RE545: " & _
            Format$(totalRE, "#,##0.000000") & vbCrLf & _
            "Energía destinos total: " & _
            Format$(totalEC + totalRE, "#,##0.000000") & _
            vbCrLf & _
            "Diferencia total: " & _
            Format$(diferenciaTotal, "#,##0.000000") & _
            vbCrLf & vbCrLf & _
            "Registros con diferencias: 0" & vbCrLf & _
            "VALIDACIÓN CORRECTA", _
            vbInformation

    Else

        MsgBox _
            "Traspaso terminado, pero la validación detectó diferencias." & _
            vbCrLf & vbCrLf & _
            "Registros procesados: " & _
            Format$(n, "#,##0") & vbCrLf & _
            "Registros con diferencias: " & _
            Format$(cantidadErrores, "#,##0") & vbCrLf & vbCrLf & _
            "Energía Medidores: " & _
            Format$(totalOrigen, "#,##0.000000") & vbCrLf & _
            "Energía Calculo E Costos: " & _
            Format$(totalEC, "#,##0.000000") & vbCrLf & _
            "Energía Calculo RE545: " & _
            Format$(totalRE, "#,##0.000000") & vbCrLf & _
            "Energía destinos total: " & _
            Format$(totalEC + totalRE, "#,##0.000000") & _
            vbCrLf & _
            "Diferencia total: " & _
            Format$(diferenciaTotal, "#,##0.000000") & _
            vbCrLf & vbCrLf & _
            "Revisar hoja: " & HOJA_ERRORES, _
            vbExclamation

    End If

Salida:

    Application.StatusBar = False

    If configuracionModificada Then

        With Application
            .ScreenUpdating = True
            .EnableEvents = True
            .Calculation = calcMode
            .DisplayAlerts = True
        End With

    End If

    Exit Sub

ErrHandler:

    Application.StatusBar = False

    MsgBox _
        "Error en el traspaso o validación:" & _
        vbCrLf & vbCrLf & _
        Err.Description, _
        vbCritical

    Resume Salida

End Sub


'====================================================================
' OBTENER ÚLTIMA FILA REAL CON DATOS
'====================================================================
Private Function UltimaFilaEnColumnas( _
    ByVal ws As Worksheet, _
    ByVal cols As Variant _
) As Long

    Dim c As Variant
    Dim f As Range
    Dim maxFila As Long

    maxFila = 0

    For Each c In cols

        Set f = ws.Columns(CStr(c)).Find( _
            What:="*", _
            After:=ws.Columns(CStr(c)).Cells(1, 1), _
            LookIn:=xlFormulas, _
            LookAt:=xlPart, _
            SearchOrder:=xlByRows, _
            SearchDirection:=xlPrevious, _
            MatchCase:=False _
        )

        If Not f Is Nothing Then
            
            If f.Row > maxFila Then
                maxFila = f.Row
            End If
            
        End If

        Set f = Nothing

    Next c

    UltimaFilaEnColumnas = maxFila

End Function


'====================================================================
' BUSCAR HOJA SIN GENERAR ERROR
'====================================================================
Private Function ObtenerHojaSiExiste( _
    ByVal wb As Workbook, _
    ByVal nombreHoja As String _
) As Worksheet

    On Error Resume Next
    
    Set ObtenerHojaSiExiste = _
        wb.Worksheets(nombreHoja)
    
    On Error GoTo 0

End Function


```

### Módulo `C_Exporta_pagos`

```vb
Attribute VB_Name = "C_Exporta_pagos"
Option Explicit

Private Const FILA_INICIO As Long = 4

Private Const COL_MES As String = "A"
Private Const COL_DIA As String = "B"
Private Const COL_HORA As String = "C"
Private Const COL_MINUTO As String = "E"
Private Const COL_BLOQUE As String = "F"
Private Const COL_CONFIG As String = "G"

Private Const TIPO_PAGO_SALIDA As String = "BESS"

Private Const EPS_CERO As Double = 0.0000001
Private Const VALIDAR_CAMBIO_HORA As Boolean = True

Sub Generar_CSV_SAE()

    Dim anioTxt As String
    Dim anioExportacion As Long
    Dim calcAnterior As XlCalculation
    Dim optimizadoActivo As Boolean

    Dim dictDatos As Object
    Dim dictPrimeraFecha As Object
    Dim primeraFechaKey As String

    On Error GoTo ErrHandler

    If ThisWorkbook.Path = "" Then
        MsgBox "Primero guarda el archivo Excel antes de generar el CSV.", vbExclamation
        Exit Sub
    End If

    anioTxt = InputBox("Ingrese el año de exportación:", "Año CSV", Year(Date))

    If Trim(anioTxt) = "" Then Exit Sub

    If Not IsNumeric(anioTxt) Then
        MsgBox "El año ingresado no es válido.", vbCritical
        Exit Sub
    End If

    anioExportacion = CLng(anioTxt)

    If anioExportacion < 2000 Or anioExportacion > 2100 Then
        MsgBox "El año ingresado parece fuera de rango.", vbCritical
        Exit Sub
    End If

    calcAnterior = Application.Calculation
    optimizadoActivo = True

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "Preparando exportación CSV..."

    PrepararHojaVerificacion

    Set dictDatos = CreateObject("Scripting.Dictionary")
    Set dictPrimeraFecha = CreateObject("Scripting.Dictionary")

    AcumularHojaCSV _
        nombreHoja:="Calculo E Costos", _
        colPago:="AZ", _
        anioExportacion:=anioExportacion, _
        dictDatos:=dictDatos, _
        dictPrimeraFecha:=dictPrimeraFecha

    AcumularHojaCSV _
        nombreHoja:="Calculo RE545", _
        colPago:="CE", _
        anioExportacion:=anioExportacion, _
        dictDatos:=dictDatos, _
        dictPrimeraFecha:=dictPrimeraFecha

    AjustarHojaVerificacion

    If dictDatos.Count = 0 Then
        MsgBox "No hay registros con pago distinto de cero para exportar.", vbInformation
        GoTo Salir
    End If

    primeraFechaKey = ObtenerPrimeraFechaKey(dictPrimeraFecha)

    GuardarCSVConsolidado _
        dictDatos:=dictDatos, _
        primeraFechaKey:=primeraFechaKey

    RegistrarTotalesPago dictDatos

Salir:

    If optimizadoActivo Then
        Application.StatusBar = False
        Application.ScreenUpdating = True
        Application.EnableEvents = True
        Application.DisplayAlerts = True
        Application.Calculation = calcAnterior
    End If

    If Err.Number = 0 Then
        MsgBox "Proceso terminado correctamente." & vbCrLf & _
               "Archivo generado en:" & vbCrLf & ThisWorkbook.Path & vbCrLf & vbCrLf & _
               "Revisa la hoja 'Verificacion_CSV'.", vbInformation
    End If

    Exit Sub

ErrHandler:
    MsgBox "Error: " & Err.Description, vbCritical
    Resume Salir

End Sub

Private Sub AcumularHojaCSV(ByVal nombreHoja As String, _
                            ByVal colPago As String, _
                            ByVal anioExportacion As Long, _
                            ByVal dictDatos As Object, _
                            ByVal dictPrimeraFecha As Object)

    Dim ws As Worksheet
    Dim ultimaFila As Long
    Dim n As Long
    Dim i As Long
    Dim filaExcel As Long

    Dim arrMes As Variant
    Dim arrDia As Variant
    Dim arrHora As Variant
    Dim arrMinuto As Variant
    Dim arrBloque As Variant
    Dim arrConfig As Variant
    Dim arrPago As Variant

    Dim mes As Long
    Dim dia As Long
    Dim horaExcel As Long
    Dim minuto As Long
    Dim bloque As Long

    Dim fechaDia As Date
    Dim fechaHora As Date
    Dim fechaCSV As String
    Dim fechaKey As String
    Dim configuracion As String
    Dim pagoNum As Double

    Dim key As String
    Dim data As Variant

    Dim dictDia As Object
    Dim dictPrevFecha As Object
    Dim keyDia As String
    Dim keyConfig As String
    Dim fechaAnterior As Date
    Dim difMin As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nombreHoja)
    On Error GoTo 0

    If ws Is Nothing Then
        MsgBox "No existe la hoja: " & nombreHoja, vbCritical
        Exit Sub
    End If

    ultimaFila = ws.Cells(ws.Rows.Count, ws.Columns(COL_MES).Column).End(xlUp).Row

    If ultimaFila < FILA_INICIO Then Exit Sub

    n = ultimaFila - FILA_INICIO + 1

    Application.StatusBar = "Leyendo datos de " & nombreHoja & "..."

    arrMes = ws.Range(COL_MES & FILA_INICIO & ":" & COL_MES & ultimaFila).Value2
    arrDia = ws.Range(COL_DIA & FILA_INICIO & ":" & COL_DIA & ultimaFila).Value2
    arrHora = ws.Range(COL_HORA & FILA_INICIO & ":" & COL_HORA & ultimaFila).Value2
    arrMinuto = ws.Range(COL_MINUTO & FILA_INICIO & ":" & COL_MINUTO & ultimaFila).Value2
    arrBloque = ws.Range(COL_BLOQUE & FILA_INICIO & ":" & COL_BLOQUE & ultimaFila).Value2
    arrConfig = ws.Range(COL_CONFIG & FILA_INICIO & ":" & COL_CONFIG & ultimaFila).Value2
    arrPago = ws.Range(colPago & FILA_INICIO & ":" & colPago & ultimaFila).Value2

    Set dictDia = CreateObject("Scripting.Dictionary")
    Set dictPrevFecha = CreateObject("Scripting.Dictionary")

    Application.StatusBar = "Consolidando datos de " & nombreHoja & "..."

    For i = 1 To n

        If Trim(CStr(arrMes(i, 1))) <> "" Then

            filaExcel = FILA_INICIO + i - 1

            mes = CLng(arrMes(i, 1))
            dia = CLng(arrDia(i, 1))
            horaExcel = CLng(arrHora(i, 1))
            minuto = CLng(arrMinuto(i, 1))
            bloque = CLng(arrBloque(i, 1))
            configuracion = Trim(CStr(arrConfig(i, 1)))

            fechaDia = DateSerial(anioExportacion, mes, dia)

            ' En la hoja la hora viene como 1, 2, 3...
            ' En el CSV queda 00, 01, 02...
            fechaHora = fechaDia + TimeSerial(horaExcel - 1, minuto, 0)

            fechaCSV = Format(fechaHora, "dd-mm-yyyy hh:nn")
            fechaKey = Format(fechaHora, "yyyymmddhhnn")

            pagoNum = ValorNumerico(arrPago(i, 1))

            ' Elimina pagos cero antes de consolidar
            If Abs(pagoNum) > EPS_CERO Then

                ' Llave de consolidación:
                ' Bloque_15min + CONFIGURACION + SAE
                key = Format(bloque, "000000") & "|" & configuracion & "|" & TIPO_PAGO_SALIDA

                If Not dictDatos.Exists(key) Then

                    ' data:
                    ' 0 Fecha_Hora
                    ' 1 CONFIGURACION
                    ' 2 Central
                    ' 3 Pago acumulado
                    ' 4 Tipo_pago
                    ' 5 Bloque_15min
                    dictDatos.Add key, Array(fechaCSV, configuracion, "", pagoNum, TIPO_PAGO_SALIDA, bloque)

                Else

                    data = dictDatos(key)
                    data(3) = CDbl(data(3)) + pagoNum

                    If CStr(data(0)) <> fechaCSV Then
                        RegistrarVerificacion nombreHoja, configuracion, filaExcel, _
                            "Fecha distinta para mismo bloque", "", _
                            bloque, data(0), fechaCSV
                    End If

                    dictDatos(key) = data

                End If

                If Not dictPrimeraFecha.Exists(fechaKey) Then
                    dictPrimeraFecha.Add fechaKey, fechaKey
                End If

            End If

            If VALIDAR_CAMBIO_HORA Then

                keyConfig = nombreHoja & "|" & configuracion
                keyDia = nombreHoja & "|" & configuracion & "|" & Format(fechaDia, "yyyy-mm-dd")

                If Not dictDia.Exists(keyDia) Then
                    dictDia.Add keyDia, 1
                Else
                    dictDia(keyDia) = dictDia(keyDia) + 1
                End If

                If horaExcel > 24 Then
                    RegistrarVerificacion nombreHoja, configuracion, filaExcel, _
                        "Hora mayor a 24", "", _
                        horaExcel, "", ""
                End If

                If dictPrevFecha.Exists(keyConfig) Then

                    fechaAnterior = dictPrevFecha(keyConfig)
                    difMin = DateDiff("n", fechaAnterior, fechaHora)

                    Select Case difMin

                        Case 15
                            ' OK

                        Case 75
                            RegistrarVerificacion nombreHoja, configuracion, filaExcel, _
                                "Cambio hora +1 detectado", "", _
                                difMin, _
                                Format(fechaAnterior, "dd-mm-yyyy hh:nn"), _
                                Format(fechaHora, "dd-mm-yyyy hh:nn")

                        Case -45
                            RegistrarVerificacion nombreHoja, configuracion, filaExcel, _
                                "Cambio hora -1 detectado", "", _
                                difMin, _
                                Format(fechaAnterior, "dd-mm-yyyy hh:nn"), _
                                Format(fechaHora, "dd-mm-yyyy hh:nn")

                        Case 0
                            RegistrarVerificacion nombreHoja, configuracion, filaExcel, _
                                "Fecha/hora duplicada", "", _
                                difMin, _
                                Format(fechaAnterior, "dd-mm-yyyy hh:nn"), _
                                Format(fechaHora, "dd-mm-yyyy hh:nn")

                        Case Else
                            RegistrarVerificacion nombreHoja, configuracion, filaExcel, _
                                "Salto inesperado", "", _
                                difMin, _
                                Format(fechaAnterior, "dd-mm-yyyy hh:nn"), _
                                Format(fechaHora, "dd-mm-yyyy hh:nn")

                    End Select

                End If

                dictPrevFecha(keyConfig) = fechaHora

            End If

        End If

    Next i

    If VALIDAR_CAMBIO_HORA Then
        RegistrarResumenDias dictDia
    End If

End Sub

Private Sub GuardarCSVConsolidado(ByVal dictDatos As Object, _
                                  ByVal primeraFechaKey As String)

    Dim arrKeys() As String
    Dim idx As Long
    Dim key As Variant
    Dim data As Variant

    Dim lineas() As String
    Dim contador As Long
    Dim contenido As String

    Dim anioMes As String
    Dim rutaSalida As String
    Dim i As Long

    If dictDatos.Count = 0 Then Exit Sub

    ReDim arrKeys(0 To dictDatos.Count - 1)

    idx = 0
    For Each key In dictDatos.Keys
        arrKeys(idx) = CStr(key)
        idx = idx + 1
    Next key

    If UBound(arrKeys) > LBound(arrKeys) Then
        QuickSortStrings arrKeys, LBound(arrKeys), UBound(arrKeys)
    End If

    ReDim lineas(0 To dictDatos.Count)

    lineas(0) = "Fecha_Hora,CONFIGURACION,Central,Pago,Tipo_pago,Bloque_15min"
    contador = 0

    For i = LBound(arrKeys) To UBound(arrKeys)

        data = dictDatos(arrKeys(i))

        If Abs(CDbl(data(3))) > EPS_CERO Then

            contador = contador + 1

            lineas(contador) = _
                CSVValue(CStr(data(0))) & "," & _
                CSVValue(CStr(data(1))) & "," & _
                CSVValue(CStr(data(2))) & "," & _
                FormatearPagoNumeroCSV(CDbl(data(3))) & "," & _
                CSVValue(CStr(data(4))) & "," & _
                CSVValue(CStr(data(5)))

        End If

    Next i

    If contador = 0 Then
        MsgBox "Luego de consolidar, todos los pagos quedaron en cero. No se generó CSV.", vbInformation
        Exit Sub
    End If

    ReDim Preserve lineas(0 To contador)

    contenido = Join(lineas, vbCrLf) & vbCrLf

    anioMes = Right(Left(primeraFechaKey, 4), 2) & Mid(primeraFechaKey, 5, 2)

    rutaSalida = ThisWorkbook.Path & "\SAE_" & anioMes & ".csv"

    Application.StatusBar = "Guardando archivo SAE_" & anioMes & ".csv..."

    GuardarUTF8 rutaSalida, contenido

End Sub

Private Sub RegistrarTotalesPago(ByVal dictDatos As Object)

    Dim wsLog As Worksheet
    Dim dictTotales As Object
    Dim key As Variant
    Dim data As Variant
    Dim config As String
    Dim pago As Double
    Dim totalGeneral As Double

    Dim arrKeys() As String
    Dim idx As Long
    Dim i As Long
    Dim filaInicio As Long
    Dim fila As Long

    Set wsLog = ThisWorkbook.Worksheets("Verificacion_CSV")
    Set dictTotales = CreateObject("Scripting.Dictionary")

    For Each key In dictDatos.Keys

        data = dictDatos(key)
        pago = CDbl(data(3))

        If Abs(pago) > EPS_CERO Then

            config = CStr(data(1))

            If Not dictTotales.Exists(config) Then
                dictTotales.Add config, pago
            Else
                dictTotales(config) = CDbl(dictTotales(config)) + pago
            End If

            totalGeneral = totalGeneral + pago

        End If

    Next key

    If dictTotales.Count = 0 Then Exit Sub

    filaInicio = wsLog.Cells(wsLog.Rows.Count, "A").End(xlUp).Row + 2

    With wsLog
        .Cells(filaInicio, "A").Value = "RESUMEN PAGOS"
        .Cells(filaInicio, "A").Font.Bold = True

        .Cells(filaInicio + 1, "A").Value = "CONFIGURACION"
        .Cells(filaInicio + 1, "B").Value = "Total Pago"
        .Range(.Cells(filaInicio + 1, "A"), .Cells(filaInicio + 1, "B")).Font.Bold = True
    End With

    ReDim arrKeys(0 To dictTotales.Count - 1)

    idx = 0
    For Each key In dictTotales.Keys
        arrKeys(idx) = CStr(key)
        idx = idx + 1
    Next key

    If UBound(arrKeys) > LBound(arrKeys) Then
        QuickSortStrings arrKeys, LBound(arrKeys), UBound(arrKeys)
    End If

    fila = filaInicio + 2

    For i = LBound(arrKeys) To UBound(arrKeys)

        wsLog.Cells(fila, "A").Value = arrKeys(i)
        wsLog.Cells(fila, "B").Value = CDbl(dictTotales(arrKeys(i)))
        wsLog.Cells(fila, "B").NumberFormat = "0.00"

        fila = fila + 1

    Next i

    wsLog.Cells(fila, "A").Value = "TOTAL GENERAL"
    wsLog.Cells(fila, "B").Value = totalGeneral
    wsLog.Cells(fila, "A").Font.Bold = True
    wsLog.Cells(fila, "B").Font.Bold = True
    wsLog.Cells(fila, "B").NumberFormat = "0.00"

    wsLog.Columns("A:H").AutoFit

End Sub

Private Function ValorNumerico(ByVal valor As Variant) As Double

    Dim txt As String
    Dim sepDecimal As String

    If IsError(valor) Then
        ValorNumerico = 0
        Exit Function
    End If

    If Trim(CStr(valor)) = "" Then
        ValorNumerico = 0
        Exit Function
    End If

    If IsNumeric(valor) Then
        ValorNumerico = CDbl(valor)
        Exit Function
    End If

    txt = Trim(CStr(valor))

    txt = Replace(txt, " ", "")
    txt = Replace(txt, "$", "")

    sepDecimal = Application.International(xlDecimalSeparator)

    If InStr(txt, ",") > 0 And InStr(txt, ".") = 0 Then
        txt = Replace(txt, ",", sepDecimal)
    End If

    If InStr(txt, ".") > 0 And InStr(txt, ",") = 0 Then
        txt = Replace(txt, ".", sepDecimal)
    End If

    If InStr(txt, ".") > 0 And InStr(txt, ",") > 0 Then
        txt = Replace(txt, ".", "")
        txt = Replace(txt, ",", sepDecimal)
    End If

    If IsNumeric(txt) Then
        ValorNumerico = CDbl(txt)
    Else
        ValorNumerico = 0
    End If

End Function

Private Function FormatearPagoNumeroCSV(ByVal valor As Double) As String

    Dim txt As String

    txt = Format$(valor, "0.00")
    txt = Replace(txt, ",", ".")

    FormatearPagoNumeroCSV = txt

End Function

Private Function CSVValue(ByVal valor As String) As String

    valor = Replace(valor, """", """""")

    If InStr(valor, ",") > 0 Or _
       InStr(valor, """") > 0 Or _
       InStr(valor, vbCr) > 0 Or _
       InStr(valor, vbLf) > 0 Then

        CSVValue = """" & valor & """"
    Else
        CSVValue = valor
    End If

End Function

Private Sub GuardarUTF8(ByVal rutaArchivo As String, ByVal contenido As String)

    Dim stream As Object

    Set stream = CreateObject("ADODB.Stream")

    With stream
        .Type = 2
        .Charset = "utf-8"
        .Open
        .WriteText contenido
        .SaveToFile rutaArchivo, 2
        .Close
    End With

End Sub

Private Function ObtenerPrimeraFechaKey(ByVal dictPrimeraFecha As Object) As String

    Dim key As Variant
    Dim menor As String

    menor = ""

    For Each key In dictPrimeraFecha.Keys
        If menor = "" Then
            menor = CStr(key)
        ElseIf CStr(key) < menor Then
            menor = CStr(key)
        End If
    Next key

    ObtenerPrimeraFechaKey = menor

End Function

Private Sub QuickSortStrings(ByRef arr() As String, ByVal first As Long, ByVal last As Long)

    Dim i As Long
    Dim j As Long
    Dim pivot As String
    Dim temp As String

    i = first
    j = last
    pivot = arr((first + last) \ 2)

    Do While i <= j

        Do While StrComp(arr(i), pivot, vbTextCompare) < 0
            i = i + 1
        Loop

        Do While StrComp(arr(j), pivot, vbTextCompare) > 0
            j = j - 1
        Loop

        If i <= j Then
            temp = arr(i)
            arr(i) = arr(j)
            arr(j) = temp

            i = i + 1
            j = j - 1
        End If

    Loop

    If first < j Then QuickSortStrings arr, first, j
    If i < last Then QuickSortStrings arr, i, last

End Sub

Private Sub PrepararHojaVerificacion()

    Dim wsLog As Worksheet

    On Error Resume Next
    Set wsLog = ThisWorkbook.Worksheets("Verificacion_CSV")
    On Error GoTo 0

    If wsLog Is Nothing Then
        Set wsLog = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        wsLog.Name = "Verificacion_CSV"
    Else
        wsLog.Cells.Clear
    End If

    With wsLog
        .Range("A1").Value = "Hoja"
        .Range("B1").Value = "Configuracion"
        .Range("C1").Value = "Fila"
        .Range("D1").Value = "Alerta"
        .Range("E1").Value = "Valor"
        .Range("F1").Value = "Desde / Día"
        .Range("G1").Value = "Hasta"
        .Range("H1").Value = "Fecha revisión"
        .Range("A1:H1").Font.Bold = True
    End With

End Sub

Private Sub RegistrarVerificacion(ByVal nombreHoja As String, _
                                  ByVal configuracion As String, _
                                  ByVal fila As Variant, _
                                  ByVal tipo As String, _
                                  ByVal detalle As String, _
                                  ByVal valor As Variant, _
                                  ByVal fechaAnterior As Variant, _
                                  ByVal fechaActual As Variant)

    Dim wsLog As Worksheet
    Dim nuevaFila As Long

    Set wsLog = ThisWorkbook.Worksheets("Verificacion_CSV")

    nuevaFila = wsLog.Cells(wsLog.Rows.Count, "A").End(xlUp).Row + 1

    With wsLog
        .Cells(nuevaFila, "A").Value = nombreHoja
        .Cells(nuevaFila, "B").Value = configuracion
        .Cells(nuevaFila, "C").Value = fila
        .Cells(nuevaFila, "D").Value = TextoCortoAlerta(tipo)
        .Cells(nuevaFila, "E").Value = valor
        .Cells(nuevaFila, "F").Value = fechaAnterior
        .Cells(nuevaFila, "G").Value = fechaActual
        .Cells(nuevaFila, "H").Value = Now
    End With

End Sub

Private Function TextoCortoAlerta(ByVal tipo As String) As String

    Select Case tipo

        Case "Hora mayor a 24"
            TextoCortoAlerta = "Hora > 24"

        Case "Cambio hora +1 detectado"
            TextoCortoAlerta = "Cambio +1"

        Case "Cambio hora -1 detectado"
            TextoCortoAlerta = "Cambio -1"

        Case "Fecha/hora duplicada"
            TextoCortoAlerta = "Hora duplicada"

        Case "Salto inesperado"
            TextoCortoAlerta = "Salto horario"

        Case "Resumen día cambio +1"
            TextoCortoAlerta = "Día 92 bloques"

        Case "Resumen día cambio -1"
            TextoCortoAlerta = "Día 100 bloques"

        Case "Resumen día revisar"
            TextoCortoAlerta = "Día revisar"

        Case "Fecha distinta para mismo bloque"
            TextoCortoAlerta = "Fecha distinta bloque"

        Case Else
            TextoCortoAlerta = tipo

    End Select

End Function

Private Sub RegistrarResumenDias(ByVal dictDia As Object)

    Dim key As Variant
    Dim partes() As String
    Dim nombreHoja As String
    Dim configuracion As String
    Dim fechaTxt As String
    Dim registros As Long

    For Each key In dictDia.Keys

        partes = Split(CStr(key), "|")

        nombreHoja = partes(0)
        configuracion = partes(1)
        fechaTxt = partes(2)
        registros = CLng(dictDia(key))

        Select Case registros

            Case 96
                ' Día normal

            Case 92
                RegistrarVerificacion nombreHoja, configuracion, "", _
                    "Resumen día cambio +1", "", _
                    registros, fechaTxt, ""

            Case 100
                RegistrarVerificacion nombreHoja, configuracion, "", _
                    "Resumen día cambio -1", "", _
                    registros, fechaTxt, ""

            Case Else
                RegistrarVerificacion nombreHoja, configuracion, "", _
                    "Resumen día revisar", "", _
                    registros, fechaTxt, ""

        End Select

    Next key

End Sub

Private Sub AjustarHojaVerificacion()

    Dim wsLog As Worksheet
    Dim ultimaFila As Long

    Set wsLog = ThisWorkbook.Worksheets("Verificacion_CSV")

    ultimaFila = wsLog.Cells(wsLog.Rows.Count, "A").End(xlUp).Row

    If ultimaFila = 1 Then
        wsLog.Range("A2").Value = "OK"
        wsLog.Range("D2").Value = "Sin alertas"
    End If

    wsLog.Columns("A:H").AutoFit

End Sub


```

### Módulo `D_Lee_medidas_origen`

```vb
Attribute VB_Name = "D_Lee_medidas_origen"
Option Explicit

Sub Cargar_Medidas_SAE_En_Medidores()

    Const FILA_ENCABEZADO_ORIGEN As Long = 1
    Const FILA_INICIO_DATOS_ORIGEN As Long = 2
    Const FILA_INICIO_DESTINO As Long = 3

    Const COLUMNA_INICIAL As Long = 1   ' Columna A
    Const COLUMNA_FINAL As Long = 9     ' Columna I

    Dim wbPrincipal As Workbook
    Dim wbOrigen As Workbook

    Dim wsDestino As Worksheet
    Dim wsOrigen As Worksheet

    Dim rutaCarpeta As String
    Dim rutaArchivo As String

    Dim ultimaFilaOrigen As Long
    Dim ultimaFilaDestino As Long
    Dim cantidadFilas As Long
    Dim cantidadColumnas As Long

    Dim arrDatos As Variant

    Dim modoCalculo As XlCalculation
    Dim archivoOrigenAbierto As Boolean
    Dim configuracionModificada As Boolean

    On Error GoTo ErrHandler

    '=====================================================
    ' ARCHIVO PRINCIPAL Y HOJA DESTINO
    '=====================================================
    Set wbPrincipal = ThisWorkbook
    Set wsDestino = wbPrincipal.Worksheets("Medidores")

    ' Validar que el archivo principal esté guardado
    If Len(wbPrincipal.Path) = 0 Then

        MsgBox _
            "Primero debes guardar el archivo que contiene la macro.", _
            vbExclamation

        Exit Sub

    End If

    '=====================================================
    ' RUTA DEL ARCHIVO DE ORIGEN
    '=====================================================
    rutaCarpeta = _
        wbPrincipal.Path & _
        Application.PathSeparator & _
        "Medidas"

    rutaArchivo = _
        rutaCarpeta & _
        Application.PathSeparator & _
        "Medidas_SAE.xlsx"

    ' Validar carpeta
    If Dir(rutaCarpeta, vbDirectory) = vbNullString Then

        MsgBox _
            "No se encontró la carpeta:" & vbCrLf & vbCrLf & _
            rutaCarpeta, _
            vbCritical

        Exit Sub

    End If

    ' Validar archivo
    If Dir(rutaArchivo, vbNormal) = vbNullString Then

        MsgBox _
            "No se encontró el archivo:" & vbCrLf & vbCrLf & _
            rutaArchivo, _
            vbCritical

        Exit Sub

    End If

    '=====================================================
    ' CONFIGURAR EXCEL PARA MAYOR VELOCIDAD
    '=====================================================
    With Application

        .ScreenUpdating = False
        .EnableEvents = False

        modoCalculo = .Calculation
        .Calculation = xlCalculationManual

        .DisplayAlerts = False
        .StatusBar = "Abriendo Medidas_SAE.xlsx..."

    End With

    configuracionModificada = True

    '=====================================================
    ' ABRIR ARCHIVO DE ORIGEN
    '=====================================================
    Set wbOrigen = Workbooks.Open( _
        Filename:=rutaArchivo, _
        UpdateLinks:=False, _
        ReadOnly:=True, _
        AddToMru:=False _
    )

    archivoOrigenAbierto = True

    ' Buscar hoja Medidas
    On Error Resume Next
    Set wsOrigen = wbOrigen.Worksheets("Medidas")
    On Error GoTo ErrHandler

    If wsOrigen Is Nothing Then

        Err.Raise _
            vbObjectError + 1000, _
            "Cargar_Medidas_SAE_En_Medidores", _
            "No existe la hoja 'Medidas' en el archivo Medidas_SAE.xlsx."

    End If

    '=====================================================
    ' DETERMINAR ÚLTIMA FILA DEL ORIGEN
    '=====================================================
    ultimaFilaOrigen = UltimaFilaEntreColumnas( _
        wsOrigen, _
        COLUMNA_INICIAL, _
        COLUMNA_FINAL _
    )

    If ultimaFilaOrigen < FILA_INICIO_DATOS_ORIGEN Then

        Err.Raise _
            vbObjectError + 1001, _
            "Cargar_Medidas_SAE_En_Medidores", _
            "La hoja 'Medidas' no contiene datos para copiar."

    End If

    Application.StatusBar = "Ordenando datos del archivo de origen..."

    '=====================================================
    ' ORDENAR DATOS EN EL ORIGEN
    '
    ' PRIMER CRITERIO:
    ' Columna G de A a Z
    '
    ' SEGUNDO CRITERIO:
    ' Columna F de menor a mayor
    '=====================================================
    With wsOrigen.Sort

        .SortFields.Clear

        ' Primer criterio: columna G de A a Z
        .SortFields.Add _
            key:=wsOrigen.Range( _
                "G" & FILA_INICIO_DATOS_ORIGEN & _
                ":G" & ultimaFilaOrigen _
            ), _
            SortOn:=xlSortOnValues, _
            Order:=xlAscending, _
            DataOption:=xlSortNormal

        ' Segundo criterio: columna F de menor a mayor
        .SortFields.Add _
            key:=wsOrigen.Range( _
                "F" & FILA_INICIO_DATOS_ORIGEN & _
                ":F" & ultimaFilaOrigen _
            ), _
            SortOn:=xlSortOnValues, _
            Order:=xlAscending, _
            DataOption:=xlSortNormal

        ' Se ordenan juntas las columnas A:I
        .SetRange wsOrigen.Range( _
            wsOrigen.Cells( _
                FILA_ENCABEZADO_ORIGEN, _
                COLUMNA_INICIAL _
            ), _
            wsOrigen.Cells( _
                ultimaFilaOrigen, _
                COLUMNA_FINAL _
            ) _
        )

        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply

    End With

    '=====================================================
    ' LEER DATOS ORDENADOS A:I
    '=====================================================
    cantidadFilas = _
        ultimaFilaOrigen - FILA_INICIO_DATOS_ORIGEN + 1

    cantidadColumnas = _
        COLUMNA_FINAL - COLUMNA_INICIAL + 1

    arrDatos = wsOrigen.Range( _
        wsOrigen.Cells( _
            FILA_INICIO_DATOS_ORIGEN, _
            COLUMNA_INICIAL _
        ), _
        wsOrigen.Cells( _
            ultimaFilaOrigen, _
            COLUMNA_FINAL _
        ) _
    ).Value2

    '=====================================================
    ' CERRAR ARCHIVO DE ORIGEN SIN GUARDAR
    '=====================================================
    wbOrigen.Close SaveChanges:=False
    archivoOrigenAbierto = False

    Set wsOrigen = Nothing
    Set wbOrigen = Nothing

    '=====================================================
    ' BORRAR DATOS ANTERIORES EN MEDIDORES
    ' DESDE A3:I HASTA LA ÚLTIMA FILA
    '=====================================================
    Application.StatusBar = "Borrando datos anteriores de Medidores..."

    ultimaFilaDestino = UltimaFilaEntreColumnas( _
        wsDestino, _
        COLUMNA_INICIAL, _
        COLUMNA_FINAL _
    )

    If ultimaFilaDestino >= FILA_INICIO_DESTINO Then

        wsDestino.Range( _
            wsDestino.Cells( _
                FILA_INICIO_DESTINO, _
                COLUMNA_INICIAL _
            ), _
            wsDestino.Cells( _
                ultimaFilaDestino, _
                COLUMNA_FINAL _
            ) _
        ).ClearContents

    End If

    '=====================================================
    ' PEGAR DATOS EN MEDIDORES DESDE A3
    '=====================================================
    Application.StatusBar = "Pegando datos en la hoja Medidores..."

    wsDestino.Cells( _
        FILA_INICIO_DESTINO, _
        COLUMNA_INICIAL _
    ).Resize( _
        cantidadFilas, _
        cantidadColumnas _
    ).Value2 = arrDatos

    Application.StatusBar = False

    MsgBox _
        "Carga terminada correctamente." & vbCrLf & vbCrLf & _
        "Archivo origen: Medidas_SAE.xlsx" & vbCrLf & _
        "Hoja origen: Medidas" & vbCrLf & _
        "Datos copiados: columnas A:I" & vbCrLf & _
        "Destino: Medidores desde A3" & vbCrLf & _
        "Registros copiados: " & cantidadFilas & vbCrLf & vbCrLf & _
        "Orden aplicado:" & vbCrLf & _
        "1. Columna G de A a Z" & vbCrLf & _
        "2. Columna F de menor a mayor", _
        vbInformation

Salida:

    Application.StatusBar = False

    If configuracionModificada Then

        With Application
            .ScreenUpdating = True
            .EnableEvents = True
            .Calculation = modoCalculo
            .DisplayAlerts = True
        End With

    End If

    Exit Sub

ErrHandler:

    If archivoOrigenAbierto Then

        On Error Resume Next
        wbOrigen.Close SaveChanges:=False
        On Error GoTo 0

    End If

    MsgBox _
        "Error al cargar Medidas_SAE.xlsx:" & vbCrLf & vbCrLf & _
        Err.Description, _
        vbCritical

    Resume Salida

End Sub


Private Function UltimaFilaEntreColumnas( _
    ByVal ws As Worksheet, _
    ByVal primeraColumna As Long, _
    ByVal ultimaColumna As Long _
) As Long

    Dim rangoBusqueda As Range
    Dim celdaEncontrada As Range

    Set rangoBusqueda = ws.Range( _
        ws.Cells(1, primeraColumna), _
        ws.Cells(ws.Rows.Count, ultimaColumna) _
    )

    Set celdaEncontrada = rangoBusqueda.Find( _
        What:="*", _
        After:=rangoBusqueda.Cells(1, 1), _
        LookIn:=xlFormulas, _
        LookAt:=xlPart, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlPrevious, _
        MatchCase:=False _
    )

    If celdaEncontrada Is Nothing Then
        UltimaFilaEntreColumnas = 0
    Else
        UltimaFilaEntreColumnas = celdaEncontrada.Row
    End If

End Function


```

### Módulo `E_lee_cmg_origen`

```vb
Attribute VB_Name = "E_lee_cmg_origen"
Option Explicit

Sub Cargar_CMg_Desde_Archivo()

    Const FILA_ENCABEZADO_ORIGEN As Long = 1
    Const FILA_INICIO_ORIGEN As Long = 2
    Const FILA_INICIO_DESTINO As Long = 2

    Const COLUMNA_INICIAL As Long = 1   ' A
    Const COLUMNA_FINAL As Long = 9     ' I

    Dim wbPrincipal As Workbook
    Dim wbOrigen As Workbook

    Dim wsDestino As Worksheet
    Dim wsOrigen As Worksheet

    Dim rutaCarpeta As String
    Dim rutaArchivo As String

    Dim ultimaFilaOrigen As Long
    Dim ultimaFilaDestino As Long

    Dim cantidadFilas As Long
    Dim cantidadColumnas As Long

    Dim arrDatos As Variant

    Dim modoCalculo As XlCalculation
    Dim archivoOrigenAbierto As Boolean
    Dim configuracionModificada As Boolean

    On Error GoTo ErrHandler

    '=====================================================
    ' ARCHIVO PRINCIPAL Y HOJA DESTINO
    '=====================================================
    Set wbPrincipal = ThisWorkbook
    Set wsDestino = wbPrincipal.Worksheets("CMg")

    ' El archivo principal debe estar guardado
    If Len(wbPrincipal.Path) = 0 Then

        MsgBox _
            "Primero debes guardar el archivo que contiene la macro.", _
            vbExclamation

        Exit Sub

    End If

    '=====================================================
    ' RUTA DEL ARCHIVO DE ORIGEN
    '=====================================================
    rutaCarpeta = _
        wbPrincipal.Path & _
        Application.PathSeparator & _
        "Cmg"

    rutaArchivo = _
        rutaCarpeta & _
        Application.PathSeparator & _
        "cmg.xlsx"

    ' Validar carpeta
    If Dir(rutaCarpeta, vbDirectory) = vbNullString Then

        MsgBox _
            "No se encontró la carpeta:" & vbCrLf & vbCrLf & _
            rutaCarpeta, _
            vbCritical

        Exit Sub

    End If

    ' Validar archivo
    If Dir(rutaArchivo, vbNormal) = vbNullString Then

        MsgBox _
            "No se encontró el archivo:" & vbCrLf & vbCrLf & _
            rutaArchivo, _
            vbCritical

        Exit Sub

    End If

    '=====================================================
    ' OPTIMIZAR EXCEL
    '=====================================================
    With Application

        .ScreenUpdating = False
        .EnableEvents = False

        modoCalculo = .Calculation
        .Calculation = xlCalculationManual

        .DisplayAlerts = False
        .StatusBar = "Abriendo cmg.xlsx..."

    End With

    configuracionModificada = True

    '=====================================================
    ' ABRIR ARCHIVO DE ORIGEN
    '=====================================================
    Set wbOrigen = Workbooks.Open( _
        Filename:=rutaArchivo, _
        UpdateLinks:=False, _
        ReadOnly:=True, _
        AddToMru:=False _
    )

    archivoOrigenAbierto = True

    ' Intentar utilizar la hoja CMg
    On Error Resume Next
    Set wsOrigen = wbOrigen.Worksheets("CMg")
    On Error GoTo ErrHandler

    ' Si no existe, usar la primera hoja
    If wsOrigen Is Nothing Then

        If wbOrigen.Worksheets.Count = 0 Then

            Err.Raise _
                vbObjectError + 1000, _
                "Cargar_CMg_Desde_Archivo", _
                "El archivo cmg.xlsx no contiene hojas."

        End If

        Set wsOrigen = wbOrigen.Worksheets(1)

    End If

    '=====================================================
    ' DETERMINAR ÚLTIMA FILA DEL ORIGEN
    ' CONSIDERANDO SOLAMENTE A:I
    '=====================================================
    ultimaFilaOrigen = UltimaFilaEntreColumnasCMg( _
        wsOrigen, _
        COLUMNA_INICIAL, _
        COLUMNA_FINAL _
    )

    If ultimaFilaOrigen < FILA_INICIO_ORIGEN Then

        Err.Raise _
            vbObjectError + 1001, _
            "Cargar_CMg_Desde_Archivo", _
            "El archivo cmg.xlsx no contiene datos desde la fila 2."

    End If

    Application.StatusBar = "Ordenando datos de cmg.xlsx..."

    '=====================================================
    ' ORDENAR EN EL ORIGEN
    '
    ' PRIMER CRITERIO:
    ' Columna D de A a Z
    '
    ' SEGUNDO CRITERIO:
    ' Columna H de menor a mayor
    '=====================================================
    With wsOrigen.Sort

        .SortFields.Clear

        ' Primer criterio: D de A a Z
        .SortFields.Add _
            key:=wsOrigen.Range( _
                "D" & FILA_INICIO_ORIGEN & _
                ":D" & ultimaFilaOrigen _
            ), _
            SortOn:=xlSortOnValues, _
            Order:=xlAscending, _
            DataOption:=xlSortNormal

        ' Segundo criterio: H de menor a mayor
        .SortFields.Add _
            key:=wsOrigen.Range( _
                "H" & FILA_INICIO_ORIGEN & _
                ":H" & ultimaFilaOrigen _
            ), _
            SortOn:=xlSortOnValues, _
            Order:=xlAscending, _
            DataOption:=xlSortNormal

        ' Ordenar juntas las columnas A:I
        .SetRange wsOrigen.Range( _
            wsOrigen.Cells( _
                FILA_ENCABEZADO_ORIGEN, _
                COLUMNA_INICIAL _
            ), _
            wsOrigen.Cells( _
                ultimaFilaOrigen, _
                COLUMNA_FINAL _
            ) _
        )

        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply

    End With

    '=====================================================
    ' LEER DATOS A:I SIN ENCABEZADOS
    '=====================================================
    cantidadFilas = _
        ultimaFilaOrigen - FILA_INICIO_ORIGEN + 1

    cantidadColumnas = _
        COLUMNA_FINAL - COLUMNA_INICIAL + 1

    arrDatos = wsOrigen.Range( _
        wsOrigen.Cells( _
            FILA_INICIO_ORIGEN, _
            COLUMNA_INICIAL _
        ), _
        wsOrigen.Cells( _
            ultimaFilaOrigen, _
            COLUMNA_FINAL _
        ) _
    ).Value2

    '=====================================================
    ' CERRAR ORIGEN SIN GUARDAR EL ORDENAMIENTO
    '=====================================================
    wbOrigen.Close SaveChanges:=False
    archivoOrigenAbierto = False

    Set wsOrigen = Nothing
    Set wbOrigen = Nothing

    '=====================================================
    ' BORRAR DATOS ANTERIORES EN EL DESTINO
    ' DESDE A2:I HACIA ABAJO
    '=====================================================
    Application.StatusBar = _
        "Borrando datos anteriores de la hoja CMg..."

    ultimaFilaDestino = UltimaFilaEntreColumnasCMg( _
        wsDestino, _
        COLUMNA_INICIAL, _
        COLUMNA_FINAL _
    )

    If ultimaFilaDestino >= FILA_INICIO_DESTINO Then

        wsDestino.Range( _
            wsDestino.Cells( _
                FILA_INICIO_DESTINO, _
                COLUMNA_INICIAL _
            ), _
            wsDestino.Cells( _
                ultimaFilaDestino, _
                COLUMNA_FINAL _
            ) _
        ).ClearContents

    End If

    '=====================================================
    ' PEGAR DATOS DESDE A2
    '=====================================================
    Application.StatusBar = _
        "Pegando datos en la hoja CMg..."

    wsDestino.Cells( _
        FILA_INICIO_DESTINO, _
        COLUMNA_INICIAL _
    ).Resize( _
        cantidadFilas, _
        cantidadColumnas _
    ).Value2 = arrDatos

    MsgBox _
        "Carga terminada correctamente." & vbCrLf & vbCrLf & _
        "Archivo origen: Cmg\cmg.xlsx" & vbCrLf & _
        "Hoja destino: CMg" & vbCrLf & _
        "Destino: A2:I" & vbCrLf & _
        "Registros copiados: " & cantidadFilas & vbCrLf & vbCrLf & _
        "Orden aplicado:" & vbCrLf & _
        "1. Columna D de A a Z" & vbCrLf & _
        "2. Columna H de menor a mayor", _
        vbInformation

Salida:

    Application.StatusBar = False

    If configuracionModificada Then

        With Application
            .ScreenUpdating = True
            .EnableEvents = True
            .Calculation = modoCalculo
            .DisplayAlerts = True
        End With

    End If

    Exit Sub

ErrHandler:

    If archivoOrigenAbierto Then

        On Error Resume Next
        wbOrigen.Close SaveChanges:=False
        On Error GoTo 0

    End If

    MsgBox _
        "Error al cargar cmg.xlsx:" & vbCrLf & vbCrLf & _
        Err.Description, _
        vbCritical

    Resume Salida

End Sub


Private Function UltimaFilaEntreColumnasCMg( _
    ByVal ws As Worksheet, _
    ByVal primeraColumna As Long, _
    ByVal ultimaColumna As Long _
) As Long

    Dim rangoBusqueda As Range
    Dim celdaEncontrada As Range

    Set rangoBusqueda = ws.Range( _
        ws.Cells(1, primeraColumna), _
        ws.Cells(ws.Rows.Count, ultimaColumna) _
    )

    Set celdaEncontrada = rangoBusqueda.Find( _
        What:="*", _
        After:=rangoBusqueda.Cells(1, 1), _
        LookIn:=xlFormulas, _
        LookAt:=xlPart, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlPrevious, _
        MatchCase:=False _
    )

    If celdaEncontrada Is Nothing Then
        UltimaFilaEntreColumnasCMg = 0
    Else
        UltimaFilaEntreColumnasCMg = celdaEncontrada.Row
    End If

End Function


```

### Módulo `F_Leer_FD`

```vb
Attribute VB_Name = "F_Leer_FD"
Option Explicit

Sub Cargar_SSCC_Desempeno_En_FD()

    Const FILA_INICIO_DATOS As Long = 12

    Dim wbDestino As Workbook
    Dim wbOrigen As Workbook

    Dim wsFD As Worksheet
    Dim wsCPF As Worksheet
    Dim wsCSF As Worksheet

    Dim rutaCarpeta As String
    Dim rutaArchivo As String

    Dim ultimaFilaCPF As Long
    Dim ultimaFilaCSF As Long

    Dim ultimaFilaDestinoCPF As Long
    Dim ultimaFilaDestinoCSF As Long

    Dim ultimaFilaFormulaAC As Long
    Dim ultimaFilaFormulaKM As Long
    Dim ultimaFilaFormulaQS As Long
    Dim ultimaFilaFormulaACAE As Long

    Dim ultimaFilaNuevaCPF As Long
    Dim ultimaFilaNuevaCSF As Long

    Dim cantidadFilasCPF As Long
    Dim cantidadFilasCSF As Long

    Dim arrCPF As Variant
    Dim arrCSF As Variant
    Dim arrEncabezado As Variant

    Dim plantillaAC As Variant
    Dim plantillaKM As Variant
    Dim plantillaQS As Variant
    Dim plantillaACAE As Variant

    Dim modoCalculo As XlCalculation
    Dim archivoAbierto As Boolean
    Dim configuracionModificada As Boolean

    On Error GoTo ErrHandler

    '=====================================================
    ' ARCHIVO Y HOJA DE DESTINO
    '=====================================================
    Set wbDestino = ThisWorkbook
    Set wsFD = wbDestino.Worksheets("FD")

    If Len(wbDestino.Path) = 0 Then

        MsgBox _
            "Primero debes guardar el archivo que contiene la macro.", _
            vbExclamation

        Exit Sub

    End If

    '=====================================================
    ' CARPETA Y ARCHIVO DE ORIGEN
    '=====================================================
    rutaCarpeta = _
        wbDestino.Path & _
        Application.PathSeparator & _
        "SSCC_Desempeño"

    If Dir(rutaCarpeta, vbDirectory) = vbNullString Then

        MsgBox _
            "No se encontró la carpeta:" & vbCrLf & vbCrLf & _
            rutaCarpeta, _
            vbCritical

        Exit Sub

    End If

    rutaArchivo = BuscarArchivoSSCCMasReciente(rutaCarpeta)

    If Len(rutaArchivo) = 0 Then

        MsgBox _
            "No se encontró ningún archivo que comience con:" & _
            vbCrLf & vbCrLf & _
            "SSCC_Desempeño_" & _
            vbCrLf & vbCrLf & _
            "Dentro de la carpeta:" & vbCrLf & _
            rutaCarpeta, _
            vbCritical

        Exit Sub

    End If

    '=====================================================
    ' GUARDAR PLANTILLAS DE FÓRMULAS DE LA FILA 12
    '=====================================================
    ValidarPlantillaFormulas wsFD.Range("A12:C12")
    ValidarPlantillaFormulas wsFD.Range("K12:M12")
    ValidarPlantillaFormulas wsFD.Range("Q12:S12")
    ValidarPlantillaFormulas wsFD.Range("AC12:AE12")

    plantillaAC = wsFD.Range("A12:C12").FormulaR1C1
    plantillaKM = wsFD.Range("K12:M12").FormulaR1C1
    plantillaQS = wsFD.Range("Q12:S12").FormulaR1C1
    plantillaACAE = wsFD.Range("AC12:AE12").FormulaR1C1

    ' Guardar hasta dónde llegaban las fórmulas anteriores
    ultimaFilaFormulaAC = UltimaFilaEnRangoSSCC(wsFD, "A", "C")
    ultimaFilaFormulaKM = UltimaFilaEnRangoSSCC(wsFD, "K", "M")
    ultimaFilaFormulaQS = UltimaFilaEnRangoSSCC(wsFD, "Q", "S")
    ultimaFilaFormulaACAE = UltimaFilaEnRangoSSCC(wsFD, "AC", "AE")

    '=====================================================
    ' OPTIMIZAR EXCEL
    '=====================================================
    With Application

        .ScreenUpdating = False
        .EnableEvents = False

        modoCalculo = .Calculation
        .Calculation = xlCalculationManual

        .DisplayAlerts = False
        .StatusBar = "Abriendo archivo SSCC Desempeño..."

    End With

    configuracionModificada = True

    '=====================================================
    ' ABRIR ARCHIVO DE ORIGEN
    '=====================================================
    Set wbOrigen = Workbooks.Open( _
        Filename:=rutaArchivo, _
        UpdateLinks:=False, _
        ReadOnly:=True, _
        AddToMru:=False _
    )

    archivoAbierto = True

    '=====================================================
    ' BUSCAR HOJAS DE ORIGEN
    '=====================================================
    On Error Resume Next

    Set wsCPF = wbOrigen.Worksheets("CPF Horario")
    Set wsCSF = wbOrigen.Worksheets("CSF Horario")

    On Error GoTo ErrHandler

    If wsCPF Is Nothing Then

        Err.Raise _
            vbObjectError + 1000, _
            "Cargar_SSCC_Desempeno_En_FD", _
            "No existe la hoja 'CPF Horario' en el archivo de origen."

    End If

    If wsCSF Is Nothing Then

        Err.Raise _
            vbObjectError + 1001, _
            "Cargar_SSCC_Desempeno_En_FD", _
            "No existe la hoja 'CSF Horario' en el archivo de origen."

    End If

    '=====================================================
    ' LEER Y FILTRAR CPF HORARIO
    '
    ' Datos originales: B12:J
    ' Filtro: columna D contiene BESS o SAE
    '=====================================================
    Application.StatusBar = _
        "Filtrando datos de CPF Horario..."

    ultimaFilaCPF = UltimaFilaEnRangoSSCC( _
        wsCPF, _
        "B", _
        "J" _
    )

    If ultimaFilaCPF >= FILA_INICIO_DATOS Then

        arrCPF = FiltrarFilasBESSoSAE( _
            wsCPF, _
            FILA_INICIO_DATOS, _
            ultimaFilaCPF, _
            2, _
            10, _
            cantidadFilasCPF _
        )

    Else

        cantidadFilasCPF = 0

    End If

    ' B9:C9 de CPF se copia en ambos sectores
    arrEncabezado = wsCPF.Range("B9:C9").Value2

    '=====================================================
    ' LEER Y FILTRAR CSF HORARIO
    '
    ' Datos originales: B12:H
    ' Filtro: columna D contiene BESS o SAE
    '=====================================================
    Application.StatusBar = _
        "Filtrando datos de CSF Horario..."

    ultimaFilaCSF = UltimaFilaEnRangoSSCC( _
        wsCSF, _
        "B", _
        "H" _
    )

    If ultimaFilaCSF >= FILA_INICIO_DATOS Then

        arrCSF = FiltrarFilasBESSoSAE( _
            wsCSF, _
            FILA_INICIO_DATOS, _
            ultimaFilaCSF, _
            2, _
            8, _
            cantidadFilasCSF _
        )

    Else

        cantidadFilasCSF = 0

    End If

    '=====================================================
    ' CERRAR ORIGEN SIN GUARDAR
    '=====================================================
    wbOrigen.Close SaveChanges:=False
    archivoAbierto = False

    Set wsCPF = Nothing
    Set wsCSF = Nothing
    Set wbOrigen = Nothing

    '=====================================================
    ' ENCABEZADOS
    '=====================================================
    wsFD.Range("D9:E9").Value2 = arrEncabezado
    wsFD.Range("T9:U9").Value2 = arrEncabezado

    '=====================================================
    ' LIMPIAR BLOQUE CSF: D12:J
    '=====================================================
    Application.StatusBar = "Limpiando bloque CSF..."

    ultimaFilaDestinoCSF = UltimaFilaEnRangoSSCC( _
        wsFD, _
        "D", _
        "J" _
    )

    If ultimaFilaDestinoCSF >= FILA_INICIO_DATOS Then

        wsFD.Range( _
            "D" & FILA_INICIO_DATOS & _
            ":J" & ultimaFilaDestinoCSF _
        ).ClearContents

    End If

    '=====================================================
    ' PEGAR FILAS FILTRADAS DE CSF
    '=====================================================
    If cantidadFilasCSF > 0 Then

        wsFD.Range("D" & FILA_INICIO_DATOS).Resize( _
            cantidadFilasCSF, _
            7 _
        ).Value2 = arrCSF

        ultimaFilaNuevaCSF = _
            FILA_INICIO_DATOS + cantidadFilasCSF - 1

    Else

        ultimaFilaNuevaCSF = FILA_INICIO_DATOS - 1

    End If

    '=====================================================
    ' AJUSTAR FÓRMULAS ASOCIADAS A CSF
    '
    ' A:C y K:M
    '=====================================================
    AjustarBloqueFormulas _
        wsFD, _
        "A", _
        "C", _
        FILA_INICIO_DATOS, _
        cantidadFilasCSF, _
        plantillaAC, _
        ultimaFilaFormulaAC

    AjustarBloqueFormulas _
        wsFD, _
        "K", _
        "M", _
        FILA_INICIO_DATOS, _
        cantidadFilasCSF, _
        plantillaKM, _
        ultimaFilaFormulaKM

    '=====================================================
    ' LIMPIAR BLOQUE CPF: T12:AB
    '=====================================================
    Application.StatusBar = "Limpiando bloque CPF..."

    ultimaFilaDestinoCPF = UltimaFilaEnRangoSSCC( _
        wsFD, _
        "T", _
        "AB" _
    )

    If ultimaFilaDestinoCPF >= FILA_INICIO_DATOS Then

        wsFD.Range( _
            "T" & FILA_INICIO_DATOS & _
            ":AB" & ultimaFilaDestinoCPF _
        ).ClearContents

    End If

    '=====================================================
    ' PEGAR FILAS FILTRADAS DE CPF
    '=====================================================
    If cantidadFilasCPF > 0 Then

        wsFD.Range("T" & FILA_INICIO_DATOS).Resize( _
            cantidadFilasCPF, _
            9 _
        ).Value2 = arrCPF

        ultimaFilaNuevaCPF = _
            FILA_INICIO_DATOS + cantidadFilasCPF - 1

    Else

        ultimaFilaNuevaCPF = FILA_INICIO_DATOS - 1

    End If

    '=====================================================
    ' AJUSTAR FÓRMULAS ASOCIADAS A CPF
    '
    ' Q:S y AC:AE
    '=====================================================
    AjustarBloqueFormulas _
        wsFD, _
        "Q", _
        "S", _
        FILA_INICIO_DATOS, _
        cantidadFilasCPF, _
        plantillaQS, _
        ultimaFilaFormulaQS

    AjustarBloqueFormulas _
        wsFD, _
        "AC", _
        "AE", _
        FILA_INICIO_DATOS, _
        cantidadFilasCPF, _
        plantillaACAE, _
        ultimaFilaFormulaACAE

    wsFD.Calculate

    MsgBox _
        "Carga terminada correctamente." & vbCrLf & vbCrLf & _
        "Filtro aplicado en columna D del origen:" & vbCrLf & _
        "Contiene BESS o SAE" & vbCrLf & vbCrLf & _
        "Filas CSF traspasadas: " & cantidadFilasCSF & vbCrLf & _
        "Destino: FD!D12:J" & vbCrLf & _
        "Fórmulas: A:C y K:M" & vbCrLf & vbCrLf & _
        "Filas CPF traspasadas: " & cantidadFilasCPF & vbCrLf & _
        "Destino: FD!T12:AB" & vbCrLf & _
        "Fórmulas: Q:S y AC:AE", _
        vbInformation

Salida:

    Application.StatusBar = False

    If configuracionModificada Then

        With Application
            .ScreenUpdating = True
            .EnableEvents = True
            .Calculation = modoCalculo
            .DisplayAlerts = True
        End With

    End If

    Exit Sub

ErrHandler:

    If archivoAbierto Then

        On Error Resume Next
        wbOrigen.Close SaveChanges:=False
        On Error GoTo 0

    End If

    MsgBox _
        "Error al cargar SSCC Desempeño:" & vbCrLf & vbCrLf & _
        Err.Description, _
        vbCritical

    Resume Salida

End Sub


Private Function FiltrarFilasBESSoSAE( _
    ByVal ws As Worksheet, _
    ByVal filaInicial As Long, _
    ByVal filaFinal As Long, _
    ByVal columnaInicial As Long, _
    ByVal columnaFinal As Long, _
    ByRef cantidadFiltrada As Long _
) As Variant

    Dim arrOrigen As Variant
    Dim arrFiltrado() As Variant

    Dim indiceColumnaD As Long
    Dim numeroColumnas As Long

    Dim i As Long
    Dim j As Long
    Dim filaDestino As Long

    numeroColumnas = _
        columnaFinal - columnaInicial + 1

    ' Posición de la columna D dentro del arreglo
    indiceColumnaD = _
        4 - columnaInicial + 1

    arrOrigen = ws.Range( _
        ws.Cells(filaInicial, columnaInicial), _
        ws.Cells(filaFinal, columnaFinal) _
    ).Value2

    cantidadFiltrada = 0

    ' Primera pasada: contar coincidencias
    For i = 1 To UBound(arrOrigen, 1)

        If EsBESSoSAE(arrOrigen(i, indiceColumnaD)) Then
            cantidadFiltrada = cantidadFiltrada + 1
        End If

    Next i

    If cantidadFiltrada = 0 Then
        FiltrarFilasBESSoSAE = Empty
        Exit Function
    End If

    ReDim arrFiltrado( _
        1 To cantidadFiltrada, _
        1 To numeroColumnas _
    )

    filaDestino = 0

    ' Segunda pasada: copiar filas coincidentes
    For i = 1 To UBound(arrOrigen, 1)

        If EsBESSoSAE(arrOrigen(i, indiceColumnaD)) Then

            filaDestino = filaDestino + 1

            For j = 1 To numeroColumnas
                arrFiltrado(filaDestino, j) = arrOrigen(i, j)
            Next j

        End If

    Next i

    FiltrarFilasBESSoSAE = arrFiltrado

End Function


Private Function EsBESSoSAE( _
    ByVal valor As Variant _
) As Boolean

    Dim texto As String

    If IsError(valor) Then Exit Function
    If IsEmpty(valor) Then Exit Function

    texto = Trim$(CStr(valor))

    If Len(texto) = 0 Then Exit Function

    EsBESSoSAE = _
        InStr(1, texto, "BESS", vbTextCompare) > 0 Or _
        InStr(1, texto, "SAE", vbTextCompare) > 0

End Function


Private Sub AjustarBloqueFormulas( _
    ByVal ws As Worksheet, _
    ByVal columnaInicial As String, _
    ByVal columnaFinal As String, _
    ByVal filaInicial As Long, _
    ByVal cantidadFilas As Long, _
    ByVal plantilla As Variant, _
    ByVal ultimaFilaAnterior As Long _
)

    Dim ultimaFilaNueva As Long

    If cantidadFilas > 0 Then

        ultimaFilaNueva = _
            filaInicial + cantidadFilas - 1

        ' Restaurar fórmula modelo
        ws.Range( _
            columnaInicial & filaInicial & _
            ":" & columnaFinal & filaInicial _
        ).FormulaR1C1 = plantilla

        ' Extender fórmulas
        If ultimaFilaNueva > filaInicial Then

            ws.Range( _
                columnaInicial & filaInicial & _
                ":" & columnaFinal & ultimaFilaNueva _
            ).FillDown

        End If

        ' Borrar fórmulas sobrantes
        If ultimaFilaAnterior > ultimaFilaNueva Then

            ws.Range( _
                columnaInicial & ultimaFilaNueva + 1 & _
                ":" & columnaFinal & ultimaFilaAnterior _
            ).ClearContents

        End If

    Else

        ' No hubo filas que cumplieran el filtro
        If ultimaFilaAnterior >= filaInicial Then

            ws.Range( _
                columnaInicial & filaInicial & _
                ":" & columnaFinal & ultimaFilaAnterior _
            ).ClearContents

        End If

    End If

End Sub


Private Sub ValidarPlantillaFormulas( _
    ByVal rango As Range _
)

    Dim celda As Range

    For Each celda In rango.Cells

        If Not celda.HasFormula Then

            Err.Raise _
                vbObjectError + 1100, _
                "ValidarPlantillaFormulas", _
                "La celda " & celda.Address(False, False) & _
                " no contiene una fórmula para usar como plantilla."

        End If

    Next celda

End Sub


Private Function BuscarArchivoSSCCMasReciente( _
    ByVal rutaCarpeta As String _
) As String

    Dim nombreArchivo As String
    Dim rutaCompleta As String
    Dim rutaSeleccionada As String

    Dim extensionArchivo As String
    Dim fechaArchivo As Date
    Dim fechaMasReciente As Date

    nombreArchivo = Dir( _
        rutaCarpeta & _
        Application.PathSeparator & _
        "SSCC_Desempeño_*.*" _
    )

    Do While Len(nombreArchivo) > 0

        If Left$(nombreArchivo, 2) <> "~$" Then

            If InStrRev(nombreArchivo, ".") > 0 Then

                extensionArchivo = LCase$(Mid$( _
                    nombreArchivo, _
                    InStrRev(nombreArchivo, ".") _
                ))

                Select Case extensionArchivo

                    Case ".xlsx", ".xlsm", ".xlsb", ".xls"

                        rutaCompleta = _
                            rutaCarpeta & _
                            Application.PathSeparator & _
                            nombreArchivo

                        fechaArchivo = FileDateTime(rutaCompleta)

                        If Len(rutaSeleccionada) = 0 Or _
                           fechaArchivo > fechaMasReciente Then

                            rutaSeleccionada = rutaCompleta
                            fechaMasReciente = fechaArchivo

                        End If

                End Select

            End If

        End If

        nombreArchivo = Dir

    Loop

    BuscarArchivoSSCCMasReciente = rutaSeleccionada

End Function


Private Function UltimaFilaEnRangoSSCC( _
    ByVal ws As Worksheet, _
    ByVal columnaInicial As String, _
    ByVal columnaFinal As String _
) As Long

    Dim rangoBusqueda As Range
    Dim celdaEncontrada As Range

    Set rangoBusqueda = ws.Range( _
        columnaInicial & "1:" & _
        columnaFinal & ws.Rows.Count _
    )

    Set celdaEncontrada = rangoBusqueda.Find( _
        What:="*", _
        After:=rangoBusqueda.Cells(1, 1), _
        LookIn:=xlFormulas, _
        LookAt:=xlPart, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlPrevious, _
        MatchCase:=False _
    )

    If celdaEncontrada Is Nothing Then
        UltimaFilaEnRangoSSCC = 0
    Else
        UltimaFilaEnRangoSSCC = celdaEncontrada.Row
    End If

End Function


```

### Módulo `G_Lee_Subastas`

```vb
Attribute VB_Name = "G_Lee_Subastas"
Option Explicit

Sub Cargar_Remuneracion_Subastas_Rapido()

    Const FILA_INICIO As Long = 3

    Dim wbDestino As Workbook
    Dim wsDestino As Worksheet

    Dim rutaArchivo As String
    Dim sql As String

    Dim conexion As Object
    Dim registros As Object

    Dim datosADO As Variant
    Dim arrBL() As Variant
    Dim arrO() As Variant
    Dim arrP() As Variant
    Dim arrQ() As Variant

    Dim plantillaMN As Variant

    Dim ultimaFilaDatosAnterior As Long
    Dim ultimaFilaFormulasAnterior As Long
    Dim ultimaFilaNueva As Long

    Dim cantidadFilas As Long
    Dim i As Long
    Dim j As Long

    Dim modoCalculo As XlCalculation
    Dim configuracionModificada As Boolean

    On Error GoTo ErrHandler

    '=====================================================
    ' ARCHIVO Y HOJA DESTINO
    '=====================================================
    Set wbDestino = ThisWorkbook
    Set wsDestino = wbDestino.Worksheets("Subastas")

    If Len(wbDestino.Path) = 0 Then

        MsgBox _
            "Primero debes guardar el archivo que contiene la macro.", _
            vbExclamation

        Exit Sub

    End If

    '=====================================================
    ' BUSCAR ARCHIVO DE ORIGEN
    '=====================================================
    rutaArchivo = BuscarArchivoSubastasMasRecienteRapido( _
        wbDestino.Path, _
        wbDestino.FullName _
    )

    If Len(rutaArchivo) = 0 Then

        MsgBox _
            "No se encontró ningún archivo que comience con:" & _
            vbCrLf & vbCrLf & _
            "3_REMUNERACIÓN_SUBASTAS_E_ID_" & _
            vbCrLf & vbCrLf & _
            "En la carpeta:" & vbCrLf & _
            wbDestino.Path, _
            vbCritical

        Exit Sub

    End If

    '=====================================================
    ' GUARDAR FÓRMULAS MODELO DE M3:N3
    '=====================================================
    ValidarFormulasModeloSubastasRapido _
        wsDestino.Range("M3:N3")

    plantillaMN = wsDestino.Range("M3:N3").FormulaR1C1

    ultimaFilaFormulasAnterior = _
        UltimaFilaEnColumnasSubastasRapido( _
            wsDestino, _
            Array("M", "N") _
        )

    ultimaFilaDatosAnterior = _
        UltimaFilaEnColumnasSubastasRapido( _
            wsDestino, _
            Array( _
                "B", "C", "D", "E", "F", "G", _
                "H", "I", "J", "K", "L", _
                "O", "P", "Q" _
            ) _
        )

    '=====================================================
    ' OPTIMIZAR EXCEL
    '=====================================================
    With Application

        .ScreenUpdating = False
        .EnableEvents = False
        .DisplayAlerts = False

        modoCalculo = .Calculation
        .Calculation = xlCalculationManual

        .StatusBar = _
            "Leyendo archivo de remuneración de subastas..."

    End With

    configuracionModificada = True

    '=====================================================
    ' ABRIR CONEXIÓN ADO
    '
    ' El archivo se lee directamente sin abrirlo en Excel.
    '=====================================================
    Set conexion = AbrirConexionExcelSubastas(rutaArchivo)

    '=====================================================
    ' CONSULTA
    '
    ' El rango DB!B:Y se interpreta así:
    '
    ' F1  = B
    ' F2  = C
    ' ...
    ' F10 = K  <- columna utilizada para filtrar
    ' F11 = L
    ' F15 = P
    ' F21 = V
    ' F24 = Y
    '
    ' Solo selecciona filas donde K contiene:
    ' BESS o SAE
    '=====================================================
    sql = _
        "SELECT " & _
        "[F1], [F2], [F3], [F4], [F5], [F6], " & _
        "[F7], [F8], [F9], [F10], [F11], " & _
        "[F15], [F24], [F21] " & _
        "FROM [DB$B3:Y1048576] " & _
        "WHERE " & _
        "([F10] LIKE '%BESS%' OR [F10] LIKE '%SAE%')"

    Set registros = CreateObject("ADODB.Recordset")

    registros.CursorLocation = 3

    registros.Open _
        sql, _
        conexion, _
        0, _
        1, _
        1

    '=====================================================
    ' LEER RESULTADOS
    '=====================================================
    If registros.EOF Then

        cantidadFilas = 0

    Else

        datosADO = registros.GetRows
        cantidadFilas = UBound(datosADO, 2) + 1

    End If

    registros.Close
    conexion.Close

    Set registros = Nothing
    Set conexion = Nothing

    '=====================================================
    ' PREPARAR ARREGLOS DE DESTINO
    '=====================================================
    If cantidadFilas > 0 Then

        ReDim arrBL(1 To cantidadFilas, 1 To 11)
        ReDim arrO(1 To cantidadFilas, 1 To 1)
        ReDim arrP(1 To cantidadFilas, 1 To 1)
        ReDim arrQ(1 To cantidadFilas, 1 To 1)

        For i = 0 To cantidadFilas - 1

            '-------------------------------------------------
            ' Origen B:L -> destino B:L
            '-------------------------------------------------
            For j = 0 To 10
                arrBL(i + 1, j + 1) = datosADO(j, i)
            Next j

            '-------------------------------------------------
            ' Origen P -> destino O
            ' Campo número 12 del resultado
            '-------------------------------------------------
            arrO(i + 1, 1) = datosADO(11, i)

            '-------------------------------------------------
            ' Origen Y -> destino P
            ' Campo número 13 del resultado
            '-------------------------------------------------
            arrP(i + 1, 1) = datosADO(12, i)

            '-------------------------------------------------
            ' Origen V -> destino Q
            ' Campo número 14 del resultado
            '-------------------------------------------------
            arrQ(i + 1, 1) = datosADO(13, i)

        Next i

        ultimaFilaNueva = _
            FILA_INICIO + cantidadFilas - 1

    Else

        ultimaFilaNueva = FILA_INICIO - 1

    End If

    '=====================================================
    ' BORRAR DATOS ANTERIORES
    '=====================================================
    Application.StatusBar = _
        "Limpiando datos anteriores de Subastas..."

    If ultimaFilaDatosAnterior >= FILA_INICIO Then

        wsDestino.Range( _
            "B" & FILA_INICIO & _
            ":L" & ultimaFilaDatosAnterior _
        ).ClearContents

        wsDestino.Range( _
            "O" & FILA_INICIO & _
            ":Q" & ultimaFilaDatosAnterior _
        ).ClearContents

    End If

    '=====================================================
    ' PEGAR DATOS FILTRADOS
    '=====================================================
    If cantidadFilas > 0 Then

        Application.StatusBar = _
            "Pegando " & cantidadFilas & _
            " filas filtradas en Subastas..."

        ' DB B:L -> Subastas B:L
        wsDestino.Range("B" & FILA_INICIO).Resize( _
            cantidadFilas, _
            11 _
        ).Value2 = arrBL

        ' DB P -> Subastas O
        wsDestino.Range("O" & FILA_INICIO).Resize( _
            cantidadFilas, _
            1 _
        ).Value2 = arrO

        ' DB Y -> Subastas P
        wsDestino.Range("P" & FILA_INICIO).Resize( _
            cantidadFilas, _
            1 _
        ).Value2 = arrP

        ' DB V -> Subastas Q
        wsDestino.Range("Q" & FILA_INICIO).Resize( _
            cantidadFilas, _
            1 _
        ).Value2 = arrQ

    End If

    '=====================================================
    ' AJUSTAR FÓRMULAS M:N
    '=====================================================
    Application.StatusBar = _
        "Ajustando fórmulas de M:N..."

    AjustarFormulasMNSubastasRapido _
        wsDestino, _
        FILA_INICIO, _
        cantidadFilas, _
        ultimaFilaFormulasAnterior, _
        plantillaMN

    ' Calcular solamente las fórmulas importadas
    If cantidadFilas > 0 Then

        wsDestino.Range( _
            "M" & FILA_INICIO & _
            ":N" & ultimaFilaNueva _
        ).Calculate

    End If

    Application.StatusBar = False

    MsgBox _
        "Importación terminada correctamente." & vbCrLf & vbCrLf & _
        "Archivo leído:" & vbCrLf & _
        Dir(rutaArchivo) & vbCrLf & vbCrLf & _
        "Filtro aplicado:" & vbCrLf & _
        "DB!K contiene BESS o SAE" & vbCrLf & vbCrLf & _
        "Filas importadas: " & cantidadFilas & vbCrLf & vbCrLf & _
        "Correspondencia:" & vbCrLf & _
        "DB B:L  ?  Subastas B:L" & vbCrLf & _
        "DB P    ?  Subastas O" & vbCrLf & _
        "DB Y    ?  Subastas P" & vbCrLf & _
        "DB V    ?  Subastas Q" & vbCrLf & _
        "Fórmulas M:N ajustadas automáticamente.", _
        vbInformation

Salida:

    Application.StatusBar = False

    If configuracionModificada Then

        With Application
            .ScreenUpdating = True
            .EnableEvents = True
            .DisplayAlerts = True
            .Calculation = modoCalculo
        End With

    End If

    Exit Sub

ErrHandler:

    On Error Resume Next

    If Not registros Is Nothing Then
        If registros.State <> 0 Then registros.Close
    End If

    If Not conexion Is Nothing Then
        If conexion.State <> 0 Then conexion.Close
    End If

    On Error GoTo 0

    MsgBox _
        "Error al importar la remuneración de subastas:" & _
        vbCrLf & vbCrLf & _
        Err.Description, _
        vbCritical

    Resume Salida

End Sub


Private Function AbrirConexionExcelSubastas( _
    ByVal rutaArchivo As String _
) As Object

    Dim conexion As Object
    Dim proveedores As Variant
    Dim proveedor As Variant

    Dim propiedadesExcel As String
    Dim cadenaConexion As String
    Dim detalleError As String

    proveedores = Array( _
        "Microsoft.ACE.OLEDB.16.0", _
        "Microsoft.ACE.OLEDB.12.0" _
    )

    propiedadesExcel = _
        PropiedadesADOExcelSubastas(rutaArchivo)

    For Each proveedor In proveedores

        Set conexion = CreateObject("ADODB.Connection")

        cadenaConexion = _
            "Provider=" & CStr(proveedor) & ";" & _
            "Data Source=" & rutaArchivo & ";" & _
            "Extended Properties=""" & _
            propiedadesExcel & """;"

        On Error Resume Next

        Err.Clear
        conexion.Open cadenaConexion

        If Err.Number = 0 Then

            On Error GoTo 0
            Set AbrirConexionExcelSubastas = conexion
            Exit Function

        End If

        detalleError = Err.Description

        On Error GoTo 0

        Set conexion = Nothing

    Next proveedor

    Err.Raise _
        vbObjectError + 1200, _
        "AbrirConexionExcelSubastas", _
        "No fue posible leer el archivo mediante ADO." & _
        vbCrLf & vbCrLf & _
        "Verifica que esté instalado Microsoft Access " & _
        "Database Engine." & vbCrLf & vbCrLf & _
        "Detalle: " & detalleError

End Function


Private Function PropiedadesADOExcelSubastas( _
    ByVal rutaArchivo As String _
) As String

    Dim extensionArchivo As String

    extensionArchivo = LCase$(Mid$( _
        rutaArchivo, _
        InStrRev(rutaArchivo, ".") _
    ))

    Select Case extensionArchivo

        Case ".xls"

            PropiedadesADOExcelSubastas = _
                "Excel 8.0;HDR=NO;IMEX=1"

        Case ".xlsb"

            PropiedadesADOExcelSubastas = _
                "Excel 12.0;HDR=NO;IMEX=1"

        Case Else

            PropiedadesADOExcelSubastas = _
                "Excel 12.0 Xml;HDR=NO;IMEX=1"

    End Select

End Function


Private Sub AjustarFormulasMNSubastasRapido( _
    ByVal ws As Worksheet, _
    ByVal filaInicial As Long, _
    ByVal cantidadFilas As Long, _
    ByVal ultimaFilaAnterior As Long, _
    ByVal plantillaMN As Variant _
)

    Dim ultimaFilaNueva As Long

    If cantidadFilas > 0 Then

        ultimaFilaNueva = _
            filaInicial + cantidadFilas - 1

        ' Restaurar las fórmulas modelo
        ws.Range( _
            "M" & filaInicial & _
            ":N" & filaInicial _
        ).FormulaR1C1 = plantillaMN

        ' Extender fórmulas
        If ultimaFilaNueva > filaInicial Then

            ws.Range( _
                "M" & filaInicial & _
                ":N" & ultimaFilaNueva _
            ).FillDown

        End If

        ' Eliminar fórmulas sobrantes
        If ultimaFilaAnterior > ultimaFilaNueva Then

            ws.Range( _
                "M" & ultimaFilaNueva + 1 & _
                ":N" & ultimaFilaAnterior _
            ).ClearContents

        End If

    Else

        ' Se conserva M3:N3 como fila modelo para futuras cargas
        ws.Range( _
            "M" & filaInicial & _
            ":N" & filaInicial _
        ).FormulaR1C1 = plantillaMN

        ' Borrar fórmulas sobrantes desde la fila 4
        If ultimaFilaAnterior > filaInicial Then

            ws.Range( _
                "M" & filaInicial + 1 & _
                ":N" & ultimaFilaAnterior _
            ).ClearContents

        End If

    End If

End Sub


Private Sub ValidarFormulasModeloSubastasRapido( _
    ByVal rangoModelo As Range _
)

    Dim celda As Range

    For Each celda In rangoModelo.Cells

        If Not celda.HasFormula Then

            Err.Raise _
                vbObjectError + 1100, _
                "ValidarFormulasModeloSubastasRapido", _
                "La celda " & _
                celda.Address(False, False) & _
                " no contiene una fórmula." & vbCrLf & _
                "M3 y N3 deben contener las fórmulas modelo."

        End If

    Next celda

End Sub


Private Function BuscarArchivoSubastasMasRecienteRapido( _
    ByVal rutaCarpeta As String, _
    ByVal archivoExcluir As String _
) As String

    Dim nombreArchivo As String
    Dim rutaCompleta As String
    Dim rutaSeleccionada As String

    Dim extensionArchivo As String
    Dim fechaArchivo As Date
    Dim fechaMasReciente As Date

    nombreArchivo = Dir( _
        rutaCarpeta & _
        Application.PathSeparator & _
        "3_REMUNERACIÓN_SUBASTAS_E_ID_*.*" _
    )

    Do While Len(nombreArchivo) > 0

        If Left$(nombreArchivo, 2) <> "~$" Then

            rutaCompleta = _
                rutaCarpeta & _
                Application.PathSeparator & _
                nombreArchivo

            If StrComp( _
                rutaCompleta, _
                archivoExcluir, _
                vbTextCompare _
            ) <> 0 Then

                If InStrRev(nombreArchivo, ".") > 0 Then

                    extensionArchivo = LCase$(Mid$( _
                        nombreArchivo, _
                        InStrRev(nombreArchivo, ".") _
                    ))

                    Select Case extensionArchivo

                        Case ".xlsx", ".xlsm", ".xlsb", ".xls"

                            fechaArchivo = _
                                FileDateTime(rutaCompleta)

                            If Len(rutaSeleccionada) = 0 Or _
                               fechaArchivo > fechaMasReciente Then

                                rutaSeleccionada = rutaCompleta
                                fechaMasReciente = fechaArchivo

                            End If

                    End Select

                End If

            End If

        End If

        nombreArchivo = Dir

    Loop

    BuscarArchivoSubastasMasRecienteRapido = _
        rutaSeleccionada

End Function


Private Function UltimaFilaEnColumnasSubastasRapido( _
    ByVal ws As Worksheet, _
    ByVal columnas As Variant _
) As Long

    Dim columna As Variant
    Dim celdaEncontrada As Range
    Dim maximaFila As Long

    maximaFila = 0

    For Each columna In columnas

        Set celdaEncontrada = _
            ws.Columns(CStr(columna)).Find( _
                What:="*", _
                LookIn:=xlFormulas, _
                LookAt:=xlPart, _
                SearchOrder:=xlByRows, _
                SearchDirection:=xlPrevious, _
                MatchCase:=False _
            )

        If Not celdaEncontrada Is Nothing Then

            If celdaEncontrada.Row > maximaFila Then
                maximaFila = celdaEncontrada.Row
            End If

        End If

        Set celdaEncontrada = Nothing

    Next columna

    UltimaFilaEnColumnasSubastasRapido = maximaFila

End Function


```

### Módulo `H_Leer_Ofertas`

```vb
Attribute VB_Name = "H_Leer_Ofertas"
Option Explicit

Sub Generar_Resumen_Ofertas_SSCC()

    Const NOMBRE_HOJA_RESUMEN As String = "Resumen Ofertas SSCC"
    Const FILA_INICIO_ORIGEN As Long = 2

    Dim wbDestino As Workbook
    Dim wbOrigen As Workbook

    Dim wsOrigen As Worksheet
    Dim wsResumen As Worksheet

    Dim rutaArchivo As String

    Dim ultimaFilaOrigen As Long
    Dim arrDatos As Variant

    Dim dicGruposServicios As Object
    Dim dicDatosGrupo As Object
    Dim dicServiciosGlobales As Object

    Dim dicServiciosGrupo As Object
    Dim dicHorasServicio As Object

    Dim claveGrupo As String
    Dim claveActual As Variant
    Dim servicioActual As Variant

    Dim listaGrupos As Variant
    Dim listaServicios As Variant

    Dim datosGrupo As Variant

    Dim nombre As String
    Dim servicio As String

    Dim anio As Variant
    Dim mes As Variant
    Dim dia As Variant
    Dim indicador As Variant

    Dim periodo As Long

    Dim arrEncabezados() As Variant
    Dim arrSalida() As Variant

    Dim cantidadGrupos As Long
    Dim cantidadServicios As Long
    Dim cantidadColumnas As Long

    Dim columnaOfertaCompleta As Long
    Dim columnaServicio As Long

    Dim fila As Long
    Dim i As Long

    Dim servicioCompleto As Boolean
    Dim todosServiciosCompletos As Boolean

    Dim tabla As ListObject
    Dim rangoTabla As Range

    Dim filasGeneradasMedidores As Long

    Dim modoCalculo As XlCalculation
    Dim configuracionModificada As Boolean
    Dim archivoOrigenAbierto As Boolean

    On Error GoTo ErrHandler

    '=====================================================
    ' ARCHIVO PRINCIPAL
    '=====================================================
    Set wbDestino = ThisWorkbook

    If Len(wbDestino.Path) = 0 Then

        MsgBox _
            "Primero debes guardar el archivo que contiene la macro.", _
            vbExclamation

        Exit Sub

    End If

    '=====================================================
    ' BUSCAR ARCHIVO QUE CONTENGA OfertasSSCC
    '=====================================================
    rutaArchivo = OSSCC_BuscarArchivoOfertas( _
        wbDestino.Path, _
        wbDestino.FullName _
    )

    If Len(rutaArchivo) = 0 Then

        MsgBox _
            "No se encontró ningún archivo de Excel cuyo nombre " & _
            "contenga el texto:" & vbCrLf & vbCrLf & _
            "OfertasSSCC" & vbCrLf & vbCrLf & _
            "Carpeta revisada:" & vbCrLf & _
            wbDestino.Path, _
            vbCritical

        Exit Sub

    End If

    '=====================================================
    ' OPTIMIZAR EXCEL
    '=====================================================
    With Application

        .ScreenUpdating = False
        .EnableEvents = False
        .DisplayAlerts = False

        modoCalculo = .Calculation
        .Calculation = xlCalculationManual

        .StatusBar = "Abriendo archivo de Ofertas SSCC..."

    End With

    configuracionModificada = True

    '=====================================================
    ' CREAR DICCIONARIOS
    '=====================================================
    Set dicGruposServicios = _
        CreateObject("Scripting.Dictionary")

    Set dicDatosGrupo = _
        CreateObject("Scripting.Dictionary")

    Set dicServiciosGlobales = _
        CreateObject("Scripting.Dictionary")

    dicGruposServicios.CompareMode = vbTextCompare
    dicDatosGrupo.CompareMode = vbTextCompare
    dicServiciosGlobales.CompareMode = vbTextCompare

    '=====================================================
    ' ABRIR ARCHIVO DE ORIGEN
    '=====================================================
    Set wbOrigen = Workbooks.Open( _
        Filename:=rutaArchivo, _
        UpdateLinks:=False, _
        ReadOnly:=True, _
        AddToMru:=False, _
        IgnoreReadOnlyRecommended:=True _
    )

    archivoOrigenAbierto = True

    '=====================================================
    ' ANALIZAR TODAS LAS HOJAS
    '=====================================================
    For Each wsOrigen In wbOrigen.Worksheets

        Application.StatusBar = _
            "Analizando hoja: " & wsOrigen.Name & "..."

        ultimaFilaOrigen = _
            OSSCC_UltimaFilaOrigen(wsOrigen)

        If ultimaFilaOrigen >= FILA_INICIO_ORIGEN Then

            ' Leer A:I en memoria
            arrDatos = wsOrigen.Range( _
                "A1:I" & ultimaFilaOrigen _
            ).Value2

            For fila = FILA_INICIO_ORIGEN To _
                        UBound(arrDatos, 1)

                '-----------------------------------------
                ' COLUMNAS
                '
                ' A = Nombre
                ' B = Año
                ' C = Mes
                ' D = Día
                ' E = Periodo
                ' H = Servicio
                ' I = Ofertó
                '-----------------------------------------
                nombre = _
                    OSSCC_TextoSeguro(arrDatos(fila, 1))

                anio = arrDatos(fila, 2)
                mes = arrDatos(fila, 3)
                dia = arrDatos(fila, 4)

                servicio = _
                    OSSCC_TextoSeguro(arrDatos(fila, 8))

                indicador = arrDatos(fila, 9)

                '-----------------------------------------
                ' FILTRO:
                ' A contiene BESS, SAE o BAT
                '-----------------------------------------
                If OSSCC_ContieneBESSoSAE(nombre) Then

                    '-------------------------------------
                    ' SERVICIO TERMINA EN _RS
                    '-------------------------------------
                    If OSSCC_ServicioTerminaEnRS( _
                        servicio _
                    ) Then

                        If OSSCC_TieneValor(anio) And _
                           OSSCC_TieneValor(mes) And _
                           OSSCC_TieneValor(dia) Then

                            claveGrupo = _
                                OSSCC_CrearClaveGrupo( _
                                    nombre, _
                                    anio, _
                                    mes, _
                                    dia _
                                )

                            '---------------------------------
                            ' CREAR GRUPO:
                            ' Nombre + Año + Mes + Día
                            '---------------------------------
                            If Not dicGruposServicios.Exists( _
                                claveGrupo _
                            ) Then

                                Set dicServiciosGrupo = _
                                    CreateObject( _
                                        "Scripting.Dictionary" _
                                    )

                                dicServiciosGrupo.CompareMode = _
                                    vbTextCompare

                                dicGruposServicios.Add _
                                    claveGrupo, _
                                    dicServiciosGrupo

                                dicDatosGrupo.Add _
                                    claveGrupo, _
                                    Array( _
                                        nombre, _
                                        anio, _
                                        mes, _
                                        dia _
                                    )

                            Else

                                Set dicServiciosGrupo = _
                                    dicGruposServicios.Item( _
                                        claveGrupo _
                                    )

                            End If

                            '---------------------------------
                            ' REGISTRAR SERVICIO GLOBAL
                            '---------------------------------
                            If Not dicServiciosGlobales.Exists( _
                                servicio _
                            ) Then

                                dicServiciosGlobales.Add _
                                    servicio, _
                                    True

                            End If

                            '---------------------------------
                            ' CREAR SERVICIO EN EL GRUPO
                            '---------------------------------
                            If Not dicServiciosGrupo.Exists( _
                                servicio _
                            ) Then

                                Set dicHorasServicio = _
                                    CreateObject( _
                                        "Scripting.Dictionary" _
                                    )

                                dicHorasServicio.CompareMode = _
                                    vbTextCompare

                                dicServiciosGrupo.Add _
                                    servicio, _
                                    dicHorasServicio

                            Else

                                Set dicHorasServicio = _
                                    dicServiciosGrupo.Item( _
                                        servicio _
                                    )

                            End If

                            '---------------------------------
                            ' CONTAR PERIODO SOLO SI I = SÍ
                            '---------------------------------
                            If OSSCC_EsRespuestaSi( _
                                indicador _
                            ) Then

                                periodo = _
                                    OSSCC_NormalizarPeriodo( _
                                        arrDatos(fila, 5) _
                                    )

                                If periodo >= 1 And _
                                   periodo <= 24 Then

                                    ' Horas repetidas cuentan
                                    ' solamente una vez
                                    dicHorasServicio( _
                                        CStr(periodo) _
                                    ) = True

                                End If

                            End If

                        End If

                    End If

                End If

            Next fila

        End If

    Next wsOrigen

    '=====================================================
    ' CERRAR ARCHIVO DE ORIGEN
    '=====================================================
    wbOrigen.Close SaveChanges:=False
    archivoOrigenAbierto = False

    Set wsOrigen = Nothing
    Set wbOrigen = Nothing

    '=====================================================
    ' PREPARAR RESULTADO
    '=====================================================
    cantidadGrupos = dicGruposServicios.Count
    cantidadServicios = dicServiciosGlobales.Count

    If cantidadGrupos = 0 Then

        Err.Raise _
            vbObjectError + 2000, _
            "Generar_Resumen_Ofertas_SSCC", _
            "No se encontraron registros que cumplan " & _
            "las condiciones BESS/SAE/BAT y servicio _RS."

    End If

    listaGrupos = _
        dicGruposServicios.Keys

    listaServicios = _
        OSSCC_ObtenerClavesOrdenadas( _
            dicServiciosGlobales _
        )

    ' Nombre, Año, Mes, Día
    ' + servicios
    ' + Oferta completa
    cantidadColumnas = 4 + cantidadServicios + 1

    columnaOfertaCompleta = cantidadColumnas

    '=====================================================
    ' CREAR O LIMPIAR HOJA RESUMEN
    '=====================================================
    Set wsResumen = OSSCC_ObtenerOCrearHoja( _
        wbDestino, _
        NOMBRE_HOJA_RESUMEN _
    )

    Do While wsResumen.ListObjects.Count > 0
        wsResumen.ListObjects(1).Unlist
    Loop

    wsResumen.Cells.Clear

    '=====================================================
    ' ENCABEZADOS
    '=====================================================
    ReDim arrEncabezados( _
        1 To 1, _
        1 To cantidadColumnas _
    )

    arrEncabezados(1, 1) = "Nombre"
    arrEncabezados(1, 2) = "Año"
    arrEncabezados(1, 3) = "Mes"
    arrEncabezados(1, 4) = "Día"

    For i = 0 To cantidadServicios - 1

        arrEncabezados(1, 5 + i) = _
            CStr(listaServicios(i))

    Next i

    arrEncabezados( _
        1, _
        columnaOfertaCompleta _
    ) = "Oferta completa"

    wsResumen.Cells(1, 1).Resize( _
        1, _
        cantidadColumnas _
    ).Value2 = arrEncabezados

    '=====================================================
    ' GENERAR FILAS DEL RESUMEN
    '=====================================================
    ReDim arrSalida( _
        1 To cantidadGrupos, _
        1 To cantidadColumnas _
    )

    For i = 0 To cantidadGrupos - 1

        claveActual = listaGrupos(i)

        datosGrupo = dicDatosGrupo.Item( _
            CStr(claveActual) _
        )

        Set dicServiciosGrupo = _
            dicGruposServicios.Item( _
                CStr(claveActual) _
            )

        arrSalida(i + 1, 1) = datosGrupo(0)
        arrSalida(i + 1, 2) = datosGrupo(1)
        arrSalida(i + 1, 3) = datosGrupo(2)
        arrSalida(i + 1, 4) = datosGrupo(3)

        todosServiciosCompletos = True

        '---------------------------------------------
        ' UNA COLUMNA POR CADA SERVICIO _RS
        '---------------------------------------------
        For columnaServicio = 0 To _
                              cantidadServicios - 1

            servicioActual = _
                listaServicios(columnaServicio)

            servicioCompleto = False

            If dicServiciosGrupo.Exists( _
                CStr(servicioActual) _
            ) Then

                Set dicHorasServicio = _
                    dicServiciosGrupo.Item( _
                        CStr(servicioActual) _
                    )

                servicioCompleto = _
                    (dicHorasServicio.Count = 24)

            End If

            If servicioCompleto Then

                arrSalida( _
                    i + 1, _
                    5 + columnaServicio _
                ) = 1

            Else

                ' Si el servicio no aparece o no tiene
                ' las 24 horas, queda en cero
                arrSalida( _
                    i + 1, _
                    5 + columnaServicio _
                ) = 0

                todosServiciosCompletos = False

            End If

        Next columnaServicio

        '---------------------------------------------
        ' OFERTA COMPLETA
        '
        ' Todos los servicios _RS encontrados en el
        ' archivo deben estar completos las 24 horas.
        '---------------------------------------------
        If todosServiciosCompletos And _
           cantidadServicios > 0 Then

            arrSalida( _
                i + 1, _
                columnaOfertaCompleta _
            ) = 1

        Else

            arrSalida( _
                i + 1, _
                columnaOfertaCompleta _
            ) = 0

        End If

    Next i

    wsResumen.Cells(2, 1).Resize( _
        cantidadGrupos, _
        cantidadColumnas _
    ).Value2 = arrSalida

    '=====================================================
    ' ORDENAR RESUMEN
    '=====================================================
    With wsResumen.Sort

        .SortFields.Clear

        .SortFields.Add _
            key:=wsResumen.Range( _
                "A2:A" & cantidadGrupos + 1 _
            ), _
            SortOn:=xlSortOnValues, _
            Order:=xlAscending, _
            DataOption:=xlSortNormal

        .SortFields.Add _
            key:=wsResumen.Range( _
                "B2:B" & cantidadGrupos + 1 _
            ), _
            SortOn:=xlSortOnValues, _
            Order:=xlAscending, _
            DataOption:=xlSortNormal

        .SortFields.Add _
            key:=wsResumen.Range( _
                "C2:C" & cantidadGrupos + 1 _
            ), _
            SortOn:=xlSortOnValues, _
            Order:=xlAscending, _
            DataOption:=xlSortNormal

        .SortFields.Add _
            key:=wsResumen.Range( _
                "D2:D" & cantidadGrupos + 1 _
            ), _
            SortOn:=xlSortOnValues, _
            Order:=xlAscending, _
            DataOption:=xlSortNormal

        .SetRange wsResumen.Range( _
            wsResumen.Cells(1, 1), _
            wsResumen.Cells( _
                cantidadGrupos + 1, _
                cantidadColumnas _
            ) _
        )

        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply

    End With

    '=====================================================
    ' CONVERTIR EN TABLA
    '=====================================================
    Set rangoTabla = wsResumen.Range( _
        wsResumen.Cells(1, 1), _
        wsResumen.Cells( _
            cantidadGrupos + 1, _
            cantidadColumnas _
        ) _
    )

    Set tabla = wsResumen.ListObjects.Add( _
        SourceType:=xlSrcRange, _
        Source:=rangoTabla, _
        XlListObjectHasHeaders:=xlYes _
    )

    On Error Resume Next
    tabla.Name = "tblResumenOfertasSSCC"
    On Error GoTo ErrHandler

    '=====================================================
    ' FORMATO DEL RESUMEN
    '=====================================================
    With wsResumen

        .Rows(1).Font.Bold = True
        .Columns.AutoFit
        .Columns("B:D").NumberFormat = "0"

        .Range( _
            .Cells(2, 5), _
            .Cells( _
                cantidadGrupos + 1, _
                cantidadColumnas _
            ) _
        ).NumberFormat = "0"

    End With

    '=====================================================
    ' TRASPASAR A MEDIDORES W:X:Y
    '=====================================================
    Application.StatusBar = _
        "Completando días y cargando Medidores W:Y..."

    filasGeneradasMedidores = _
        OSSCC_CargarResumenEnMedidores( _
            wbDestino, _
            wsResumen _
        )

    Application.StatusBar = False

    MsgBox _
        "Proceso terminado correctamente." & vbCrLf & vbCrLf & _
        "Archivo analizado:" & vbCrLf & _
        Dir(rutaArchivo) & vbCrLf & vbCrLf & _
        "Registros en Resumen Ofertas SSCC: " & _
        cantidadGrupos & vbCrLf & _
        "Servicios _RS encontrados: " & _
        cantidadServicios & vbCrLf & vbCrLf & _
        "Filas creadas en Medidores W:Y: " & _
        filasGeneradasMedidores & vbCrLf & vbCrLf & _
        "Los días faltantes y los nombres de Medidores!G " & _
        "que no estaban representados directa ni mediante " & _
        "Diccionario!E:F:G fueron completados con oferta 0.", _
        vbInformation

Salida:

    Application.StatusBar = False

    If configuracionModificada Then

        With Application
            .ScreenUpdating = True
            .EnableEvents = True
            .DisplayAlerts = True
            .Calculation = modoCalculo
        End With

    End If

    Exit Sub

ErrHandler:

    On Error Resume Next

    If archivoOrigenAbierto Then
        wbOrigen.Close SaveChanges:=False
    End If

    On Error GoTo 0

    MsgBox _
        "Error al generar el resumen de Ofertas SSCC:" & _
        vbCrLf & vbCrLf & _
        Err.Description, _
        vbCritical

    Resume Salida

End Sub


Private Function OSSCC_CargarResumenEnMedidores( _
    ByVal wb As Workbook, _
    ByVal wsResumen As Worksheet _
) As Long

    Const FILA_INICIO_RESUMEN As Long = 2
    Const FILA_INICIO_MEDIDORES As Long = 3
    Const FILA_INICIO_DICCIONARIO As Long = 2

    Const COLUMNA_NOMBRE_MEDIDORES As Long = 7   ' G
    Const COLUMNA_DIC_1 As Long = 5              ' E
    Const COLUMNA_DIC_2 As Long = 6              ' F
    Const COLUMNA_DIC_3 As Long = 7              ' G
    Const COLUMNA_DESTINO_NOMBRE As Long = 23    ' W
    Const COLUMNA_DESTINO_DIA As Long = 24       ' X
    Const COLUMNA_DESTINO_OFERTA As Long = 25    ' Y

    Const MAX_NOMBRES_AVISO As Long = 30

    Dim wsMedidores As Worksheet
    Dim wsDiccionario As Worksheet

    Dim columnaOfertaCompleta As Long
    Dim ultimaFilaResumen As Long
    Dim ultimaFilaG As Long
    Dim ultimaFilaDiccionario As Long
    Dim ultimaFilaDestinoAnterior As Long

    Dim arrResumen As Variant
    Dim arrNombresG As Variant
    Dim arrDiccionario As Variant
    Dim arrSalida() As Variant
    Dim arrAlias(1 To 3) As String

    Dim dicNombres As Object
    Dim dicOfertas As Object
    Dim dicPeriodos As Object
    Dim dicEquivalencias As Object
    Dim dicNoEncontradosDiccionario As Object

    Dim clavesNombres As Variant
    Dim clavesNoEncontrados As Variant
    Dim claveNombre As String
    Dim claveOferta As String
    Dim clavePeriodo As String
    Dim claveAlias As String

    Dim nombre As String
    Dim nombreMostrar As String
    Dim cadenaEquivalencias As String
    Dim mensajeNoEncontrados As String

    Dim anio As Long
    Dim mes As Long
    Dim dia As Long
    Dim diasDelMes As Long

    Dim valorOferta As Long

    Dim periodoSeleccionado As Variant
    Dim partesPeriodo As Variant
    Dim partesEquivalentes As Variant
    Dim aliasActual As Variant

    Dim fila As Long
    Dim i As Long
    Dim d As Long
    Dim k As Long
    Dim filaSalida As Long
    Dim cantidadNombres As Long
    Dim cantidadFilasSalida As Long
    Dim cantidadMostrar As Long

    Dim nombreYaRepresentado As Boolean
    Dim encontradoEnDiccionario As Boolean
    Dim separador As String
    Dim separadorAlias As String

    Set wsMedidores = wb.Worksheets("Medidores")

    On Error Resume Next
    Set wsDiccionario = wb.Worksheets("Diccionario")
    On Error GoTo 0

    If wsDiccionario Is Nothing Then

        Err.Raise _
            vbObjectError + 2105, _
            "OSSCC_CargarResumenEnMedidores", _
            "No se encontró la hoja 'Diccionario'." & _
            vbCrLf & vbCrLf & _
            "Las columnas E, F y G deben contener nombres " & _
            "equivalentes usados en Medidores/Subastas/Ofertas."

    End If

    separador = Chr$(30)
    separadorAlias = Chr$(29)

    '=====================================================
    ' BUSCAR COLUMNA "OFERTA COMPLETA"
    '=====================================================
    columnaOfertaCompleta = _
        OSSCC_BuscarColumnaEncabezado( _
            wsResumen, _
            "Oferta completa" _
        )

    If columnaOfertaCompleta = 0 Then

        Err.Raise _
            vbObjectError + 2100, _
            "OSSCC_CargarResumenEnMedidores", _
            "No se encontró el encabezado 'Oferta completa' " & _
            "en la hoja Resumen Ofertas SSCC."

    End If

    '=====================================================
    ' ÚLTIMA FILA DEL RESUMEN
    '=====================================================
    ultimaFilaResumen = _
        OSSCC_UltimaFilaValoresEnColumnas( _
            wsResumen, _
            Array( _
                1, _
                2, _
                3, _
                4, _
                columnaOfertaCompleta _
            ) _
        )

    If ultimaFilaResumen < FILA_INICIO_RESUMEN Then

        Err.Raise _
            vbObjectError + 2101, _
            "OSSCC_CargarResumenEnMedidores", _
            "La hoja Resumen Ofertas SSCC no contiene datos."

    End If

    '=====================================================
    ' LEER RESUMEN EN MEMORIA
    '=====================================================
    arrResumen = wsResumen.Range( _
        wsResumen.Cells(FILA_INICIO_RESUMEN, 1), _
        wsResumen.Cells( _
            ultimaFilaResumen, _
            columnaOfertaCompleta _
        ) _
    ).Value2

    Set dicNombres = _
        CreateObject("Scripting.Dictionary")

    Set dicOfertas = _
        CreateObject("Scripting.Dictionary")

    Set dicPeriodos = _
        CreateObject("Scripting.Dictionary")

    Set dicEquivalencias = _
        CreateObject("Scripting.Dictionary")

    Set dicNoEncontradosDiccionario = _
        CreateObject("Scripting.Dictionary")

    dicNombres.CompareMode = vbTextCompare
    dicOfertas.CompareMode = vbTextCompare
    dicPeriodos.CompareMode = vbTextCompare
    dicEquivalencias.CompareMode = vbTextCompare
    dicNoEncontradosDiccionario.CompareMode = vbTextCompare

    '=====================================================
    ' LEER EQUIVALENCIAS DE DICCIONARIO!E:G
    '
    ' Las columnas E, F y G se consideran equivalentes
    ' dentro de una misma fila.
    '
    ' Ejemplo:
    ' E = BESS ABC
    ' F = SAE ABC
    ' G = BESS_ABC_RS
    '
    ' Si cualquiera de los tres nombres aparece en Medidores
    ' o en el resumen, se revisan también los otros dos.
    '=====================================================
    ultimaFilaDiccionario = _
        OSSCC_UltimaFilaValoresEnColumnas( _
            wsDiccionario, _
            Array( _
                COLUMNA_DIC_1, _
                COLUMNA_DIC_2, _
                COLUMNA_DIC_3 _
            ) _
        )

    If ultimaFilaDiccionario >= _
       FILA_INICIO_DICCIONARIO Then

        arrDiccionario = wsDiccionario.Range( _
            wsDiccionario.Cells( _
                FILA_INICIO_DICCIONARIO, _
                COLUMNA_DIC_1 _
            ), _
            wsDiccionario.Cells( _
                ultimaFilaDiccionario, _
                COLUMNA_DIC_3 _
            ) _
        ).Value2

        For fila = 1 To UBound(arrDiccionario, 1)

            arrAlias(1) = _
                OSSCC_NormalizarNombreClave( _
                    arrDiccionario(fila, 1) _
                )

            arrAlias(2) = _
                OSSCC_NormalizarNombreClave( _
                    arrDiccionario(fila, 2) _
                )

            arrAlias(3) = _
                OSSCC_NormalizarNombreClave( _
                    arrDiccionario(fila, 3) _
                )

            ' Construir una única cadena con todos los nombres
            ' no vacíos de la fila, evitando duplicados.
            cadenaEquivalencias = ""

            For k = 1 To 3

                claveAlias = arrAlias(k)

                If Len(claveAlias) > 0 Then

                    If InStr( _
                        1, _
                        separadorAlias & cadenaEquivalencias & separadorAlias, _
                        separadorAlias & claveAlias & separadorAlias, _
                        vbTextCompare _
                    ) = 0 Then

                        If Len(cadenaEquivalencias) > 0 Then
                            cadenaEquivalencias = _
                                cadenaEquivalencias & separadorAlias
                        End If

                        cadenaEquivalencias = _
                            cadenaEquivalencias & claveAlias

                    End If

                End If

            Next k

            If Len(cadenaEquivalencias) > 0 Then

                ' Cada nombre de E/F/G permite recuperar todos
                ' los equivalentes existentes en esa misma fila.
                For k = 1 To 3

                    claveAlias = arrAlias(k)

                    If Len(claveAlias) > 0 Then
                        dicEquivalencias(claveAlias) = _
                            cadenaEquivalencias
                    End If

                Next k

            End If

        Next fila

    End If

    '=====================================================
    ' LEER NOMBRES, DÍAS Y OFERTA DEL RESUMEN
    '=====================================================
    For fila = 1 To UBound(arrResumen, 1)

        nombreMostrar = _
            OSSCC_LimpiarNombreMostrar( _
                arrResumen(fila, 1) _
            )

        If Len(nombreMostrar) > 0 Then

            If IsNumeric(arrResumen(fila, 2)) And _
               IsNumeric(arrResumen(fila, 3)) And _
               IsNumeric(arrResumen(fila, 4)) Then

                anio = CLng(arrResumen(fila, 2))
                mes = CLng(arrResumen(fila, 3))
                dia = CLng(arrResumen(fila, 4))

                If mes >= 1 And mes <= 12 Then

                    clavePeriodo = _
                        CStr(anio) & separador & CStr(mes)

                    If Not dicPeriodos.Exists( _
                        clavePeriodo _
                    ) Then

                        dicPeriodos.Add _
                            clavePeriodo, _
                            True

                    End If

                    claveNombre = _
                        OSSCC_NormalizarNombreClave( _
                            nombreMostrar _
                        )

                    ' Los nombres provenientes del resumen
                    ' son los que deben conservarse en W.
                    If Not dicNombres.Exists( _
                        claveNombre _
                    ) Then

                        dicNombres.Add _
                            claveNombre, _
                            nombreMostrar

                    End If

                    claveOferta = _
                        claveNombre & _
                        separador & _
                        CStr(dia)

                    valorOferta = _
                        OSSCC_ValorOfertaBinario( _
                            arrResumen( _
                                fila, _
                                columnaOfertaCompleta _
                            ) _
                        )

                    dicOfertas(claveOferta) = valorOferta

                End If

            End If

        End If

    Next fila

    '=====================================================
    ' VALIDAR UN SOLO AÑO Y MES
    '
    ' W:Y solamente guarda Nombre, Día y Oferta.
    ' Si hubiera más de un mes, no podrían distinguirse.
    '=====================================================
    If dicPeriodos.Count = 0 Then

        Err.Raise _
            vbObjectError + 2102, _
            "OSSCC_CargarResumenEnMedidores", _
            "No fue posible determinar el año y mes " & _
            "del resumen."

    End If

    If dicPeriodos.Count > 1 Then

        Err.Raise _
            vbObjectError + 2103, _
            "OSSCC_CargarResumenEnMedidores", _
            "El resumen contiene más de un año o mes." & _
            vbCrLf & vbCrLf & _
            "No es posible pegar varios meses en W:Y porque " & _
            "el destino solamente contiene Nombre, Día y Oferta."

    End If

    periodoSeleccionado = dicPeriodos.Keys()(0)
    partesPeriodo = Split( _
        CStr(periodoSeleccionado), _
        separador _
    )

    anio = CLng(partesPeriodo(0))
    mes = CLng(partesPeriodo(1))

    diasDelMes = Day( _
        DateSerial(anio, mes + 1, 0) _
    )

    '=====================================================
    ' LEER NOMBRES DE MEDIDORES!G3:G
    '
    ' Antes de agregar un nombre con oferta cero:
    ' 1. Revisar si el mismo nombre ya está en el resumen.
    ' 2. Si no está directo, buscar Medidores!G en cualquiera
    '    de las columnas Diccionario!E, F o G.
    ' 3. Si lo encuentra, revisar TODOS los nombres equivalentes
    '    de esa fila E:G contra los nombres del resumen.
    ' 4. Si el nombre no existe en Diccionario!E:G, registrarlo
    '    para mostrar una advertencia al finalizar.
    ' 5. Si no está representado de ninguna forma, agregarlo
    '    igualmente y completar todos sus días con oferta 0.
    '=====================================================
    ultimaFilaG = _
        OSSCC_UltimaFilaValoresColumna( _
            wsMedidores, _
            COLUMNA_NOMBRE_MEDIDORES _
        )

    If ultimaFilaG >= FILA_INICIO_MEDIDORES Then

        arrNombresG = wsMedidores.Range( _
            wsMedidores.Cells( _
                FILA_INICIO_MEDIDORES, _
                COLUMNA_NOMBRE_MEDIDORES _
            ), _
            wsMedidores.Cells( _
                ultimaFilaG, _
                COLUMNA_NOMBRE_MEDIDORES _
            ) _
        ).Value2

        For fila = 1 To UBound(arrNombresG, 1)

            nombreMostrar = _
                OSSCC_LimpiarNombreMostrar( _
                    arrNombresG(fila, 1) _
                )

            If Len(nombreMostrar) > 0 Then

                claveNombre = _
                    OSSCC_NormalizarNombreClave( _
                        nombreMostrar _
                    )

                nombreYaRepresentado = _
                    dicNombres.Exists(claveNombre)

                encontradoEnDiccionario = _
                    dicEquivalencias.Exists(claveNombre)

                ' Si el nombre no está directo en el resumen,
                ' revisar todos sus equivalentes de E:F:G.
                If Not nombreYaRepresentado Then

                    If encontradoEnDiccionario Then

                        partesEquivalentes = Split( _
                            CStr(dicEquivalencias.Item( _
                                claveNombre _
                            )), _
                            separadorAlias _
                        )

                        For Each aliasActual In partesEquivalentes

                            If dicNombres.Exists( _
                                CStr(aliasActual) _
                            ) Then

                                nombreYaRepresentado = True
                                Exit For

                            End If

                        Next aliasActual

                    Else

                        ' El nombre de Medidores!G no existe en
                        ' ninguna de las columnas E, F o G.
                        If Not dicNoEncontradosDiccionario.Exists( _
                            claveNombre _
                        ) Then

                            dicNoEncontradosDiccionario.Add _
                                claveNombre, _
                                nombreMostrar

                        End If

                    End If

                End If

                ' Solo se agrega con ofertas cero si ni el
                ' nombre directo ni ninguno de sus equivalentes
                ' E/F/G está representado en el resumen.
                If Not nombreYaRepresentado Then

                    If Not dicNombres.Exists(claveNombre) Then
                        dicNombres.Add _
                            claveNombre, _
                            nombreMostrar
                    End If

                End If

            End If

        Next fila

    End If

    cantidadNombres = dicNombres.Count

    If cantidadNombres = 0 Then

        Err.Raise _
            vbObjectError + 2104, _
            "OSSCC_CargarResumenEnMedidores", _
            "No se encontraron nombres en el resumen ni " & _
            "en Medidores!G3:G."

    End If

    '=====================================================
    ' ORDENAR NOMBRES
    '=====================================================
    clavesNombres = _
        OSSCC_ObtenerClavesOrdenadasPorValor( _
            dicNombres _
        )

    cantidadFilasSalida = _
        cantidadNombres * diasDelMes

    ReDim arrSalida( _
        1 To cantidadFilasSalida, _
        1 To 3 _
    )

    '=====================================================
    ' CREAR TODOS LOS DÍAS PARA CADA NOMBRE
    '=====================================================
    filaSalida = 0

    For i = 0 To cantidadNombres - 1

        claveNombre = CStr(clavesNombres(i))
        nombre = CStr(dicNombres.Item(claveNombre))

        For d = 1 To diasDelMes

            filaSalida = filaSalida + 1

            arrSalida(filaSalida, 1) = nombre
            arrSalida(filaSalida, 2) = d

            claveOferta = _
                claveNombre & _
                separador & _
                CStr(d)

            If dicOfertas.Exists(claveOferta) Then

                arrSalida(filaSalida, 3) = _
                    dicOfertas.Item(claveOferta)

            Else

                ' Día faltante o nombre de Medidores que no
                ' estaba representado directa ni mediante
                ' equivalencia E/F/G en el resumen.
                arrSalida(filaSalida, 3) = 0

            End If

        Next d

    Next i

    '=====================================================
    ' BORRAR RESULTADO ANTERIOR W3:Y
    '=====================================================
    ultimaFilaDestinoAnterior = _
        OSSCC_UltimaFilaValoresEnColumnas( _
            wsMedidores, _
            Array( _
                COLUMNA_DESTINO_NOMBRE, _
                COLUMNA_DESTINO_DIA, _
                COLUMNA_DESTINO_OFERTA _
            ) _
        )

    If ultimaFilaDestinoAnterior >= _
       FILA_INICIO_MEDIDORES Then

        wsMedidores.Range( _
            wsMedidores.Cells( _
                FILA_INICIO_MEDIDORES, _
                COLUMNA_DESTINO_NOMBRE _
            ), _
            wsMedidores.Cells( _
                ultimaFilaDestinoAnterior, _
                COLUMNA_DESTINO_OFERTA _
            ) _
        ).ClearContents

    End If

    '=====================================================
    ' PEGAR EN W3:Y
    '=====================================================
    wsMedidores.Cells( _
        FILA_INICIO_MEDIDORES, _
        COLUMNA_DESTINO_NOMBRE _
    ).Resize( _
        cantidadFilasSalida, _
        3 _
    ).Value2 = arrSalida

    '=====================================================
    ' AVISAR NOMBRES NO ENCONTRADOS EN DICCIONARIO E:F:G
    '=====================================================
    If dicNoEncontradosDiccionario.Count > 0 Then

        clavesNoEncontrados = _
            OSSCC_ObtenerClavesOrdenadasPorValor( _
                dicNoEncontradosDiccionario _
            )

        cantidadMostrar = dicNoEncontradosDiccionario.Count
        If cantidadMostrar > MAX_NOMBRES_AVISO Then
            cantidadMostrar = MAX_NOMBRES_AVISO
        End If

        mensajeNoEncontrados = _
            "ADVERTENCIA:" & vbCrLf & vbCrLf & _
            CStr(dicNoEncontradosDiccionario.Count) & _
            " nombre(s) de Medidores!G no fueron encontrados " & _
            "en ninguna de las columnas Diccionario!E:F:G." & _
            vbCrLf & vbCrLf & _
            "Se incorporaron igualmente con oferta 0:" & _
            vbCrLf & vbCrLf

        For i = 0 To cantidadMostrar - 1

            mensajeNoEncontrados = _
                mensajeNoEncontrados & "- " & _
                CStr(dicNoEncontradosDiccionario.Item( _
                    clavesNoEncontrados(i) _
                )) & vbCrLf

        Next i

        If dicNoEncontradosDiccionario.Count > _
           MAX_NOMBRES_AVISO Then

            mensajeNoEncontrados = _
                mensajeNoEncontrados & vbCrLf & _
                "... y " & _
                CStr( _
                    dicNoEncontradosDiccionario.Count - _
                    MAX_NOMBRES_AVISO _
                ) & _
                " nombre(s) adicional(es)."

        End If

        MsgBox _
            mensajeNoEncontrados, _
            vbExclamation, _
            "Nombres no encontrados en Diccionario"

    End If

    OSSCC_CargarResumenEnMedidores = _
        cantidadFilasSalida

End Function



Private Function OSSCC_ContieneBESSoSAE( _
    ByVal texto As String _
) As Boolean

    OSSCC_ContieneBESSoSAE = _
        InStr(1, texto, "BESS", vbTextCompare) > 0 Or _
        InStr(1, texto, "SAE", vbTextCompare) > 0 Or _
        InStr(1, texto, "BAT", vbTextCompare) > 0

End Function


Private Function OSSCC_ServicioTerminaEnRS( _
    ByVal servicio As String _
) As Boolean

    servicio = Trim$(servicio)

    If Len(servicio) < 3 Then

        OSSCC_ServicioTerminaEnRS = False

    Else

        OSSCC_ServicioTerminaEnRS = _
            StrComp( _
                Right$(servicio, 3), _
                "_RS", _
                vbTextCompare _
            ) = 0

    End If

End Function


Private Function OSSCC_EsRespuestaSi( _
    ByVal valor As Variant _
) As Boolean

    Dim texto As String

    If IsError(valor) Or IsEmpty(valor) Then Exit Function

    texto = UCase$(Trim$(CStr(valor)))

    texto = Replace(texto, Chr$(160), "")
    texto = Replace(texto, " ", "")
    texto = Replace(texto, vbTab, "")

    texto = Replace(texto, "Í", "I")
    texto = Replace(texto, "Ì", "I")
    texto = Replace(texto, "Ï", "I")
    texto = Replace(texto, "Î", "I")

    texto = Replace(texto, ".", "")
    texto = Replace(texto, ",", "")
    texto = Replace(texto, ";", "")
    texto = Replace(texto, ":", "")

    OSSCC_EsRespuestaSi = (texto = "SI")

End Function


Private Function OSSCC_NormalizarPeriodo( _
    ByVal valor As Variant _
) As Long

    Dim numero As Double
    Dim texto As String

    OSSCC_NormalizarPeriodo = -1

    If IsError(valor) Or IsEmpty(valor) Then Exit Function

    ' En el archivo mostrado, PERIOD usa valores 1 a 24.
    If IsNumeric(valor) Then

        numero = CDbl(valor)

        If numero = Fix(numero) Then

            If numero >= 1 And numero <= 24 Then

                OSSCC_NormalizarPeriodo = CLng(numero)
                Exit Function

            End If

        End If

        ' Hora Excel guardada como fracción del día
        If numero >= 0 And numero < 1 Then

            OSSCC_NormalizarPeriodo = _
                Hour(CDate(numero)) + 1

            Exit Function

        End If

    End If

    texto = Trim$(CStr(valor))

    If Len(texto) = 0 Then Exit Function

    If IsNumeric(texto) Then

        numero = CDbl(texto)

        If numero = Fix(numero) And _
           numero >= 1 And numero <= 24 Then

            OSSCC_NormalizarPeriodo = CLng(numero)

        End If

        Exit Function

    End If

    If texto = "24:00" Or _
       texto = "24:00:00" Then

        OSSCC_NormalizarPeriodo = 24
        Exit Function

    End If

    If IsDate(texto) Then

        OSSCC_NormalizarPeriodo = _
            Hour(CDate(texto)) + 1

    End If

End Function


Private Function OSSCC_CrearClaveGrupo( _
    ByVal nombre As String, _
    ByVal anio As Variant, _
    ByVal mes As Variant, _
    ByVal dia As Variant _
) As String

    Dim separador As String

    separador = Chr$(30)

    OSSCC_CrearClaveGrupo = _
        Trim$(nombre) & separador & _
        OSSCC_ValorClave(anio) & separador & _
        OSSCC_ValorClave(mes) & separador & _
        OSSCC_ValorClave(dia)

End Function


Private Function OSSCC_ValorClave( _
    ByVal valor As Variant _
) As String

    If IsError(valor) Or IsEmpty(valor) Then

        OSSCC_ValorClave = ""

    ElseIf IsNumeric(valor) Then

        OSSCC_ValorClave = _
            Format$(CDbl(valor), "0.###############")

    Else

        OSSCC_ValorClave = Trim$(CStr(valor))

    End If

End Function


Private Function OSSCC_TextoSeguro( _
    ByVal valor As Variant _
) As String

    If IsError(valor) Or IsEmpty(valor) Then

        OSSCC_TextoSeguro = ""

    Else

        OSSCC_TextoSeguro = Trim$(CStr(valor))

    End If

End Function


Private Function OSSCC_TieneValor( _
    ByVal valor As Variant _
) As Boolean

    If IsError(valor) Or IsEmpty(valor) Then

        OSSCC_TieneValor = False

    Else

        OSSCC_TieneValor = _
            Len(Trim$(CStr(valor))) > 0

    End If

End Function


Private Function OSSCC_ValorOfertaBinario( _
    ByVal valor As Variant _
) As Long

    Dim texto As String

    If IsError(valor) Or IsEmpty(valor) Then

        OSSCC_ValorOfertaBinario = 0
        Exit Function

    End If

    If IsNumeric(valor) Then

        If CDbl(valor) = 1 Then
            OSSCC_ValorOfertaBinario = 1
        Else
            OSSCC_ValorOfertaBinario = 0
        End If

        Exit Function

    End If

    texto = Trim$(CStr(valor))

    If texto = "1" Then
        OSSCC_ValorOfertaBinario = 1
    Else
        OSSCC_ValorOfertaBinario = 0
    End If

End Function


Private Function OSSCC_LimpiarNombreMostrar( _
    ByVal valor As Variant _
) As String

    Dim texto As String

    If IsError(valor) Or IsEmpty(valor) Then Exit Function

    texto = CStr(valor)

    texto = Replace(texto, Chr$(160), " ")
    texto = Trim$(texto)

    Do While InStr(1, texto, "  ", vbBinaryCompare) > 0
        texto = Replace(texto, "  ", " ")
    Loop

    OSSCC_LimpiarNombreMostrar = texto

End Function


Private Function OSSCC_NormalizarNombreClave( _
    ByVal valor As Variant _
) As String

    OSSCC_NormalizarNombreClave = _
        UCase$(OSSCC_LimpiarNombreMostrar(valor))

End Function


Private Function OSSCC_ObtenerClavesOrdenadas( _
    ByVal diccionario As Object _
) As Variant

    Dim claves As Variant
    Dim temporal As Variant

    Dim i As Long
    Dim j As Long

    claves = diccionario.Keys

    If diccionario.Count > 1 Then

        For i = LBound(claves) To UBound(claves) - 1

            For j = i + 1 To UBound(claves)

                If StrComp( _
                    CStr(claves(i)), _
                    CStr(claves(j)), _
                    vbTextCompare _
                ) > 0 Then

                    temporal = claves(i)
                    claves(i) = claves(j)
                    claves(j) = temporal

                End If

            Next j

        Next i

    End If

    OSSCC_ObtenerClavesOrdenadas = claves

End Function


Private Function OSSCC_ObtenerClavesOrdenadasPorValor( _
    ByVal diccionario As Object _
) As Variant

    Dim claves As Variant
    Dim temporal As Variant

    Dim i As Long
    Dim j As Long

    claves = diccionario.Keys

    If diccionario.Count > 1 Then

        For i = LBound(claves) To UBound(claves) - 1

            For j = i + 1 To UBound(claves)

                If StrComp( _
                    CStr(diccionario.Item(claves(i))), _
                    CStr(diccionario.Item(claves(j))), _
                    vbTextCompare _
                ) > 0 Then

                    temporal = claves(i)
                    claves(i) = claves(j)
                    claves(j) = temporal

                End If

            Next j

        Next i

    End If

    OSSCC_ObtenerClavesOrdenadasPorValor = claves

End Function


Private Function OSSCC_BuscarArchivoOfertas( _
    ByVal rutaCarpeta As String, _
    ByVal archivoExcluir As String _
) As String

    Dim nombreArchivo As String
    Dim rutaCompleta As String
    Dim rutaSeleccionada As String

    Dim extensionArchivo As String
    Dim fechaArchivo As Date
    Dim fechaMasReciente As Date

    nombreArchivo = Dir( _
        rutaCarpeta & _
        Application.PathSeparator & _
        "*OfertasSSCC*.*" _
    )

    Do While Len(nombreArchivo) > 0

        If Left$(nombreArchivo, 2) <> "~$" Then

            rutaCompleta = _
                rutaCarpeta & _
                Application.PathSeparator & _
                nombreArchivo

            If StrComp( _
                rutaCompleta, _
                archivoExcluir, _
                vbTextCompare _
            ) <> 0 Then

                If InStrRev(nombreArchivo, ".") > 0 Then

                    extensionArchivo = _
                        LCase$(Mid$( _
                            nombreArchivo, _
                            InStrRev(nombreArchivo, ".") _
                        ))

                    Select Case extensionArchivo

                        Case ".xlsx", ".xlsm", ".xlsb", ".xls"

                            fechaArchivo = _
                                FileDateTime(rutaCompleta)

                            If Len(rutaSeleccionada) = 0 Or _
                               fechaArchivo > fechaMasReciente Then

                                rutaSeleccionada = rutaCompleta
                                fechaMasReciente = fechaArchivo

                            End If

                    End Select

                End If

            End If

        End If

        nombreArchivo = Dir

    Loop

    OSSCC_BuscarArchivoOfertas = rutaSeleccionada

End Function


Private Function OSSCC_UltimaFilaOrigen( _
    ByVal ws As Worksheet _
) As Long

    Dim rangoBusqueda As Range
    Dim celdaEncontrada As Range

    Set rangoBusqueda = ws.Range( _
        ws.Cells(1, 1), _
        ws.Cells(ws.Rows.Count, 9) _
    )

    Set celdaEncontrada = rangoBusqueda.Find( _
        What:="*", _
        After:=rangoBusqueda.Cells(1, 1), _
        LookIn:=xlFormulas, _
        LookAt:=xlPart, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlPrevious, _
        MatchCase:=False _
    )

    If celdaEncontrada Is Nothing Then
        OSSCC_UltimaFilaOrigen = 0
    Else
        OSSCC_UltimaFilaOrigen = celdaEncontrada.Row
    End If

End Function


Private Function OSSCC_UltimaFilaValoresColumna( _
    ByVal ws As Worksheet, _
    ByVal numeroColumna As Long _
) As Long

    Dim celdaEncontrada As Range

    Set celdaEncontrada = ws.Columns( _
        numeroColumna _
    ).Find( _
        What:="*", _
        After:=ws.Cells(1, numeroColumna), _
        LookIn:=xlValues, _
        LookAt:=xlPart, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlPrevious, _
        MatchCase:=False _
    )

    If celdaEncontrada Is Nothing Then
        OSSCC_UltimaFilaValoresColumna = 0
    Else
        OSSCC_UltimaFilaValoresColumna = _
            celdaEncontrada.Row
    End If

End Function


Private Function OSSCC_UltimaFilaValoresEnColumnas( _
    ByVal ws As Worksheet, _
    ByVal columnas As Variant _
) As Long

    Dim columna As Variant
    Dim celdaEncontrada As Range
    Dim maximaFila As Long

    maximaFila = 0

    For Each columna In columnas

        Set celdaEncontrada = ws.Columns( _
            CLng(columna) _
        ).Find( _
            What:="*", _
            After:=ws.Cells(1, CLng(columna)), _
            LookIn:=xlValues, _
            LookAt:=xlPart, _
            SearchOrder:=xlByRows, _
            SearchDirection:=xlPrevious, _
            MatchCase:=False _
        )

        If Not celdaEncontrada Is Nothing Then

            If celdaEncontrada.Row > maximaFila Then
                maximaFila = celdaEncontrada.Row
            End If

        End If

        Set celdaEncontrada = Nothing

    Next columna

    OSSCC_UltimaFilaValoresEnColumnas = maximaFila

End Function


Private Function OSSCC_BuscarColumnaEncabezado( _
    ByVal ws As Worksheet, _
    ByVal encabezadoBuscado As String _
) As Long

    Dim ultimaColumna As Long
    Dim columna As Long
    Dim texto As String

    ultimaColumna = ws.Cells( _
        1, _
        ws.Columns.Count _
    ).End(xlToLeft).Column

    For columna = 1 To ultimaColumna

        texto = Trim$(CStr( _
            ws.Cells(1, columna).Value2 _
        ))

        If StrComp( _
            texto, _
            encabezadoBuscado, _
            vbTextCompare _
        ) = 0 Then

            OSSCC_BuscarColumnaEncabezado = columna
            Exit Function

        End If

    Next columna

    OSSCC_BuscarColumnaEncabezado = 0

End Function

Private Function OSSCC_ObtenerOCrearHoja( _
    ByVal wb As Workbook, _
    ByVal nombreHoja As String _
) As Worksheet

    On Error Resume Next

    Set OSSCC_ObtenerOCrearHoja = _
        wb.Worksheets(nombreHoja)

    On Error GoTo 0

    If OSSCC_ObtenerOCrearHoja Is Nothing Then

        Set OSSCC_ObtenerOCrearHoja = _
            wb.Worksheets.Add( _
                After:=wb.Worksheets( _
                    wb.Worksheets.Count _
                ) _
            )

        OSSCC_ObtenerOCrearHoja.Name = nombreHoja

    End If

End Function






```

### Módulo `I_Ofertas_ventana`

```vb
Attribute VB_Name = "I_Ofertas_ventana"
Option Explicit

Sub Resumir_Medidores_Central_Ventana_Oferta_Completa()

    Const FILA_INICIO As Long = 3
    Const TOLERANCIA As Double = 0.000001

    Dim ws As Worksheet

    Dim ultimaFilaOrigen As Long
    Dim ultimaFilaResumen As Long
    Dim ultimaFilaSalida As Long

    Dim arrG As Variant
    Dim arrL As Variant
    Dim arrR As Variant
    Dim arrSalida() As Variant

    Dim dicSuma As Object
    Dim dicDatos As Object

    Dim clave As String
    Dim claves As Variant
    Dim datosGrupo As Variant

    Dim valorG As Variant
    Dim valorL As Variant
    Dim valorR As Variant

    Dim ofertaTotal As Double
    Dim ofertaEsperada As Double

    Dim horaInicio As Long
    Dim ofertaEsperadaInicial As Long
    Dim ofertaEsperadaFinal As Long

    Dim ultimaVentana As Double
    Dim encontroVentanaNumerica As Boolean

    Dim ventanaNumerica As Double
    Dim esVentanaNumerica As Boolean

    Dim completa As Long
    Dim separador As String

    Dim i As Long
    Dim cantidadGrupos As Long

    Dim modoCalculo As XlCalculation
    Dim configuracionModificada As Boolean

    On Error GoTo ErrHandler

    Set ws = ThisWorkbook.Worksheets("Medidores")

    separador = Chr$(30)

    '=====================================================
    ' LEER HORA DE INICIO DESDE S1
    '=====================================================
    horaInicio = RESOF_ObtenerHoraInicio(ws.Range("S1").Value)

    If horaInicio < 1 Or horaInicio > 24 Then

        MsgBox _
            "El valor de Medidores!S1 no corresponde a una hora válida." & _
            vbCrLf & vbCrLf & _
            "Debe contener una hora entre 1 y 24, por ejemplo 10.", _
            vbCritical

        Exit Sub

    End If

    ' Si S1 = 10:
    ' Primera ventana = 9 horas * 4 = 36
    ' Última ventana = 15 horas * 4 = 60
    ofertaEsperadaInicial = (horaInicio - 1) * 4
    ofertaEsperadaFinal = (25 - horaInicio) * 4

    '=====================================================
    ' ACELERAR EXCEL
    '=====================================================
    With Application

        .ScreenUpdating = False
        .EnableEvents = False

        modoCalculo = .Calculation
        .Calculation = xlCalculationManual

        .StatusBar = "Generando resumen de ofertas..."

    End With

    configuracionModificada = True

    '=====================================================
    ' OBTENER ÚLTIMA FILA DEL ORIGEN
    '
    ' G = Central
    ' L = Ventana T
    ' R = Oferta
    '=====================================================
    ultimaFilaOrigen = RESOF_UltimaFilaColumnas( _
        ws, _
        Array("G", "L", "R") _
    )

    If ultimaFilaOrigen < FILA_INICIO Then

        MsgBox _
            "No existen datos para resumir desde la fila 3.", _
            vbExclamation

        GoTo Salida

    End If

    '=====================================================
    ' LEER DATOS EN MEMORIA
    '=====================================================
    arrG = ws.Range( _
        "G" & FILA_INICIO & _
        ":G" & ultimaFilaOrigen _
    ).Value2

    arrL = ws.Range( _
        "L" & FILA_INICIO & _
        ":L" & ultimaFilaOrigen _
    ).Value2

    arrR = ws.Range( _
        "R" & FILA_INICIO & _
        ":R" & ultimaFilaOrigen _
    ).Value2

    '=====================================================
    ' IDENTIFICAR LA ÚLTIMA VENTANA DEL MES
    '
    ' Se utiliza el mayor valor numérico existente en L.
    '=====================================================
    encontroVentanaNumerica = False
    ultimaVentana = 0

    For i = 1 To UBound(arrL, 1)

        If RESOF_EsNumeroValido(arrL(i, 1)) Then

            ventanaNumerica = CDbl(arrL(i, 1))

            If Not encontroVentanaNumerica Then

                ultimaVentana = ventanaNumerica
                encontroVentanaNumerica = True

            ElseIf ventanaNumerica > ultimaVentana Then

                ultimaVentana = ventanaNumerica

            End If

        End If

    Next i

    If Not encontroVentanaNumerica Then

        Err.Raise _
            vbObjectError + 1000, _
            "Resumir_Medidores_Central_Ventana_Oferta_Completa", _
            "No se encontraron valores numéricos en la columna L."

    End If

    '=====================================================
    ' CREAR DICCIONARIOS
    '=====================================================
    Set dicSuma = CreateObject("Scripting.Dictionary")
    Set dicDatos = CreateObject("Scripting.Dictionary")

    dicSuma.CompareMode = vbTextCompare
    dicDatos.CompareMode = vbTextCompare

    '=====================================================
    ' AGRUPAR POR G + L Y SUMAR R
    '=====================================================
    For i = 1 To UBound(arrG, 1)

        valorG = arrG(i, 1)
        valorL = arrL(i, 1)
        valorR = arrR(i, 1)

        ' Procesar únicamente filas con Central y Ventana
        If RESOF_TieneValor(valorG) And _
           RESOF_TieneValor(valorL) Then

            clave = _
                RESOF_ValorClave(valorG) & separador & _
                RESOF_ValorClave(valorL)

            If Not dicSuma.Exists(clave) Then

                dicSuma.Add clave, 0#

                dicDatos.Add clave, Array( _
                    RESOF_ValorSalida(valorG), _
                    RESOF_ValorSalida(valorL) _
                )

            End If

            ' Sumar únicamente valores numéricos de R
            If RESOF_EsNumeroValido(valorR) Then

                dicSuma(clave) = _
                    CDbl(dicSuma(clave)) + CDbl(valorR)

            End If

        End If

    Next i

    cantidadGrupos = dicSuma.Count

    '=====================================================
    ' BORRAR RESUMEN ANTERIOR DESDE AB3:AE
    '=====================================================
    ultimaFilaResumen = RESOF_UltimaFilaColumnas( _
        ws, _
        Array("AB", "AC", "AD", "AE") _
    )

    If ultimaFilaResumen >= FILA_INICIO Then

        ws.Range( _
            "AB" & FILA_INICIO & _
            ":AE" & ultimaFilaResumen _
        ).ClearContents

    End If

    '=====================================================
    ' ENCABEZADOS
    '=====================================================
    ws.Range("AB2:AE2").Value = Array( _
        "Central", _
        "Ventana T", _
        "Oferta", _
        "Completa" _
    )

    ws.Range("AB2:AE2").Font.Bold = True

    If cantidadGrupos = 0 Then

        MsgBox _
            "No se encontraron filas válidas para generar el resumen.", _
            vbExclamation

        GoTo Salida

    End If

    '=====================================================
    ' PREPARAR RESULTADO
    '=====================================================
    ReDim arrSalida(1 To cantidadGrupos, 1 To 4)

    claves = dicSuma.Keys

    For i = 0 To cantidadGrupos - 1

        datosGrupo = dicDatos.Item(CStr(claves(i)))

        valorG = datosGrupo(0)
        valorL = datosGrupo(1)
        ofertaTotal = CDbl(dicSuma.Item(CStr(claves(i))))

        esVentanaNumerica = RESOF_EsNumeroValido(valorL)

        ' Valor normal por defecto
        ofertaEsperada = 96

        If esVentanaNumerica Then

            ventanaNumerica = CDbl(valorL)

            '---------------------------------------------
            ' VENTANA 0
            '---------------------------------------------
            If Abs(ventanaNumerica) < TOLERANCIA Then

                ofertaEsperada = ofertaEsperadaInicial

            '---------------------------------------------
            ' ÚLTIMA VENTANA DEL MES
            '---------------------------------------------
            ElseIf Abs( _
                ventanaNumerica - ultimaVentana _
            ) < TOLERANCIA Then

                ofertaEsperada = ofertaEsperadaFinal

            End If

        End If

        If Abs(ofertaTotal - ofertaEsperada) < TOLERANCIA Then
            completa = 1
        Else
            completa = 0
        End If

        arrSalida(i + 1, 1) = valorG
        arrSalida(i + 1, 2) = valorL
        arrSalida(i + 1, 3) = ofertaTotal
        arrSalida(i + 1, 4) = completa

    Next i

    '=====================================================
    ' PEGAR DESDE AB3
    '=====================================================
    ws.Range("AB" & FILA_INICIO).Resize( _
        cantidadGrupos, _
        4 _
    ).Value2 = arrSalida

    ultimaFilaSalida = _
        FILA_INICIO + cantidadGrupos - 1

    '=====================================================
    ' ORDENAR:
    '
    ' 1. Central
    ' 2. Ventana T
    '=====================================================
    With ws.Sort

        .SortFields.Clear

        .SortFields.Add _
            key:=ws.Range( _
                "AB" & FILA_INICIO & _
                ":AB" & ultimaFilaSalida _
            ), _
            SortOn:=xlSortOnValues, _
            Order:=xlAscending, _
            DataOption:=xlSortNormal

        .SortFields.Add _
            key:=ws.Range( _
                "AC" & FILA_INICIO & _
                ":AC" & ultimaFilaSalida _
            ), _
            SortOn:=xlSortOnValues, _
            Order:=xlAscending, _
            DataOption:=xlSortNormal

        .SetRange ws.Range( _
            "AB2:AE" & ultimaFilaSalida _
        )

        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply

    End With

    '=====================================================
    ' FORMATO
    '=====================================================
    ws.Columns("AB:AE").AutoFit

    ws.Range( _
        "AD3:AD" & ultimaFilaSalida _
    ).NumberFormat = "#,##0.00"

    ws.Range( _
        "AE3:AE" & ultimaFilaSalida _
    ).NumberFormat = "0"

    MsgBox _
        "Resumen generado correctamente." & vbCrLf & vbCrLf & _
        "Hora de inicio en S1: " & horaInicio & vbCrLf & _
        "Última ventana detectada: " & ultimaVentana & vbCrLf & vbCrLf & _
        "Oferta esperada ventana 0: " & _
        ofertaEsperadaInicial & vbCrLf & _
        "Oferta esperada ventanas normales: 96" & vbCrLf & _
        "Oferta esperada última ventana: " & _
        ofertaEsperadaFinal & vbCrLf & vbCrLf & _
        "Grupos generados: " & cantidadGrupos & vbCrLf & _
        "Resultado: Medidores!AB3:AE" & ultimaFilaSalida, _
        vbInformation

Salida:

    Application.StatusBar = False

    If configuracionModificada Then

        With Application
            .ScreenUpdating = True
            .EnableEvents = True
            .Calculation = modoCalculo
        End With

    End If

    Exit Sub

ErrHandler:

    MsgBox _
        "Error al generar el resumen:" & vbCrLf & vbCrLf & _
        Err.Description, _
        vbCritical

    Resume Salida

End Sub


Private Function RESOF_ObtenerHoraInicio( _
    ByVal valor As Variant _
) As Long

    Dim numero As Double
    Dim texto As String

    RESOF_ObtenerHoraInicio = -1

    If IsError(valor) Or IsEmpty(valor) Then Exit Function

    If IsNumeric(valor) Then

        numero = CDbl(valor)

        ' Hora de Excel, por ejemplo 10:00
        If numero >= 0 And numero < 1 Then

            RESOF_ObtenerHoraInicio = Hour(CDate(numero))
            Exit Function

        End If

        ' Hora escrita como número entero
        If numero = Fix(numero) And _
           numero >= 1 And numero <= 24 Then

            RESOF_ObtenerHoraInicio = CLng(numero)
            Exit Function

        End If

    End If

    texto = Trim$(CStr(valor))

    If Len(texto) = 0 Then Exit Function

    If IsDate(texto) Then

        RESOF_ObtenerHoraInicio = Hour(CDate(texto))

        If RESOF_ObtenerHoraInicio = 0 And _
           Left$(texto, 2) = "24" Then

            RESOF_ObtenerHoraInicio = 24

        End If

    End If

End Function


Private Function RESOF_UltimaFilaColumnas( _
    ByVal ws As Worksheet, _
    ByVal columnas As Variant _
) As Long

    Dim columna As Variant
    Dim celdaEncontrada As Range
    Dim maximaFila As Long

    maximaFila = 0

    For Each columna In columnas

        Set celdaEncontrada = ws.Columns( _
            CStr(columna) _
        ).Find( _
            What:="*", _
            LookIn:=xlFormulas, _
            LookAt:=xlPart, _
            SearchOrder:=xlByRows, _
            SearchDirection:=xlPrevious, _
            MatchCase:=False _
        )

        If Not celdaEncontrada Is Nothing Then

            If celdaEncontrada.Row > maximaFila Then
                maximaFila = celdaEncontrada.Row
            End If

        End If

        Set celdaEncontrada = Nothing

    Next columna

    RESOF_UltimaFilaColumnas = maximaFila

End Function


Private Function RESOF_TieneValor( _
    ByVal valor As Variant _
) As Boolean

    If IsError(valor) Or IsEmpty(valor) Then

        RESOF_TieneValor = False

    Else

        RESOF_TieneValor = _
            Len(Trim$(CStr(valor))) > 0

    End If

End Function


Private Function RESOF_EsNumeroValido( _
    ByVal valor As Variant _
) As Boolean

    If IsError(valor) Or IsEmpty(valor) Then

        RESOF_EsNumeroValido = False

    Else

        RESOF_EsNumeroValido = _
            IsNumeric(valor) And _
            Len(Trim$(CStr(valor))) > 0

    End If

End Function


Private Function RESOF_ValorClave( _
    ByVal valor As Variant _
) As String

    If IsError(valor) Or IsEmpty(valor) Then

        RESOF_ValorClave = ""

    ElseIf IsNumeric(valor) Then

        RESOF_ValorClave = _
            Format$(CDbl(valor), "0.###############")

    Else

        RESOF_ValorClave = _
            UCase$(Trim$(CStr(valor)))

    End If

End Function


Private Function RESOF_ValorSalida( _
    ByVal valor As Variant _
) As Variant

    If IsError(valor) Or IsEmpty(valor) Then

        RESOF_ValorSalida = vbNullString

    Else

        RESOF_ValorSalida = valor

    End If

End Function


```

### Módulo `J_Calculo_Ecostos`

```vb
Attribute VB_Name = "J_Calculo_Ecostos"

'==========================================================
' CONFIGURACIÓN GENERAL
'==========================================================

Private Const FILA_INICIO As Long = 4
Private Const FILA_INICIO_RESUMEN As Long = 8

Private Const HOJA_DESTINO As String = "Calculo E Costos"
Private Const HOJA_SUBASTAS As String = "Subastas"
Private Const HOJA_RESUMEN As String = "Resumen"
Private Const HOJA_PRORRATA As String = "Prorrata SSCC"
Private Const HOJA_DICCIONARIO As String = "Diccionario"
Private Const HOJA_FD As String = "FD"
Private Const HOJA_LOG As String = "Log Calculo E Costos"

Private Const SEP As String = "¦"


'==========================================================
' MACRO PRINCIPAL
'
' COLUMNAS CALCULADAS:
'
' L, M, N, O
' R, S, T, U
' W, X, Y
' AB, AC, AD, AE, AF
' AG hasta AX
' AZ
'
' COLUMNAS NO MODIFICADAS:
'
' P, Q, Z, AA, AY
'==========================================================

Public Sub Actualizar_Calculos_Columnas()

    Dim inicioProceso As Date
    Dim inicioTimer As Double
    Dim duracionCalculo As Double

    Dim wb As Workbook

    Dim wsDestino As Worksheet
    Dim wsSubastas As Worksheet
    Dim wsResumen As Worksheet
    Dim wsProrrata As Worksheet
    Dim wsDiccionario As Worksheet
    Dim wsFD As Worksheet

    Dim ultimaFila As Long
    Dim ultimaSubastasDK As Long
    Dim ultimaSubastasUW As Long
    Dim ultimaProrrata As Long
    Dim ultimaDiccionario As Long
    Dim ultimaResumenBC As Long
    Dim ultimaFDAL As Long
    Dim ultimaFDQAD As Long

    Dim cantidadFilas As Long

    Dim datosDestino As Variant
    Dim datosSubastasDK As Variant
    Dim datosSubastasUW As Variant
    Dim datosProrrata As Variant
    Dim datosDiccionario As Variant
    Dim datosFDAL As Variant
    Dim datosFDQAD As Variant

    Dim dicSubastas As Object
    Dim dicUmbralesSubastas As Object
    Dim dicProrrata As Object
    Dim dicMapeo As Object
    Dim dicFDAL As Object
    Dim dicFDQAD As Object
    Dim dicResumen As Object
    Dim dicFaltantesFD As Object

    Dim dicGrupos As Object
    Dim dicSumaIPorP As Object
    Dim dicSumaJPorP As Object

    Dim filasGrupo As Collection
    Dim clavesGrupos As Variant

    Dim claveSubasta As String
    Dim claveGrupo As String
    Dim claveP As String
    Dim claveAuxiliar As String
    Dim claveFDY As String
    Dim claveFDAC As String

    Dim valoresF() As Double
    Dim valoresI() As Double
    Dim valoresJ() As Double
    Dim valoresK() As Double
    Dim valoresQ() As Double
    Dim valoresW() As Double

    Dim resultadoL() As Byte

    'L:M:N:O
    Dim salidaLMNO() As Variant

    'R:S:T:U
    Dim salidaRSTU() As Variant

    'W:X:Y
    Dim salidaWXY() As Variant

    'AB:AC:AD:AE:AF
    Dim salidaABAF() As Variant

    'AG:AX
    Dim salidaAGAX() As Variant

    'AZ
    Dim salidaAZ() As Variant

    Dim indicesF() As Long
    Dim indicesRanking() As Long

    Dim indicesCalificaI() As Long
    Dim indicesNoCalificaI() As Long

    Dim indicesCalificaJ() As Long
    Dim indicesNoCalificaJ() As Long

    Dim i As Long
    Dim j As Long

    Dim posicion As Long
    Dim finBloque As Long
    Dim cantidadGrupo As Long

    Dim cantidadCalificaI As Long
    Dim cantidadNoCalificaI As Long

    Dim cantidadCalificaJ As Long
    Dim cantidadNoCalificaJ As Long

    Dim filaIndice As Long
    Dim filaOrigen As Long
    Dim siguienteIndice As Long

    Dim valorFActual As Double
    Dim valorQActual As Double

    Dim sumaBloqueI As Double
    Dim sumaBloqueJ As Double

    Dim acumuladoI As Double
    Dim acumuladoJ As Double

    Dim valorS As Double
    Dim valorT As Double

    Dim maximoN As Double
    Dim maximoO As Double

    Dim valorUmbralM As Variant
    Dim umbralM As Double

    Dim valorFactor As Variant
    Dim factorGrupo As Double

    Dim datosPar As Variant
    Dim datosFD As Variant
    Dim datosUmbral As Variant

    Dim valorAG As Double
    Dim valorAH As Double

    Dim valorAM As Variant
    Dim valorAN As Variant
    Dim valorAP As Variant
    Dim valorAQ As Variant

    Dim valorAS As Variant
    Dim valorAT As Variant

    Dim valorAE As Variant
    Dim valorAF As Variant

    Dim valorAU As Double
    Dim valorAV As Double
    Dim valorAW As Double
    Dim valorAX As Double

    Dim mapeoG As Variant

    Dim bloqueY As Long
    Dim bloqueAC As Long

    Dim sumaIP As Double
    Dim sumaJP As Double

    Dim sumaABAE As Double
    Dim cantidadABAE As Long
    Dim errorAEGrupo As Boolean

    Dim sumaADAF As Double
    Dim cantidadADAF As Long
    Dim errorAFGrupo As Boolean

    Dim promedioABAE As Double
    Dim promedioADAF As Double

    Dim sumaABW As Double
    Dim cantidadABW As Long

    Dim sumaADW As Double
    Dim cantidadADW As Long

    Dim promedioABW As Double
    Dim promedioADW As Double

    Dim umbralBajada As Double
    Dim umbralSubida As Double
    Dim umbralesValidos As Boolean

    Dim sumaAXGrupo As Double
    Dim sumaUGrupo As Double
    Dim valorAZGrupo As Double

    Dim vAB As Variant
    Dim vAD As Variant

    Dim calculoAnterior As XlCalculation
    Dim eventosAnteriores As Boolean
    Dim pantallaAnterior As Boolean
    Dim barraAnterior As Variant

    Dim estadoGuardado As Boolean
    Dim procesoCorrecto As Boolean

    Dim numeroError As Long
    Dim descripcionError As String
    Dim nombreHojaError As String

    Dim mensajeFaltantesFD As String
    Dim tipoMensajeFinal As VbMsgBoxStyle

    inicioProceso = Now
    inicioTimer = Timer

    On Error GoTo ManejoError

    Set wb = ThisWorkbook

    '======================================================
    ' VALIDAR HOJAS
    '======================================================

    If Not ExisteHoja(HOJA_DESTINO, wb) Then
        MsgBox "No existe la hoja '" & HOJA_DESTINO & "'.", _
               vbExclamation, "Cálculo E Costos"
        Exit Sub
    End If

    If Not ExisteHoja(HOJA_SUBASTAS, wb) Then
        MsgBox "No existe la hoja '" & HOJA_SUBASTAS & "'.", _
               vbExclamation, "Cálculo E Costos"
        Exit Sub
    End If

    If Not ExisteHoja(HOJA_RESUMEN, wb) Then
        MsgBox "No existe la hoja '" & HOJA_RESUMEN & "'.", _
               vbExclamation, "Cálculo E Costos"
        Exit Sub
    End If

    If Not ExisteHoja(HOJA_PRORRATA, wb) Then
        MsgBox "No existe la hoja '" & HOJA_PRORRATA & "'.", _
               vbExclamation, "Cálculo E Costos"
        Exit Sub
    End If

    If Not ExisteHoja(HOJA_DICCIONARIO, wb) Then
        MsgBox "No existe la hoja '" & HOJA_DICCIONARIO & "'.", _
               vbExclamation, "Cálculo E Costos"
        Exit Sub
    End If

    If Not ExisteHoja(HOJA_FD, wb) Then
        MsgBox "No existe la hoja '" & HOJA_FD & "'.", _
               vbExclamation, "Cálculo E Costos"
        Exit Sub
    End If

    Set wsDestino = wb.Worksheets(HOJA_DESTINO)
    Set wsSubastas = wb.Worksheets(HOJA_SUBASTAS)
    Set wsResumen = wb.Worksheets(HOJA_RESUMEN)
    Set wsProrrata = wb.Worksheets(HOJA_PRORRATA)
    Set wsDiccionario = wb.Worksheets(HOJA_DICCIONARIO)
    Set wsFD = wb.Worksheets(HOJA_FD)

    '======================================================
    ' UMBRAL PARA COLUMNA M
    '======================================================

    valorUmbralM = wsResumen.Range("H8").Value2

    If IsError(valorUmbralM) Then

        MsgBox "Resumen!H8 contiene un error.", _
               vbExclamation, "Cálculo E Costos"
        Exit Sub

    ElseIf IsNumeric(valorUmbralM) Then

        umbralM = CDbl(valorUmbralM)

    ElseIf Len(Trim$(TextoSeguro(valorUmbralM))) = 0 Then

        umbralM = 0

    Else

        MsgBox "Resumen!H8 debe contener un valor numérico.", _
               vbExclamation, "Cálculo E Costos"
        Exit Sub

    End If

    '======================================================
    ' OBTENER ÚLTIMAS FILAS
    '======================================================

    'Detecta la última fila considerando únicamente columnas de entrada.
    'Así no se cortan registros si la última fila tiene G vacío pero
    'contiene datos en otra columna necesaria para el cálculo.
    ultimaFila = ObtenerUltimaFila( _
        wsDestino, _
        Array("A", "B", "C", "D", "F", "G", "I", "J", "K", "P", "Q"), _
        FILA_INICIO)

    If ultimaFila < FILA_INICIO Then

        MsgBox "No hay datos desde la fila " & FILA_INICIO & ".", _
               vbInformation, "Cálculo E Costos"
        Exit Sub

    End If

    ultimaSubastasDK = ObtenerUltimaFila( _
        wsSubastas, Array("D", "G", "H", "I", "K"), 1)

    ultimaSubastasUW = ObtenerUltimaFila( _
        wsSubastas, Array("U", "V", "W"), 1)

    ultimaProrrata = ObtenerUltimaFila( _
        wsProrrata, Array("A", "B", "C", "D"), 1)

    ultimaDiccionario = wsDiccionario.Cells( _
        wsDiccionario.Rows.Count, "A").End(xlUp).Row

    'Resumen B:C es una tabla variable que comienza en la fila 8.
    'Su término ya no queda limitado a la fila 12.
    ultimaResumenBC = ObtenerUltimaFila( _
        wsResumen, Array("B", "C"), FILA_INICIO_RESUMEN)

    ultimaFDAL = wsFD.Cells( _
        wsFD.Rows.Count, "A").End(xlUp).Row

    ultimaFDQAD = wsFD.Cells( _
        wsFD.Rows.Count, "Q").End(xlUp).Row

    cantidadFilas = ultimaFila - FILA_INICIO + 1

    '======================================================
    ' GUARDAR ESTADO DE EXCEL
    '======================================================

    calculoAnterior = Application.Calculation
    eventosAnteriores = Application.EnableEvents
    pantallaAnterior = Application.ScreenUpdating
    barraAnterior = Application.StatusBar

    estadoGuardado = True

    '======================================================
    ' MÁXIMA VELOCIDAD
    '======================================================

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "Cargando datos en memoria..."

    '======================================================
    ' CARGAR DATOS EN MATRICES
    '======================================================

    datosDestino = wsDestino.Range( _
        "A" & FILA_INICIO & ":Q" & ultimaFila).Value2

    datosSubastasDK = wsSubastas.Range( _
        "D1:K" & ultimaSubastasDK).Value2

    datosSubastasUW = wsSubastas.Range( _
        "U1:W" & ultimaSubastasUW).Value2

    datosProrrata = wsProrrata.Range( _
        "A1:D" & ultimaProrrata).Value2

    datosDiccionario = wsDiccionario.Range( _
        "A1:B" & ultimaDiccionario).Value2

    datosFDAL = wsFD.Range( _
        "A1:L" & ultimaFDAL).Value2

    datosFDQAD = wsFD.Range( _
        "Q1:AD" & ultimaFDQAD).Value2

    '======================================================
    ' CREAR DICCIONARIOS
    '======================================================

    Application.StatusBar = "Creando índices de búsqueda..."

    Set dicSubastas = _
        CrearDiccionarioSubastas(datosSubastasDK)

    Set dicUmbralesSubastas = _
        CrearDiccionarioUmbralesSubastas(datosSubastasUW)

    Set dicProrrata = _
        CrearDiccionarioProrrata(datosProrrata)

    Set dicMapeo = CrearDiccionarioPrimerValor( _
        datosDiccionario, 1, 2)

    Set dicFDAL = CrearDiccionarioFDAL(datosFDAL)
    Set dicFDQAD = CrearDiccionarioFDQAD(datosFDQAD)

    Set dicFaltantesFD = CreateObject("Scripting.Dictionary")
    dicFaltantesFD.CompareMode = vbTextCompare

    If ultimaResumenBC >= FILA_INICIO_RESUMEN Then

        Set dicResumen = CrearDiccionarioPrimerValor( _
            wsResumen.Range( _
                "B" & FILA_INICIO_RESUMEN & _
                ":C" & ultimaResumenBC).Value2, _
            1, 2)

    Else

        'Si no hay tabla de factores, se crea un diccionario vacío.
        'AE/AF conservarán el comportamiento actual: #N/D al no
        'encontrarse la clave G correspondiente.
        Set dicResumen = CreateObject("Scripting.Dictionary")
        dicResumen.CompareMode = vbTextCompare

    End If

    Set dicGrupos = CreateObject("Scripting.Dictionary")
    dicGrupos.CompareMode = vbTextCompare

    Set dicSumaIPorP = CreateObject("Scripting.Dictionary")
    dicSumaIPorP.CompareMode = vbTextCompare

    Set dicSumaJPorP = CreateObject("Scripting.Dictionary")
    dicSumaJPorP.CompareMode = vbTextCompare

    '======================================================
    ' DIMENSIONAR ARREGLOS
    '======================================================

    ReDim valoresF(1 To cantidadFilas)
    ReDim valoresI(1 To cantidadFilas)
    ReDim valoresJ(1 To cantidadFilas)
    ReDim valoresK(1 To cantidadFilas)
    ReDim valoresQ(1 To cantidadFilas)
    ReDim valoresW(1 To cantidadFilas)

    ReDim resultadoL(1 To cantidadFilas)

    ReDim salidaLMNO(1 To cantidadFilas, 1 To 4)
    ReDim salidaRSTU(1 To cantidadFilas, 1 To 4)
    ReDim salidaWXY(1 To cantidadFilas, 1 To 3)
    ReDim salidaABAF(1 To cantidadFilas, 1 To 5)
    ReDim salidaAGAX(1 To cantidadFilas, 1 To 18)
    ReDim salidaAZ(1 To cantidadFilas, 1 To 1)

    '======================================================
    ' PRIMER RECORRIDO
    '
    ' A = 1
    ' B = 2
    ' C = 3
    ' D = 4
    ' F = 6
    ' G = 7
    ' I = 9
    ' J = 10
    ' K = 11
    ' P = 16
    ' Q = 17
    '======================================================

    Application.StatusBar = "Calculando L, M, W y X..."

    For i = 1 To cantidadFilas

        valoresF(i) = NumeroSeguro(datosDestino(i, 6))
        valoresI(i) = NumeroSeguro(datosDestino(i, 9))
        valoresJ(i) = NumeroSeguro(datosDestino(i, 10))
        valoresK(i) = NumeroSeguro(datosDestino(i, 11))
        valoresQ(i) = NumeroSeguro(datosDestino(i, 17))

        '--------------------------------------------------
        ' L
        '--------------------------------------------------

        claveSubasta = CrearClave4( _
            datosDestino(i, 7), _
            datosDestino(i, 1), _
            datosDestino(i, 2), _
            datosDestino(i, 3))

        If dicSubastas.Exists(claveSubasta) Then
            resultadoL(i) = 1
        Else
            resultadoL(i) = 0
        End If

        salidaLMNO(i, 1) = resultadoL(i)

        '--------------------------------------------------
        ' M
        '--------------------------------------------------

        If valoresK(i) > umbralM Then
            salidaLMNO(i, 2) = 1
        Else
            salidaLMNO(i, 2) = 0
        End If

        '--------------------------------------------------
        ' X = P
        '--------------------------------------------------

        salidaWXY(i, 2) = datosDestino(i, 16)

        '--------------------------------------------------
        ' W
        '
        ' Fila 4:
        ' W4 = C4
        '
        ' Desde fila 5:
        ' SI(X actual=X anterior;W anterior+1;1)
        '
        ' Como X=P, se compara P actual con P anterior.
        '--------------------------------------------------

        If i = 1 Then

            valoresW(i) = NumeroSeguro(datosDestino(i, 3))
            salidaWXY(i, 1) = datosDestino(i, 3)

        ElseIf ValoresIgualesExcel( _
            datosDestino(i, 16), _
            datosDestino(i - 1, 16)) Then

            valoresW(i) = valoresW(i - 1) + 1#
            salidaWXY(i, 1) = valoresW(i)

        Else

            valoresW(i) = 1#
            salidaWXY(i, 1) = 1

        End If

        '--------------------------------------------------
        ' GRUPO G + P
        '--------------------------------------------------

        claveGrupo = CrearClave2( _
            datosDestino(i, 7), _
            datosDestino(i, 16))

        If dicGrupos.Exists(claveGrupo) Then

            Set filasGrupo = dicGrupos.Item(claveGrupo)

        Else

            Set filasGrupo = New Collection
            dicGrupos.Add claveGrupo, filasGrupo

        End If

        filasGrupo.Add i

        '--------------------------------------------------
        ' SUMAS GLOBALES POR P
        '--------------------------------------------------

        claveP = NormalizarValor(datosDestino(i, 16))

        AgregarSuma dicSumaIPorP, claveP, valoresI(i)
        AgregarSuma dicSumaJPorP, claveP, valoresJ(i)

    Next i

    clavesGrupos = dicGrupos.Keys

    '======================================================
    ' PROCESAR GRUPOS G + P
    '
    ' N, O, R, Y, AB, AC, AD, AE Y AF
    '======================================================

    Application.StatusBar = "Calculando columnas agrupadas..."

    For i = LBound(clavesGrupos) To UBound(clavesGrupos)

        Set filasGrupo = dicGrupos.Item(clavesGrupos(i))

        cantidadGrupo = filasGrupo.Count

        ReDim indicesF(1 To cantidadGrupo)
        ReDim indicesRanking(1 To cantidadGrupo)

        ReDim indicesCalificaI(1 To cantidadGrupo)
        ReDim indicesNoCalificaI(1 To cantidadGrupo)

        ReDim indicesCalificaJ(1 To cantidadGrupo)
        ReDim indicesNoCalificaJ(1 To cantidadGrupo)

        cantidadCalificaI = 0
        cantidadNoCalificaI = 0

        cantidadCalificaJ = 0
        cantidadNoCalificaJ = 0

        For j = 1 To cantidadGrupo

            filaIndice = CLng(filasGrupo(j))

            indicesF(j) = filaIndice
            indicesRanking(j) = filaIndice

            If resultadoL(filaIndice) = 1 And _
               valoresI(filaIndice) <> 0 Then

                cantidadCalificaI = cantidadCalificaI + 1
                indicesCalificaI(cantidadCalificaI) = filaIndice

            Else

                cantidadNoCalificaI = cantidadNoCalificaI + 1
                indicesNoCalificaI(cantidadNoCalificaI) = filaIndice

            End If

            If resultadoL(filaIndice) = 1 And _
               valoresJ(filaIndice) <> 0 Then

                cantidadCalificaJ = cantidadCalificaJ + 1
                indicesCalificaJ(cantidadCalificaJ) = filaIndice

            Else

                cantidadNoCalificaJ = cantidadNoCalificaJ + 1
                indicesNoCalificaJ(cantidadNoCalificaJ) = filaIndice

            End If

        Next j

        '==================================================
        ' N Y O
        '==================================================

        If cantidadGrupo > 1 Then
            QuickSortFDesc indicesF, 1, cantidadGrupo, valoresF
        End If

        acumuladoI = 0
        acumuladoJ = 0
        posicion = 1

        Do While posicion <= cantidadGrupo

            valorFActual = valoresF(indicesF(posicion))
            finBloque = posicion

            Do While finBloque < cantidadGrupo

                siguienteIndice = indicesF(finBloque + 1)

                If valoresF(siguienteIndice) = valorFActual Then
                    finBloque = finBloque + 1
                Else
                    Exit Do
                End If

            Loop

            sumaBloqueI = 0
            sumaBloqueJ = 0

            For j = posicion To finBloque

                filaIndice = indicesF(j)

                If resultadoL(filaIndice) = 1 Then

                    sumaBloqueI = _
                        sumaBloqueI + valoresI(filaIndice)

                    sumaBloqueJ = _
                        sumaBloqueJ + valoresJ(filaIndice)

                End If

            Next j

            acumuladoI = acumuladoI + sumaBloqueI
            acumuladoJ = acumuladoJ + sumaBloqueJ

            For j = posicion To finBloque

                filaIndice = indicesF(j)

                salidaLMNO(filaIndice, 3) = acumuladoI
                salidaLMNO(filaIndice, 4) = -acumuladoJ

            Next j

            posicion = finBloque + 1

        Loop

        '==================================================
        ' R
        '==================================================

        If cantidadGrupo > 1 Then

            QuickSortQFDesc _
                indicesRanking, _
                1, _
                cantidadGrupo, _
                valoresQ, _
                valoresF

        End If

        posicion = 1

        Do While posicion <= cantidadGrupo

            filaIndice = indicesRanking(posicion)

            valorQActual = valoresQ(filaIndice)
            valorFActual = valoresF(filaIndice)

            finBloque = posicion

            Do While finBloque < cantidadGrupo

                siguienteIndice = indicesRanking(finBloque + 1)

                If valoresQ(siguienteIndice) = valorQActual And _
                   valoresF(siguienteIndice) = valorFActual Then

                    finBloque = finBloque + 1

                Else

                    Exit Do

                End If

            Loop

            For j = posicion To finBloque

                filaIndice = indicesRanking(j)
                salidaRSTU(filaIndice, 1) = posicion

            Next j

            posicion = finBloque + 1

        Loop

        '==================================================
        ' Y Y AB
        '==================================================

        If cantidadCalificaI > 1 Then

            QuickSortQDescFilaAsc _
                indicesCalificaI, _
                1, _
                cantidadCalificaI, _
                valoresQ

        End If

        For j = 1 To cantidadGrupo

            filaIndice = CLng(filasGrupo(j))

            If j <= cantidadCalificaI Then

                filaOrigen = indicesCalificaI(j)

                If IsError(datosDestino(filaOrigen, 6)) Then
                    salidaWXY(filaIndice, 3) = vbNullString
                Else
                    salidaWXY(filaIndice, 3) = _
                        datosDestino(filaOrigen, 6)
                End If

                salidaABAF(filaIndice, 1) = _
                    valoresQ(filaOrigen)

            Else

                filaOrigen = _
                    indicesNoCalificaI(j - cantidadCalificaI)

                If IsError(datosDestino(filaOrigen, 6)) Then
                    salidaWXY(filaIndice, 3) = vbNullString
                Else
                    salidaWXY(filaIndice, 3) = _
                        datosDestino(filaOrigen, 6)
                End If

                salidaABAF(filaIndice, 1) = vbNullString

            End If

        Next j

        '==================================================
        ' AC Y AD
        '==================================================

        If cantidadCalificaJ > 1 Then

            QuickSortQAscFilaAsc _
                indicesCalificaJ, _
                1, _
                cantidadCalificaJ, _
                valoresQ

        End If

        For j = 1 To cantidadGrupo

            filaIndice = CLng(filasGrupo(j))

            If j <= cantidadCalificaJ Then

                filaOrigen = indicesCalificaJ(j)

                If IsError(datosDestino(filaOrigen, 6)) Then
                    salidaABAF(filaIndice, 2) = vbNullString
                Else
                    salidaABAF(filaIndice, 2) = _
                        datosDestino(filaOrigen, 6)
                End If

                salidaABAF(filaIndice, 3) = _
                    valoresQ(filaOrigen)

            Else

                filaOrigen = _
                    indicesNoCalificaJ(j - cantidadCalificaJ)

                If IsError(datosDestino(filaOrigen, 6)) Then
                    salidaABAF(filaIndice, 2) = vbNullString
                Else
                    salidaABAF(filaIndice, 2) = _
                        datosDestino(filaOrigen, 6)
                End If

                salidaABAF(filaIndice, 3) = vbNullString

            End If

        Next j

        '==================================================
        ' MÁXIMOS N Y O DEL GRUPO
        '==================================================

        filaIndice = CLng(filasGrupo(1))

        maximoN = CDbl(salidaLMNO(filaIndice, 3))
        maximoO = CDbl(salidaLMNO(filaIndice, 4))

        For j = 2 To cantidadGrupo

            filaIndice = CLng(filasGrupo(j))

            If CDbl(salidaLMNO(filaIndice, 3)) > maximoN Then
                maximoN = CDbl(salidaLMNO(filaIndice, 3))
            End If

            If CDbl(salidaLMNO(filaIndice, 4)) > maximoO Then
                maximoO = CDbl(salidaLMNO(filaIndice, 4))
            End If

        Next j

        '==================================================
        ' AE Y AF
        '==================================================

        filaIndice = CLng(filasGrupo(1))

        claveAuxiliar = _
            NormalizarValor(datosDestino(filaIndice, 7))

        If dicResumen.Exists(claveAuxiliar) Then
            valorFactor = dicResumen.Item(claveAuxiliar)
        Else
            valorFactor = CVErr(xlErrNA)
        End If

        For j = 1 To cantidadGrupo

            filaIndice = CLng(filasGrupo(j))

            If IsError(valorFactor) Then

                salidaABAF(filaIndice, 4) = valorFactor
                salidaABAF(filaIndice, 5) = valorFactor

            ElseIf Not IsNumeric(valorFactor) Then

                salidaABAF(filaIndice, 4) = CVErr(xlErrValue)
                salidaABAF(filaIndice, 5) = CVErr(xlErrValue)

            ElseIf CDbl(valorFactor) = 0 Then

                salidaABAF(filaIndice, 4) = CVErr(xlErrDiv0)
                salidaABAF(filaIndice, 5) = CVErr(xlErrDiv0)

            Else

                factorGrupo = CDbl(valorFactor)

                salidaABAF(filaIndice, 4) = _
                    CalcularAsignacionEnergia( _
                        valoresW(filaIndice), _
                        maximoN, _
                        factorGrupo)

                salidaABAF(filaIndice, 5) = _
                    -CalcularAsignacionEnergia( _
                        valoresW(filaIndice), _
                        maximoO, _
                        factorGrupo)

            End If

        Next j

    Next i

    '======================================================
    ' S, T Y U
    '======================================================

    Application.StatusBar = "Calculando S, T y U..."

    For i = 1 To cantidadFilas

        valorS = valoresI(i) * valoresQ(i)
        valorT = valoresJ(i) * valoresQ(i)

        salidaRSTU(i, 2) = valorS
        salidaRSTU(i, 3) = valorT

        salidaRSTU(i, 4) = _
            CDbl(resultadoL(i)) * (valorS + valorT)

    Next i

    '======================================================
    ' AG HASTA AT
    '======================================================

    Application.StatusBar = "Calculando AG hasta AT..."

    For i = 1 To cantidadFilas

        '--------------------------------------------------
        ' AG, AH, AJ Y AK
        '--------------------------------------------------

        claveAuxiliar = CrearClave2( _
            datosDestino(i, 7), _
            datosDestino(i, 4))

        If dicProrrata.Exists(claveAuxiliar) Then

            datosPar = dicProrrata.Item(claveAuxiliar)

            valorAG = CDbl(datosPar(0))
            valorAH = CDbl(datosPar(1))

        Else

            valorAG = 0
            valorAH = 0

        End If

        salidaAGAX(i, 1) = valorAG
        salidaAGAX(i, 2) = valorAH
        salidaAGAX(i, 3) = 0

        salidaAGAX(i, 4) = valorAG
        salidaAGAX(i, 5) = valorAH
        salidaAGAX(i, 6) = 0

        '--------------------------------------------------
        ' AM, AN, AP Y AQ
        '--------------------------------------------------

        claveAuxiliar = _
            NormalizarValor(datosDestino(i, 7))

        If dicMapeo.Exists(claveAuxiliar) Then

            mapeoG = dicMapeo.Item(claveAuxiliar)

            If IsError(mapeoG) Then

                valorAM = mapeoG
                valorAN = mapeoG
                valorAP = mapeoG
                valorAQ = mapeoG

            Else

                bloqueY = CalcularBloque( _
                    NumeroSeguro(salidaWXY(i, 3)))

                bloqueAC = CalcularBloque( _
                    NumeroSeguro(salidaABAF(i, 2)))

                claveFDY = NormalizarValor( _
                    CStr(bloqueY) & TextoSeguro(mapeoG))

                claveFDAC = NormalizarValor( _
                    CStr(bloqueAC) & TextoSeguro(mapeoG))

                If dicFDQAD.Exists(claveFDY) Then

                    datosFD = dicFDQAD.Item(claveFDY)

                    valorAP = datosFD(0)
                    valorAM = datosFD(1)

                Else

                    'No se encontró la clave en FD Q:AD:
                    'AM y AP toman 0 y se registra el faltante.
                    valorAM = 0
                    valorAP = 0

                    RegistrarFaltanteFD _
                        dicFaltantesFD, _
                        "FD Q:AD", _
                        claveFDY, _
                        FILA_INICIO + i - 1, _
                        "AM / AP"

                End If

                If dicFDAL.Exists(claveFDY) Then

                    datosFD = dicFDAL.Item(claveFDY)
                    valorAN = datosFD(1)

                Else

                    'No se encontró la clave en FD A:L:
                    'AN toma 0 y se registra el faltante.
                    valorAN = 0

                    RegistrarFaltanteFD _
                        dicFaltantesFD, _
                        "FD A:L", _
                        claveFDY, _
                        FILA_INICIO + i - 1, _
                        "AN"

                End If

                If dicFDAL.Exists(claveFDAC) Then

                    datosFD = dicFDAL.Item(claveFDAC)
                    valorAQ = datosFD(0)

                Else

                    'No se encontró la clave en FD A:L:
                    'AQ toma 0 y se registra el faltante.
                    valorAQ = 0

                    RegistrarFaltanteFD _
                        dicFaltantesFD, _
                        "FD A:L", _
                        claveFDAC, _
                        FILA_INICIO + i - 1, _
                        "AQ"

                End If

            End If

        Else

            valorAM = CVErr(xlErrNA)
            valorAN = CVErr(xlErrNA)
            valorAP = CVErr(xlErrNA)
            valorAQ = CVErr(xlErrNA)

        End If

        salidaAGAX(i, 7) = valorAM
        salidaAGAX(i, 8) = valorAN
        salidaAGAX(i, 9) = 0

        salidaAGAX(i, 10) = valorAP
        salidaAGAX(i, 11) = valorAQ
        salidaAGAX(i, 12) = 0

        valorAE = salidaABAF(i, 4)
        valorAF = salidaABAF(i, 5)

        valorAS = CalcularCostoPonderado( _
            valorAG, valorAH, valorAM, valorAN, valorAE)

        valorAT = CalcularCostoPonderado( _
            valorAG, valorAH, valorAP, valorAQ, valorAF)

        salidaAGAX(i, 13) = valorAS
        salidaAGAX(i, 14) = valorAT

    Next i

    '======================================================
    ' AU, AV, AW, AX Y AZ
    '======================================================

    Application.StatusBar = "Calculando AU hasta AZ..."

    For i = LBound(clavesGrupos) To UBound(clavesGrupos)

        Set filasGrupo = dicGrupos.Item(clavesGrupos(i))

        cantidadGrupo = filasGrupo.Count
        filaIndice = CLng(filasGrupo(1))

        claveP = _
            NormalizarValor(datosDestino(filaIndice, 16))

        sumaIP = ObtenerSuma(dicSumaIPorP, claveP)
        sumaJP = ObtenerSuma(dicSumaJPorP, claveP)

        sumaABAE = 0
        cantidadABAE = 0
        errorAEGrupo = False

        sumaADAF = 0
        cantidadADAF = 0
        errorAFGrupo = False

        For j = 1 To cantidadGrupo

            filaIndice = CLng(filasGrupo(j))

            valorAE = salidaABAF(filaIndice, 4)
            valorAF = salidaABAF(filaIndice, 5)

            If IsError(valorAE) Then

                errorAEGrupo = True

            ElseIf EsNumeroValido(valorAE) Then

                If CDbl(valorAE) <> 0 Then

                    vAB = salidaABAF(filaIndice, 1)

                    If EsNumeroValido(vAB) Then

                        sumaABAE = sumaABAE + CDbl(vAB)
                        cantidadABAE = cantidadABAE + 1

                    End If

                End If

            End If

            If IsError(valorAF) Then

                errorAFGrupo = True

            ElseIf EsNumeroValido(valorAF) Then

                If CDbl(valorAF) <> 0 Then

                    vAD = salidaABAF(filaIndice, 3)

                    If EsNumeroValido(vAD) Then

                        sumaADAF = sumaADAF + CDbl(vAD)
                        cantidadADAF = cantidadADAF + 1

                    End If

                End If

            End If

        Next j

        If cantidadABAE > 0 And Not errorAEGrupo Then
            promedioABAE = sumaABAE / cantidadABAE
        Else
            promedioABAE = 0
        End If

        If cantidadADAF > 0 And Not errorAFGrupo Then
            promedioADAF = sumaADAF / cantidadADAF
        Else
            promedioADAF = 0
        End If

        '--------------------------------------------------
        ' UMBRALES DE SUBASTAS
        '--------------------------------------------------

        filaIndice = CLng(filasGrupo(1))

        claveAuxiliar = NormalizarValor( _
            TextoSeguro(datosDestino(filaIndice, 7)) & _
            "&" & _
            TextoSeguro(datosDestino(filaIndice, 16)))

        umbralesValidos = False

        If dicUmbralesSubastas.Exists(claveAuxiliar) Then

            datosUmbral = _
                dicUmbralesSubastas.Item(claveAuxiliar)

            If Not IsError(datosUmbral(0)) And _
               Not IsError(datosUmbral(1)) Then

                umbralSubida = NumeroSeguro(datosUmbral(0))
                umbralBajada = NumeroSeguro(datosUmbral(1))

                umbralesValidos = True

            End If

        End If

        '--------------------------------------------------
        ' PROMEDIOS PARA AW
        '--------------------------------------------------

        sumaABW = 0
        cantidadABW = 0

        sumaADW = 0
        cantidadADW = 0

        If umbralesValidos Then

            For j = 1 To cantidadGrupo

                filaIndice = CLng(filasGrupo(j))

                If valoresW(filaIndice) <= _
                   umbralBajada * 4# Then

                    vAB = salidaABAF(filaIndice, 1)

                    If EsNumeroValido(vAB) Then

                        sumaABW = sumaABW + CDbl(vAB)
                        cantidadABW = cantidadABW + 1

                    End If

                End If

                If valoresW(filaIndice) <= _
                   umbralSubida * 4# Then

                    vAD = salidaABAF(filaIndice, 3)

                    If EsNumeroValido(vAD) Then

                        sumaADW = sumaADW + CDbl(vAD)
                        cantidadADW = cantidadADW + 1

                    End If

                End If

            Next j

        End If

        If cantidadABW > 0 Then
            promedioABW = sumaABW / cantidadABW
        Else
            promedioABW = 0
        End If

        If cantidadADW > 0 Then
            promedioADW = sumaADW / cantidadADW
        Else
            promedioADW = 0
        End If

        sumaAXGrupo = 0
        sumaUGrupo = 0

        '--------------------------------------------------
        ' AU, AV, AW Y AX
        '--------------------------------------------------

        For j = 1 To cantidadGrupo

            filaIndice = CLng(filasGrupo(j))

            valorAE = salidaABAF(filaIndice, 4)
            valorAF = salidaABAF(filaIndice, 5)

            valorAS = salidaAGAX(filaIndice, 13)
            valorAT = salidaAGAX(filaIndice, 14)

            'AU
            valorAU = 0

            If cantidadABAE > 0 And _
               Not errorAEGrupo And _
               sumaIP > 10 And _
               EsNumeroValido(valorAE) Then

                valorAU = _
                    promedioABAE * CDbl(valorAE)

            End If

            'AV
            valorAV = 0

            If cantidadADAF > 0 And _
               Not errorAFGrupo And _
               sumaJP < -10 And _
               EsNumeroValido(valorAF) Then

                valorAV = _
                    promedioADAF * CDbl(valorAF)

            End If

            'AW
            valorAW = 0

            If umbralesValidos And _
               cantidadABW > 0 And _
               cantidadADW > 0 And _
               EsNumeroValido(valorAE) And _
               EsNumeroValido(valorAF) And _
               EsNumeroValido(valorAS) And _
               EsNumeroValido(valorAT) Then

                valorAW = _
                    (CDbl(valorAE) - CDbl(valorAS)) * _
                    promedioABW - _
                    (CDbl(valorAF) - CDbl(valorAT)) * _
                    promedioADW

            End If

            'AX
            valorAX = valorAU + valorAV - valorAW

            salidaAGAX(filaIndice, 15) = valorAU
            salidaAGAX(filaIndice, 16) = valorAV
            salidaAGAX(filaIndice, 17) = valorAW
            salidaAGAX(filaIndice, 18) = valorAX

            sumaAXGrupo = sumaAXGrupo + valorAX

            sumaUGrupo = sumaUGrupo + _
                NumeroSeguro(salidaRSTU(filaIndice, 4))

        Next j

        '--------------------------------------------------
        ' AZ
        '--------------------------------------------------

        valorAZGrupo = _
            (sumaAXGrupo - sumaUGrupo) / cantidadGrupo

        If valorAZGrupo < 0 Then
            valorAZGrupo = 0
        End If

        For j = 1 To cantidadGrupo

            filaIndice = CLng(filasGrupo(j))
            salidaAZ(filaIndice, 1) = valorAZGrupo

        Next j

    Next i

    '======================================================
    ' ESCRITURA MASIVA
    '======================================================

    Application.StatusBar = "Escribiendo resultados..."

    wsDestino.Range( _
        "L" & FILA_INICIO & ":O" & ultimaFila).Value2 = _
        salidaLMNO

    wsDestino.Range( _
        "R" & FILA_INICIO & ":U" & ultimaFila).Value2 = _
        salidaRSTU

    wsDestino.Range( _
        "W" & FILA_INICIO & ":Y" & ultimaFila).Value2 = _
        salidaWXY

    wsDestino.Range( _
        "AB" & FILA_INICIO & ":AF" & ultimaFila).Value2 = _
        salidaABAF

    wsDestino.Range( _
        "AG" & FILA_INICIO & ":AX" & ultimaFila).Value2 = _
        salidaAGAX

    wsDestino.Range( _
        "AZ" & FILA_INICIO & ":AZ" & ultimaFila).Value2 = _
        salidaAZ

    '======================================================
    ' REGISTRAR ÚLTIMA CORRIDA
    '======================================================

    duracionCalculo = SegundosTranscurridos(inicioTimer)

    RegistrarLogCalculo _
        wb:=wb, _
        wsDestino:=wsDestino, _
        cantidadFilas:=cantidadFilas, _
        inicioProceso:=inicioProceso, _
        duracionSegundos:=duracionCalculo, _
        ultimaSubastasDK:=ultimaSubastasDK, _
        ultimaSubastasUW:=ultimaSubastasUW, _
        ultimaProrrata:=ultimaProrrata, _
        ultimaDiccionario:=ultimaDiccionario, _
        ultimaResumenBC:=ultimaResumenBC, _
        ultimaFDAL:=ultimaFDAL, _
        ultimaFDQAD:=ultimaFDQAD

    procesoCorrecto = True

SalidaLimpia:

    If estadoGuardado Then

        Application.ScreenUpdating = pantallaAnterior
        Application.EnableEvents = eventosAnteriores
        Application.Calculation = calculoAnterior
        Application.StatusBar = barraAnterior

    End If

    If procesoCorrecto Then

        wsDestino.Activate

        mensajeFaltantesFD = _
            CrearMensajeFaltantesFD(dicFaltantesFD)

        If dicFaltantesFD.Count > 0 Then
            tipoMensajeFinal = vbExclamation
        Else
            tipoMensajeFinal = vbInformation
        End If

        MsgBox _
            "Proceso terminado correctamente." & _
            vbCrLf & vbCrLf & _
            "Columnas calculadas:" & vbCrLf & _
            "L:O, R:U, W:Y, AB:AX y AZ" & _
            vbCrLf & vbCrLf & _
            "Filas procesadas: " & _
            Format$(cantidadFilas, "#,##0") & _
            vbCrLf & _
            "Duración: " & _
            Format$(duracionCalculo, "0.000") & _
            " segundos" & _
            mensajeFaltantesFD & _
            vbCrLf & vbCrLf & _
            "El log contiene solamente la última corrida.", _
            tipoMensajeFinal, _
            "Cálculo E Costos"

    End If

    Exit Sub

ManejoError:

    numeroError = Err.Number
    descripcionError = Err.Description

    If wsDestino Is Nothing Then
        nombreHojaError = HOJA_DESTINO
    Else
        nombreHojaError = wsDestino.Name
    End If

    duracionCalculo = SegundosTranscurridos(inicioTimer)

    On Error Resume Next

    If Not wb Is Nothing Then

        RegistrarLogError _
            wb:=wb, _
            nombreHoja:=nombreHojaError, _
            inicioProceso:=inicioProceso, _
            duracionSegundos:=duracionCalculo, _
            numeroError:=numeroError, _
            descripcionError:=descripcionError

    End If

    On Error GoTo 0

    MsgBox _
        "Se produjo un error durante el proceso." & _
        vbCrLf & vbCrLf & _
        "Número: " & numeroError & vbCrLf & _
        "Descripción: " & descripcionError, _
        vbCritical, _
        "Cálculo E Costos"

    GoTo SalidaLimpia

End Sub


'==========================================================
' COMPARAR X ACTUAL CON X ANTERIOR
'
' Reproduce:
' =SI(X5=X4;W4+1;1)
'==========================================================

Private Function ValoresIgualesExcel( _
    ByVal valor1 As Variant, _
    ByVal valor2 As Variant) As Boolean

    If IsError(valor1) Or IsError(valor2) Then

        ValoresIgualesExcel = False
        Exit Function

    End If

    If IsEmpty(valor1) And IsEmpty(valor2) Then

        ValoresIgualesExcel = True
        Exit Function

    End If

    If VarType(valor1) = vbString And _
       VarType(valor2) = vbString Then

        ValoresIgualesExcel = _
            (StrComp(CStr(valor1), CStr(valor2), vbTextCompare) = 0)

        Exit Function

    End If

    If IsNumeric(valor1) And _
       IsNumeric(valor2) And _
       VarType(valor1) <> vbString And _
       VarType(valor2) <> vbString Then

        ValoresIgualesExcel = _
            (CDbl(valor1) = CDbl(valor2))

        Exit Function

    End If

    If VarType(valor1) = vbBoolean And _
       VarType(valor2) = vbBoolean Then

        ValoresIgualesExcel = _
            (CBool(valor1) = CBool(valor2))

        Exit Function

    End If

    ValoresIgualesExcel = False

End Function


'==========================================================
' AS Y AT
'==========================================================

Private Function CalcularCostoPonderado( _
    ByVal cantidad1 As Double, _
    ByVal cantidad2 As Double, _
    ByVal precio1 As Variant, _
    ByVal precio2 As Variant, _
    ByVal energia As Variant) As Variant

    Dim factor As Double

    If IsError(energia) Then

        CalcularCostoPonderado = energia
        Exit Function

    End If

    If Not IsNumeric(energia) Then

        CalcularCostoPonderado = CVErr(xlErrValue)
        Exit Function

    End If

    If cantidad1 + cantidad2 > 0 Then

        If IsError(precio1) Then

            CalcularCostoPonderado = precio1
            Exit Function

        End If

        If IsError(precio2) Then

            CalcularCostoPonderado = precio2
            Exit Function

        End If

        factor = _
            NumeroSeguro(precio1) * cantidad1 + _
            NumeroSeguro(precio2) * cantidad2

    Else

        factor = 1#

    End If

    CalcularCostoPonderado = _
        factor * CDbl(energia)

End Function


'==========================================================
' AE Y AF
'==========================================================

Private Function CalcularAsignacionEnergia( _
    ByVal bloque As Double, _
    ByVal energiaMaxima As Double, _
    ByVal factor As Double) As Double

    Dim cantidadBloques As Double
    Dim parteEntera As Double
    Dim fraccion As Double
    Dim proporcion As Double

    cantidadBloques = _
        4# * energiaMaxima / factor / 1000#

    parteEntera = Int(cantidadBloques)
    fraccion = cantidadBloques - parteEntera

    If bloque <= cantidadBloques Then

        proporcion = 1#

    ElseIf fraccion <> 0 And _
           bloque = parteEntera + 1# Then

        proporcion = fraccion

    Else

        proporcion = 0#

    End If

    CalcularAsignacionEnergia = _
        proporcion * factor / 4# * 1000#

End Function


Private Function CalcularBloque( _
    ByVal valor As Double) As Long

    CalcularBloque = _
        CLng(Int((valor - 1#) / 4#) + 1#)

End Function


'==========================================================
' DICCIONARIO SUBASTAS PARA L
'==========================================================

Private Function CrearDiccionarioSubastas( _
    ByVal datos As Variant) As Object

    Dim dic As Object
    Dim tipo As String
    Dim clave As String
    Dim i As Long

    Set dic = CreateObject("Scripting.Dictionary")
    dic.CompareMode = vbTextCompare

    For i = 1 To UBound(datos, 1)

        tipo = UCase$(Trim$(TextoSeguro(datos(i, 1))))

        If tipo = "BAJADA" Or tipo = "SUBIDA" Then

            clave = CrearClave4( _
                datos(i, 8), _
                datos(i, 4), _
                datos(i, 5), _
                datos(i, 6))

            If Not dic.Exists(clave) Then
                dic.Add clave, True
            End If

        End If

    Next i

    Set CrearDiccionarioSubastas = dic

End Function


'==========================================================
' DICCIONARIO SUBASTAS U:W
'==========================================================

Private Function CrearDiccionarioUmbralesSubastas( _
    ByVal datos As Variant) As Object

    Dim dic As Object
    Dim clave As String
    Dim i As Long

    Set dic = CreateObject("Scripting.Dictionary")
    dic.CompareMode = vbTextCompare

    For i = 1 To UBound(datos, 1)

        clave = NormalizarValor(datos(i, 1))

        If Len(clave) > 0 Then

            If Not dic.Exists(clave) Then

                dic.Add clave, Array( _
                    datos(i, 2), _
                    datos(i, 3))

            End If

        End If

    Next i

    Set CrearDiccionarioUmbralesSubastas = dic

End Function


'==========================================================
' DICCIONARIO PRORRATA SSCC
'==========================================================

Private Function CrearDiccionarioProrrata( _
    ByVal datos As Variant) As Object

    Dim dic As Object
    Dim clave As String
    Dim valores As Variant
    Dim i As Long

    Set dic = CreateObject("Scripting.Dictionary")
    dic.CompareMode = vbTextCompare

    For i = 1 To UBound(datos, 1)

        clave = CrearClave2( _
            datos(i, 1), _
            datos(i, 2))

        If dic.Exists(clave) Then

            valores = dic.Item(clave)

            valores(0) = _
                CDbl(valores(0)) + NumeroSeguro(datos(i, 3))

            valores(1) = _
                CDbl(valores(1)) + NumeroSeguro(datos(i, 4))

            dic.Item(clave) = valores

        Else

            dic.Add clave, Array( _
                NumeroSeguro(datos(i, 3)), _
                NumeroSeguro(datos(i, 4)))

        End If

    Next i

    Set CrearDiccionarioProrrata = dic

End Function


'==========================================================
' DICCIONARIO DE PRIMERA COINCIDENCIA
'==========================================================

Private Function CrearDiccionarioPrimerValor( _
    ByVal datos As Variant, _
    ByVal columnaClave As Long, _
    ByVal columnaValor As Long) As Object

    Dim dic As Object
    Dim clave As String
    Dim i As Long

    Set dic = CreateObject("Scripting.Dictionary")
    dic.CompareMode = vbTextCompare

    For i = 1 To UBound(datos, 1)

        clave = NormalizarValor(datos(i, columnaClave))

        If Len(clave) > 0 Then

            If Not dic.Exists(clave) Then

                dic.Add _
                    clave, _
                    datos(i, columnaValor)

            End If

        End If

    Next i

    Set CrearDiccionarioPrimerValor = dic

End Function


'==========================================================
' FD A:L
'
' Clave = A
' Valor 0 = K
' Valor 1 = L
'==========================================================

Private Function CrearDiccionarioFDAL( _
    ByVal datos As Variant) As Object

    Dim dic As Object
    Dim clave As String
    Dim i As Long

    Set dic = CreateObject("Scripting.Dictionary")
    dic.CompareMode = vbTextCompare

    For i = 1 To UBound(datos, 1)

        clave = NormalizarValor(datos(i, 1))

        If Len(clave) > 0 Then

            If Not dic.Exists(clave) Then

                dic.Add clave, Array( _
                    datos(i, 11), _
                    datos(i, 12))

            End If

        End If

    Next i

    Set CrearDiccionarioFDAL = dic

End Function


'==========================================================
' FD Q:AD
'
' Clave = Q
' Valor 0 = AC
' Valor 1 = AD
'==========================================================

Private Function CrearDiccionarioFDQAD( _
    ByVal datos As Variant) As Object

    Dim dic As Object
    Dim clave As String
    Dim i As Long

    Set dic = CreateObject("Scripting.Dictionary")
    dic.CompareMode = vbTextCompare

    For i = 1 To UBound(datos, 1)

        clave = NormalizarValor(datos(i, 1))

        If Len(clave) > 0 Then

            If Not dic.Exists(clave) Then

                dic.Add clave, Array( _
                    datos(i, 13), _
                    datos(i, 14))

            End If

        End If

    Next i

    Set CrearDiccionarioFDQAD = dic

End Function


'==========================================================
' REGISTRO DE CLAVES NO ENCONTRADAS EN FD
'
' Se registra cada combinación:
' sector FD + clave + fila destino + columnas destino.
' El cálculo continúa utilizando 0.
'==========================================================

Private Sub RegistrarFaltanteFD( _
    ByVal dic As Object, _
    ByVal sectorFD As String, _
    ByVal clave As String, _
    ByVal filaDestino As Long, _
    ByVal columnasDestino As String)

    Dim claveRegistro As String

    claveRegistro = _
        NormalizarValor(sectorFD) & SEP & _
        NormalizarValor(clave) & SEP & _
        CStr(filaDestino) & SEP & _
        NormalizarValor(columnasDestino)

    If Not dic.Exists(claveRegistro) Then

        dic.Add claveRegistro, Array( _
            sectorFD, _
            clave, _
            filaDestino, _
            columnasDestino)

    End If

End Sub


'==========================================================
' MENSAJE FINAL DE FALTANTES FD
'
' Para evitar un MsgBox excesivamente grande, muestra
' hasta 15 registros y luego informa cuántos adicionales
' existen. Todos los faltantes ya fueron tratados como 0.
'==========================================================

Private Function CrearMensajeFaltantesFD( _
    ByVal dic As Object) As String

    Const MAXIMO_MOSTRAR As Long = 15

    Dim claves As Variant
    Dim datos As Variant
    Dim mensaje As String

    Dim i As Long
    Dim cantidadMostrar As Long

    If dic Is Nothing Then
        CrearMensajeFaltantesFD = vbNullString
        Exit Function
    End If

    If dic.Count = 0 Then
        CrearMensajeFaltantesFD = vbNullString
        Exit Function
    End If

    mensaje = _
        vbCrLf & vbCrLf & _
        "ADVERTENCIA - DATOS NO ENCONTRADOS EN FD" & _
        vbCrLf & _
        "Los siguientes casos se reemplazaron por 0:" & _
        vbCrLf

    claves = dic.Keys

    If dic.Count < MAXIMO_MOSTRAR Then
        cantidadMostrar = dic.Count
    Else
        cantidadMostrar = MAXIMO_MOSTRAR
    End If

    For i = 0 To cantidadMostrar - 1

        datos = dic.Item(claves(i))

        mensaje = _
            mensaje & vbCrLf & _
            "Fila " & CStr(datos(2)) & _
            " | " & CStr(datos(0)) & _
            " | Clave: " & CStr(datos(1)) & _
            " | Destino: " & CStr(datos(3))

    Next i

    If dic.Count > MAXIMO_MOSTRAR Then

        mensaje = _
            mensaje & _
            vbCrLf & _
            "... y " & _
            Format$(dic.Count - MAXIMO_MOSTRAR, "#,##0") & _
            " registro(s) adicional(es)."

    End If

    mensaje = _
        mensaje & _
        vbCrLf & vbCrLf & _
        "Total de registros sin coincidencia en FD: " & _
        Format$(dic.Count, "#,##0")

    CrearMensajeFaltantesFD = mensaje

End Function


'==========================================================
' SUMAS EN DICCIONARIOS
'==========================================================

Private Sub AgregarSuma( _
    ByVal dic As Object, _
    ByVal clave As String, _
    ByVal valor As Double)

    If dic.Exists(clave) Then

        dic.Item(clave) = _
            CDbl(dic.Item(clave)) + valor

    Else

        dic.Add clave, valor

    End If

End Sub


Private Function ObtenerSuma( _
    ByVal dic As Object, _
    ByVal clave As String) As Double

    If dic.Exists(clave) Then
        ObtenerSuma = CDbl(dic.Item(clave))
    Else
        ObtenerSuma = 0
    End If

End Function


'==========================================================
' ORDENAR POR F DESCENDENTE
'==========================================================

Private Sub QuickSortFDesc( _
    ByRef indices() As Long, _
    ByVal primero As Long, _
    ByVal ultimo As Long, _
    ByRef valoresF() As Double)

    Dim izq As Long
    Dim der As Long
    Dim temp As Long
    Dim pivote As Double

    izq = primero
    der = ultimo

    pivote = _
        valoresF(indices((primero + ultimo) \ 2))

    Do While izq <= der

        Do While izq <= ultimo

            If valoresF(indices(izq)) > pivote Then
                izq = izq + 1
            Else
                Exit Do
            End If

        Loop

        Do While der >= primero

            If valoresF(indices(der)) < pivote Then
                der = der - 1
            Else
                Exit Do
            End If

        Loop

        If izq <= der Then

            temp = indices(izq)
            indices(izq) = indices(der)
            indices(der) = temp

            izq = izq + 1
            der = der - 1

        End If

    Loop

    If primero < der Then
        QuickSortFDesc indices, primero, der, valoresF
    End If

    If izq < ultimo Then
        QuickSortFDesc indices, izq, ultimo, valoresF
    End If

End Sub


'==========================================================
' ORDENAR POR Q DESCENDENTE Y F DESCENDENTE
'==========================================================

Private Sub QuickSortQFDesc( _
    ByRef indices() As Long, _
    ByVal primero As Long, _
    ByVal ultimo As Long, _
    ByRef valoresQ() As Double, _
    ByRef valoresF() As Double)

    Dim izq As Long
    Dim der As Long
    Dim temp As Long

    Dim indicePivote As Long
    Dim qPivote As Double
    Dim fPivote As Double

    izq = primero
    der = ultimo

    indicePivote = _
        indices((primero + ultimo) \ 2)

    qPivote = valoresQ(indicePivote)
    fPivote = valoresF(indicePivote)

    Do While izq <= der

        Do While izq <= ultimo

            If EsMayorQF( _
                indices(izq), _
                qPivote, _
                fPivote, _
                valoresQ, _
                valoresF) Then

                izq = izq + 1

            Else

                Exit Do

            End If

        Loop

        Do While der >= primero

            If EsMenorQF( _
                indices(der), _
                qPivote, _
                fPivote, _
                valoresQ, _
                valoresF) Then

                der = der - 1

            Else

                Exit Do

            End If

        Loop

        If izq <= der Then

            temp = indices(izq)
            indices(izq) = indices(der)
            indices(der) = temp

            izq = izq + 1
            der = der - 1

        End If

    Loop

    If primero < der Then

        QuickSortQFDesc _
            indices, primero, der, valoresQ, valoresF

    End If

    If izq < ultimo Then

        QuickSortQFDesc _
            indices, izq, ultimo, valoresQ, valoresF

    End If

End Sub


Private Function EsMayorQF( _
    ByVal indice As Long, _
    ByVal qPivote As Double, _
    ByVal fPivote As Double, _
    ByRef valoresQ() As Double, _
    ByRef valoresF() As Double) As Boolean

    EsMayorQF = _
        valoresQ(indice) > qPivote Or _
        (valoresQ(indice) = qPivote And _
         valoresF(indice) > fPivote)

End Function


Private Function EsMenorQF( _
    ByVal indice As Long, _
    ByVal qPivote As Double, _
    ByVal fPivote As Double, _
    ByRef valoresQ() As Double, _
    ByRef valoresF() As Double) As Boolean

    EsMenorQF = _
        valoresQ(indice) < qPivote Or _
        (valoresQ(indice) = qPivote And _
         valoresF(indice) < fPivote)

End Function


'==========================================================
' ORDENAR Q DESCENDENTE / ASCENDENTE
' CON FILA ASCENDENTE
'==========================================================

Private Sub QuickSortQDescFilaAsc( _
    ByRef indices() As Long, _
    ByVal primero As Long, _
    ByVal ultimo As Long, _
    ByRef valoresQ() As Double)

    QuickSortQFila _
        indices, primero, ultimo, valoresQ, True

End Sub


Private Sub QuickSortQAscFilaAsc( _
    ByRef indices() As Long, _
    ByVal primero As Long, _
    ByVal ultimo As Long, _
    ByRef valoresQ() As Double)

    QuickSortQFila _
        indices, primero, ultimo, valoresQ, False

End Sub


Private Sub QuickSortQFila( _
    ByRef indices() As Long, _
    ByVal primero As Long, _
    ByVal ultimo As Long, _
    ByRef valoresQ() As Double, _
    ByVal descendente As Boolean)

    Dim izq As Long
    Dim der As Long
    Dim temp As Long
    Dim indicePivote As Long

    izq = primero
    der = ultimo

    indicePivote = _
        indices((primero + ultimo) \ 2)

    Do While izq <= der

        Do While izq <= ultimo

            If CompararQFila( _
                indices(izq), _
                indicePivote, _
                valoresQ, _
                descendente) Then

                izq = izq + 1

            Else

                Exit Do

            End If

        Loop

        Do While der >= primero

            If CompararQFila( _
                indicePivote, _
                indices(der), _
                valoresQ, _
                descendente) Then

                der = der - 1

            Else

                Exit Do

            End If

        Loop

        If izq <= der Then

            temp = indices(izq)
            indices(izq) = indices(der)
            indices(der) = temp

            izq = izq + 1
            der = der - 1

        End If

    Loop

    If primero < der Then

        QuickSortQFila _
            indices, primero, der, valoresQ, descendente

    End If

    If izq < ultimo Then

        QuickSortQFila _
            indices, izq, ultimo, valoresQ, descendente

    End If

End Sub


Private Function CompararQFila( _
    ByVal indiceA As Long, _
    ByVal indiceB As Long, _
    ByRef valoresQ() As Double, _
    ByVal descendente As Boolean) As Boolean

    If descendente Then

        CompararQFila = _
            valoresQ(indiceA) > valoresQ(indiceB) Or _
            (valoresQ(indiceA) = valoresQ(indiceB) And _
             indiceA < indiceB)

    Else

        CompararQFila = _
            valoresQ(indiceA) < valoresQ(indiceB) Or _
            (valoresQ(indiceA) = valoresQ(indiceB) And _
             indiceA < indiceB)

    End If

End Function


'==========================================================
' LOG DE EJECUCIÓN CORRECTA
'
' Borra todos los registros anteriores y deja solamente
' los 35 registros correspondientes a la última corrida.
'==========================================================

Private Sub RegistrarLogCalculo( _
    ByVal wb As Workbook, _
    ByVal wsDestino As Worksheet, _
    ByVal cantidadFilas As Long, _
    ByVal inicioProceso As Date, _
    ByVal duracionSegundos As Double, _
    ByVal ultimaSubastasDK As Long, _
    ByVal ultimaSubastasUW As Long, _
    ByVal ultimaProrrata As Long, _
    ByVal ultimaDiccionario As Long, _
    ByVal ultimaResumenBC As Long, _
    ByVal ultimaFDAL As Long, _
    ByVal ultimaFDQAD As Long)

    Const CANTIDAD_REGISTROS As Long = 35

    Dim wsLog As Worksheet
    Dim detalle() As Variant

    Dim filaDetalle As Long
    Dim idEjecucion As String
    Dim ultimaFilaDestino As Long

    Set wsLog = ObtenerHojaLog(wb)

    PrepararHojaLog wsLog
    LimpiarLogAnterior wsLog

    idEjecucion = _
        Format$(inicioProceso, "yyyymmdd_hhnnss")

    ultimaFilaDestino = _
        FILA_INICIO + cantidadFilas - 1

    ReDim detalle(1 To CANTIDAD_REGISTROS, 1 To 10)

    filaDetalle = 0

    AgregarDetalleLog detalle, filaDetalle, _
        idEjecucion, inicioProceso, wsDestino.Name, "L", _
        "Marca con 1 si la combinación G-A-B-C de la fila existe en Subastas como registro BAJADA o SUBIDA; si no existe, deja 0.", _
        wsDestino.Name & "!G4:G" & ultimaFilaDestino & ", A4:C" & ultimaFilaDestino & "; " & _
        HOJA_SUBASTAS & "!D1:K" & ultimaSubastasDK & " (D=tipo; G,H,I,K forman la clave equivalente).", _
        "Crea la clave G|A|B|C; busca K|G|H|I solo en filas con D=BAJADA o SUBIDA. Destino: " & _
        wsDestino.Name & "!L4:L" & ultimaFilaDestino & ".", _
        cantidadFilas, duracionSegundos

    AgregarDetalleLog detalle, filaDetalle, _
        idEjecucion, inicioProceso, wsDestino.Name, "M", _
        "Compara K de cada fila con el umbral único de Resumen!H8: deja 1 cuando K>H8 y 0 en caso contrario.", _
        wsDestino.Name & "!K4:K" & ultimaFilaDestino & "; " & _
        HOJA_RESUMEN & "!H8.", _
        "Lectura del umbral una sola vez y comparación fila a fila en memoria. Destino: " & _
        wsDestino.Name & "!M4:M" & ultimaFilaDestino & ".", _
        cantidadFilas, duracionSegundos

    AgregarDetalleLog detalle, filaDetalle, _
        idEjecucion, inicioProceso, wsDestino.Name, "N", _
        "Para cada grupo con igual G y P, suma I de las filas con L=1 cuyo F es mayor o igual al F de la fila evaluada.", _
        wsDestino.Name & "!G4:G" & ultimaFilaDestino & ", P4:P" & ultimaFilaDestino & ", " & _
        "F4:F" & ultimaFilaDestino & ", I4:I" & ultimaFilaDestino & " y L4:L" & ultimaFilaDestino & " calculada previamente.", _
        "Agrupa por G-P, ordena F de mayor a menor y acumula I por bloques de igual F; los empates reciben el mismo acumulado. Destino: " & _
        wsDestino.Name & "!N4:N" & ultimaFilaDestino & ".", _
        cantidadFilas, duracionSegundos

    AgregarDetalleLog detalle, filaDetalle, _
        idEjecucion, inicioProceso, wsDestino.Name, "O", _
        "Para cada grupo con igual G y P, calcula el negativo de la suma J de las filas con L=1 cuyo F es mayor o igual al F evaluado.", _
        wsDestino.Name & "!G4:G" & ultimaFilaDestino & ", P4:P" & ultimaFilaDestino & ", " & _
        "F4:F" & ultimaFilaDestino & ", J4:J" & ultimaFilaDestino & " y L4:L" & ultimaFilaDestino & " calculada previamente.", _
        "Agrupa por G-P, ordena F de mayor a menor, acumula J por bloques de igual F y cambia el signo del acumulado. Destino: " & _
        wsDestino.Name & "!O4:O" & ultimaFilaDestino & ".", _
        cantidadFilas, duracionSegundos

    AgregarDetalleLog detalle, filaDetalle, _
        idEjecucion, inicioProceso, wsDestino.Name, "R", _
        "Asigna una posición dentro de cada grupo G-P, ordenando primero Q descendente y luego F descendente; iguales Q y F comparten posición.", _
        wsDestino.Name & "!G4:G" & ultimaFilaDestino & ", P4:P" & ultimaFilaDestino & ", " & _
        "Q4:Q" & ultimaFilaDestino & " y F4:F" & ultimaFilaDestino & ".", _
        "Ordenamiento en memoria por Q? y F?; la posición corresponde al inicio del bloque empatado. Destino: " & _
        wsDestino.Name & "!R4:R" & ultimaFilaDestino & ".", _
        cantidadFilas, duracionSegundos

    AgregarDetalleLog detalle, filaDetalle, _
        idEjecucion, inicioProceso, wsDestino.Name, "S", _
        "Multiplica el valor I de cada fila por su valor Q.", _
        wsDestino.Name & "!I4:I" & ultimaFilaDestino & " y Q4:Q" & ultimaFilaDestino & ".", _
        "Operación fila a fila: S=I*Q. Destino: " & wsDestino.Name & "!S4:S" & ultimaFilaDestino & ".", _
        cantidadFilas, duracionSegundos

    AgregarDetalleLog detalle, filaDetalle, _
        idEjecucion, inicioProceso, wsDestino.Name, "T", _
        "Multiplica el valor J de cada fila por su valor Q.", _
        wsDestino.Name & "!J4:J" & ultimaFilaDestino & " y Q4:Q" & ultimaFilaDestino & ".", _
        "Operación fila a fila: T=J*Q. Destino: " & wsDestino.Name & "!T4:T" & ultimaFilaDestino & ".", _
        cantidadFilas, duracionSegundos

    AgregarDetalleLog detalle, filaDetalle, _
        idEjecucion, inicioProceso, wsDestino.Name, "U", _
        "Calcula únicamente el costo de filas reconocidas en L: multiplica L por la suma de S y T.", _
        wsDestino.Name & "!L4:L" & ultimaFilaDestino & ", S4:S" & ultimaFilaDestino & " y T4:T" & ultimaFilaDestino & _
        "; S y T provienen de I, J y Q.", _
        "Operación equivalente a U=L*(S+T)=L*((I*Q)+(J*Q)). Destino: " & _
        wsDestino.Name & "!U4:U" & ultimaFilaDestino & ".", _
        cantidadFilas, duracionSegundos

    AgregarDetalleLog detalle, filaDetalle, _
        idEjecucion, inicioProceso, wsDestino.Name, "W", _
        "Genera un contador secuencial por bloques consecutivos de P: W4 toma C4; desde la fila 5 aumenta W anterior en 1 si P actual=P anterior y reinicia en 1 cuando cambia.", _
        wsDestino.Name & "!C4 y P4:P" & ultimaFilaDestino & "; utiliza además el valor W calculado para la fila anterior.", _
        "Recorrido estrictamente secuencial para respetar la dependencia entre filas. Destino: " & _
        wsDestino.Name & "!W4:W" & ultimaFilaDestino & ".", _
        cantidadFilas, duracionSegundos

    AgregarDetalleLog detalle, filaDetalle, _
        idEjecucion, inicioProceso, wsDestino.Name, "X", _
        "Copia exactamente P en X, sin transformar el valor.", _
        wsDestino.Name & "!P4:P" & ultimaFilaDestino & ".", _
        "Asignación directa X=P. Destino: " & wsDestino.Name & "!X4:X" & ultimaFilaDestino & ".", _
        cantidadFilas, duracionSegundos

    AgregarDetalleLog detalle, filaDetalle, _
        idEjecucion, inicioProceso, wsDestino.Name, "Y", _
        "Dentro de cada grupo G-P, coloca primero los F de filas con L=1 e I distinto de 0, ordenados por Q descendente; después coloca los F restantes.", _
        wsDestino.Name & "!G4:G" & ultimaFilaDestino & ", P4:P" & ultimaFilaDestino & ", " & _
        "L4:L" & ultimaFilaDestino & ", I4:I" & ultimaFilaDestino & ", Q4:Q" & ultimaFilaDestino & " y F4:F" & ultimaFilaDestino & ".", _
        "Separa filas calificadas/no calificadas; ordena las calificadas por Q? y fila original?, y distribuye los F sobre las posiciones del grupo. Destino: " & _
        wsDestino.Name & "!Y4:Y" & ultimaFilaDestino & ".", _
        cantidadFilas, duracionSegundos

    AgregarDetalleLog detalle, filaDetalle, _
        idEjecucion, inicioProceso, wsDestino.Name, "AB", _
        "Dentro de cada grupo G-P, coloca los Q de filas con L=1 e I distinto de 0, ordenados de mayor a menor; las posiciones restantes quedan vacías.", _
        wsDestino.Name & "!G4:G" & ultimaFilaDestino & ", P4:P" & ultimaFilaDestino & ", " & _
        "L4:L" & ultimaFilaDestino & ", I4:I" & ultimaFilaDestino & " y Q4:Q" & ultimaFilaDestino & ".", _
        "Usa el mismo orden aplicado para Y (Q? y fila original?). Destino: " & _
        wsDestino.Name & "!AB4:AB" & ultimaFilaDestino & ".", _
        cantidadFilas, duracionSegundos

    AgregarDetalleLog detalle, filaDetalle, _
        idEjecucion, inicioProceso, wsDestino.Name, "AC", _
        "Dentro de cada grupo G-P, coloca primero los F de filas con L=1 y J distinto de 0, ordenados por Q ascendente; después coloca los F restantes.", _
        wsDestino.Name & "!G4:G" & ultimaFilaDestino & ", P4:P" & ultimaFilaDestino & ", " & _
        "L4:L" & ultimaFilaDestino & ", J4:J" & ultimaFilaDestino & ", Q4:Q" & ultimaFilaDestino & " y F4:F" & ultimaFilaDestino & ".", _
        "Separa filas calificadas/no calificadas; ordena las calificadas por Q? y fila original?, y distribuye los F sobre las posiciones del grupo. Destino: " & _
        wsDestino.Name & "!AC4:AC" & ultimaFilaDestino & ".", _
        cantidadFilas, duracionSegundos

    AgregarDetalleLog detalle, filaDetalle, _
        idEjecucion, inicioProceso, wsDestino.Name, "AD", _
        "Dentro de cada grupo G-P, coloca los Q de filas con L=1 y J distinto de 0, ordenados de menor a mayor; las posiciones restantes quedan vacías.", _
        wsDestino.Name & "!G4:G" & ultimaFilaDestino & ", P4:P" & ultimaFilaDestino & ", " & _
        "L4:L" & ultimaFilaDestino & ", J4:J" & ultimaFilaDestino & " y Q4:Q" & ultimaFilaDestino & ".", _
        "Usa el mismo orden aplicado para AC (Q? y fila original?). Destino: " & _
        wsDestino.Name & "!AD4:AD" & ultimaFilaDestino & ".", _
        cantidadFilas, duracionSegundos

    AgregarDetalleLog detalle, filaDetalle, _
        idEjecucion, inicioProceso, wsDestino.Name, "AE", _
        "Distribuye el máximo N de cada grupo G-P según el número de bloque W y el factor asociado a G en Resumen B:C.", _
        wsDestino.Name & "!G4:G" & ultimaFilaDestino & ", P4:P" & ultimaFilaDestino & ", N4:N" & ultimaFilaDestino & _
        " y W4:W" & ultimaFilaDestino & "; " & HOJA_RESUMEN & "!B" & FILA_INICIO_RESUMEN & ":C" & ultimaResumenBC & " (B=G, C=factor).", _
        "Calcula bloques=4*MAX(N)/factor/1000; asigna factor/4*1000 a bloques completos y la proporción al bloque fraccionario. Destino: " & _
        wsDestino.Name & "!AE4:AE" & ultimaFilaDestino & ".", _
        cantidadFilas, duracionSegundos

    AgregarDetalleLog detalle, filaDetalle, _
        idEjecucion, inicioProceso, wsDestino.Name, "AF", _
        "Distribuye con signo negativo el máximo O de cada grupo G-P según W y el factor asociado a G en Resumen B:C.", _
        wsDestino.Name & "!G4:G" & ultimaFilaDestino & ", P4:P" & ultimaFilaDestino & ", O4:O" & ultimaFilaDestino & _
        " y W4:W" & ultimaFilaDestino & "; " & HOJA_RESUMEN & "!B" & FILA_INICIO_RESUMEN & ":C" & ultimaResumenBC & " (B=G, C=factor).", _
        "Aplica la misma asignación por bloques de AE y cambia el signo del resultado. Destino: " & _
        wsDestino.Name & "!AF4:AF" & ultimaFilaDestino & ".", _
        cantidadFilas, duracionSegundos

    AgregarDetalleLog detalle, filaDetalle, _
        idEjecucion, inicioProceso, wsDestino.Name, "AG", _
        "Suma la columna C de Prorrata SSCC para todas las filas donde A coincide con G y B coincide con D de la fila destino.", _
        wsDestino.Name & "!G4:G" & ultimaFilaDestino & " y D4:D" & ultimaFilaDestino & "; " & _
        HOJA_PRORRATA & "!A1:D" & ultimaProrrata & " (A=G, B=D, C=valor a sumar).", _
        "Diccionario de sumas por clave G-D; cuando no hay coincidencia deja 0. Destino: " & _
        wsDestino.Name & "!AG4:AG" & ultimaFilaDestino & ".", _
        cantidadFilas, duracionSegundos

    AgregarDetalleLog detalle, filaDetalle, _
        idEjecucion, inicioProceso, wsDestino.Name, "AH", _
        "Suma la columna D de Prorrata SSCC para todas las filas donde A coincide con G y B coincide con D de la fila destino.", _
        wsDestino.Name & "!G4:G" & ultimaFilaDestino & " y D4:D" & ultimaFilaDestino & "; " & _
        HOJA_PRORRATA & "!A1:D" & ultimaProrrata & " (A=G, B=D, D=valor a sumar).", _
        "Diccionario de sumas por clave G-D; cuando no hay coincidencia deja 0. Destino: " & _
        wsDestino.Name & "!AH4:AH" & ultimaFilaDestino & ".", _
        cantidadFilas, duracionSegundos

    AgregarDetalleLog detalle, filaDetalle, _
        idEjecucion, inicioProceso, wsDestino.Name, "AI", _
        "Asigna el valor constante 0 en todas las filas procesadas.", _
        "No utiliza rangos de origen.", _
        "Asignación directa. Destino: " & wsDestino.Name & "!AI4:AI" & ultimaFilaDestino & ".", _
        cantidadFilas, duracionSegundos

    AgregarDetalleLog detalle, filaDetalle, _
        idEjecucion, inicioProceso, wsDestino.Name, "AJ", _
        "Repite en AJ la misma suma calculada para AG: columna C de Prorrata SSCC por coincidencia G-D.", _
        wsDestino.Name & "!G4:G" & ultimaFilaDestino & " y D4:D" & ultimaFilaDestino & "; " & _
        HOJA_PRORRATA & "!A1:D" & ultimaProrrata & ".", _
        "Reutiliza en memoria el resultado AG de la misma fila. Destino: " & _
        wsDestino.Name & "!AJ4:AJ" & ultimaFilaDestino & ".", _
        cantidadFilas, duracionSegundos

    AgregarDetalleLog detalle, filaDetalle, _
        idEjecucion, inicioProceso, wsDestino.Name, "AK", _
        "Repite en AK la misma suma calculada para AH: columna D de Prorrata SSCC por coincidencia G-D.", _
        wsDestino.Name & "!G4:G" & ultimaFilaDestino & " y D4:D" & ultimaFilaDestino & "; " & _
        HOJA_PRORRATA & "!A1:D" & ultimaProrrata & ".", _
        "Reutiliza en memoria el resultado AH de la misma fila. Destino: " & _
        wsDestino.Name & "!AK4:AK" & ultimaFilaDestino & ".", _
        cantidadFilas, duracionSegundos

    AgregarDetalleLog detalle, filaDetalle, _
        idEjecucion, inicioProceso, wsDestino.Name, "AL", _
        "Asigna el valor constante 0 en todas las filas procesadas.", _
        "No utiliza rangos de origen.", _
        "Asignación directa. Destino: " & wsDestino.Name & "!AL4:AL" & ultimaFilaDestino & ".", _
        cantidadFilas, duracionSegundos

    AgregarDetalleLog detalle, filaDetalle, _
        idEjecucion, inicioProceso, wsDestino.Name, "AM", _
        "Obtiene el código equivalente de G en Diccionario A:B; calcula el bloque de Y y busca en FD la clave bloque+código para devolver FD columna AD.", _
        wsDestino.Name & "!G4:G" & ultimaFilaDestino & " y Y4:Y" & ultimaFilaDestino & "; " & _
        HOJA_DICCIONARIO & "!A1:B" & ultimaDiccionario & "; " & HOJA_FD & "!Q1:AD" & ultimaFDQAD & " (Q=clave, AD=resultado).", _
        "Bloque=ENTERO((Y-1)/4)+1; clave=texto del bloque concatenado con Diccionario!B. Primera coincidencia; si falta devuelve #N/A. Destino: " & _
        wsDestino.Name & "!AM4:AM" & ultimaFilaDestino & ".", _
        cantidadFilas, duracionSegundos

    AgregarDetalleLog detalle, filaDetalle, _
        idEjecucion, inicioProceso, wsDestino.Name, "AN", _
        "Obtiene el código equivalente de G en Diccionario A:B; calcula el bloque de Y y busca en FD la clave bloque+código para devolver FD columna L.", _
        wsDestino.Name & "!G4:G" & ultimaFilaDestino & " y Y4:Y" & ultimaFilaDestino & "; " & _
        HOJA_DICCIONARIO & "!A1:B" & ultimaDiccionario & "; " & HOJA_FD & "!A1:L" & ultimaFDAL & " (A=clave, L=resultado).", _
        "Bloque=ENTERO((Y-1)/4)+1; clave=texto del bloque concatenado con Diccionario!B. Primera coincidencia; si falta devuelve #N/A. Destino: " & _
        wsDestino.Name & "!AN4:AN" & ultimaFilaDestino & ".", _
        cantidadFilas, duracionSegundos

    AgregarDetalleLog detalle, filaDetalle, _
        idEjecucion, inicioProceso, wsDestino.Name, "AO", _
        "Asigna el valor constante 0 en todas las filas procesadas.", _
        "No utiliza rangos de origen.", _
        "Asignación directa. Destino: " & wsDestino.Name & "!AO4:AO" & ultimaFilaDestino & ".", _
        cantidadFilas, duracionSegundos

    AgregarDetalleLog detalle, filaDetalle, _
        idEjecucion, inicioProceso, wsDestino.Name, "AP", _
        "Obtiene el código equivalente de G en Diccionario A:B; calcula el bloque de Y y busca en FD la clave bloque+código para devolver FD columna AC.", _
        wsDestino.Name & "!G4:G" & ultimaFilaDestino & " y Y4:Y" & ultimaFilaDestino & "; " & _
        HOJA_DICCIONARIO & "!A1:B" & ultimaDiccionario & "; " & HOJA_FD & "!Q1:AD" & ultimaFDQAD & " (Q=clave, AC=resultado).", _
        "Bloque=ENTERO((Y-1)/4)+1; clave=texto del bloque concatenado con Diccionario!B. Primera coincidencia; si falta devuelve #N/A. Destino: " & _
        wsDestino.Name & "!AP4:AP" & ultimaFilaDestino & ".", _
        cantidadFilas, duracionSegundos

    AgregarDetalleLog detalle, filaDetalle, _
        idEjecucion, inicioProceso, wsDestino.Name, "AQ", _
        "Obtiene el código equivalente de G en Diccionario A:B; calcula el bloque de AC y busca en FD la clave bloque+código para devolver FD columna K.", _
        wsDestino.Name & "!G4:G" & ultimaFilaDestino & " y AC4:AC" & ultimaFilaDestino & "; " & _
        HOJA_DICCIONARIO & "!A1:B" & ultimaDiccionario & "; " & HOJA_FD & "!A1:L" & ultimaFDAL & " (A=clave, K=resultado).", _
        "Bloque=ENTERO((AC-1)/4)+1; clave=texto del bloque concatenado con Diccionario!B. Primera coincidencia; si falta devuelve #N/A. Destino: " & _
        wsDestino.Name & "!AQ4:AQ" & ultimaFilaDestino & ".", _
        cantidadFilas, duracionSegundos

    AgregarDetalleLog detalle, filaDetalle, _
        idEjecucion, inicioProceso, wsDestino.Name, "AR", _
        "Asigna el valor constante 0 en todas las filas procesadas.", _
        "No utiliza rangos de origen.", _
        "Asignación directa. Destino: " & wsDestino.Name & "!AR4:AR" & ultimaFilaDestino & ".", _
        cantidadFilas, duracionSegundos

    AgregarDetalleLog detalle, filaDetalle, _
        idEjecucion, inicioProceso, wsDestino.Name, "AS", _
        "Calcula AE por un factor formado con AG, AH, AM y AN. Si AG+AH es mayor que 0 usa AM*AG+AN*AH; si no, usa factor 1.", _
        wsDestino.Name & "!AG4:AH" & ultimaFilaDestino & ", AM4:AN" & ultimaFilaDestino & " y AE4:AE" & ultimaFilaDestino & ".", _
        "Por fila: AS=AE*(AM*AG+AN*AH) cuando AG+AH>0; de lo contrario AS=AE. Los errores de AE, AM o AN se propagan. Destino: " & _
        wsDestino.Name & "!AS4:AS" & ultimaFilaDestino & ".", _
        cantidadFilas, duracionSegundos

    AgregarDetalleLog detalle, filaDetalle, _
        idEjecucion, inicioProceso, wsDestino.Name, "AT", _
        "Calcula AF por un factor formado con AJ, AK, AP y AQ. Si AJ+AK es mayor que 0 usa AP*AJ+AQ*AK; si no, usa factor 1.", _
        wsDestino.Name & "!AJ4:AK" & ultimaFilaDestino & ", AP4:AQ" & ultimaFilaDestino & " y AF4:AF" & ultimaFilaDestino & ".", _
        "Por fila: AT=AF*(AP*AJ+AQ*AK) cuando AJ+AK>0; de lo contrario AT=AF. Los errores de AF, AP o AQ se propagan. Destino: " & _
        wsDestino.Name & "!AT4:AT" & ultimaFilaDestino & ".", _
        cantidadFilas, duracionSegundos

    AgregarDetalleLog detalle, filaDetalle, _
        idEjecucion, inicioProceso, wsDestino.Name, "AU", _
        "Para cada grupo G-P obtiene el promedio de AB solo en filas con AE numérico y distinto de 0; lo multiplica por AE si la suma total de I para ese P es mayor que 10.", _
        wsDestino.Name & "!G4:G" & ultimaFilaDestino & ", P4:P" & ultimaFilaDestino & ", AB4:AB" & ultimaFilaDestino & ", " & _
        "AE4:AE" & ultimaFilaDestino & " e I4:I" & ultimaFilaDestino & ".", _
        "El promedio se calcula por G-P, pero la condición SUMA(I)>10 se calcula por P considerando todas las G. Si algún AE del grupo es error, AU queda 0. Destino: " & _
        wsDestino.Name & "!AU4:AU" & ultimaFilaDestino & ".", _
        cantidadFilas, duracionSegundos

    AgregarDetalleLog detalle, filaDetalle, _
        idEjecucion, inicioProceso, wsDestino.Name, "AV", _
        "Para cada grupo G-P obtiene el promedio de AD solo en filas con AF numérico y distinto de 0; lo multiplica por AF si la suma total de J para ese P es menor que -10.", _
        wsDestino.Name & "!G4:G" & ultimaFilaDestino & ", P4:P" & ultimaFilaDestino & ", AD4:AD" & ultimaFilaDestino & ", " & _
        "AF4:AF" & ultimaFilaDestino & " y J4:J" & ultimaFilaDestino & ".", _
        "El promedio se calcula por G-P, pero la condición SUMA(J)<-10 se calcula por P considerando todas las G. Si algún AF del grupo es error, AV queda 0. Destino: " & _
        wsDestino.Name & "!AV4:AV" & ultimaFilaDestino & ".", _
        cantidadFilas, duracionSegundos

    AgregarDetalleLog detalle, filaDetalle, _
        idEjecucion, inicioProceso, wsDestino.Name, "AW", _
        "Calcula una diferencia valorizada usando promedios de AB y AD limitados por umbrales de subida y bajada definidos en Subastas U:W.", _
        wsDestino.Name & "!G4:G" & ultimaFilaDestino & ", P4:P" & ultimaFilaDestino & ", W4:W" & ultimaFilaDestino & ", " & _
        "AB4:AB" & ultimaFilaDestino & ", AD4:AD" & ultimaFilaDestino & ", AE4:AF" & ultimaFilaDestino & ", AS4:AT" & ultimaFilaDestino & "; " & _
        HOJA_SUBASTAS & "!U1:W" & ultimaSubastasUW & " (U=clave G&P, V=umbral subida, W=umbral bajada).", _
        "Por G-P promedia AB donde W<=umbralBajada*4 y AD donde W<=umbralSubida*4; luego AW=(AE-AS)*promAB-(AF-AT)*promAD. Destino: " & _
        wsDestino.Name & "!AW4:AW" & ultimaFilaDestino & ".", _
        cantidadFilas, duracionSegundos

    AgregarDetalleLog detalle, filaDetalle, _
        idEjecucion, inicioProceso, wsDestino.Name, "AX", _
        "Combina los tres componentes económicos calculados previamente.", _
        wsDestino.Name & "!AU4:AU" & ultimaFilaDestino & ", AV4:AV" & ultimaFilaDestino & " y AW4:AW" & ultimaFilaDestino & ".", _
        "Operación fila a fila: AX=AU+AV-AW. Destino: " & wsDestino.Name & "!AX4:AX" & ultimaFilaDestino & ".", _
        cantidadFilas, duracionSegundos

    AgregarDetalleLog detalle, filaDetalle, _
        idEjecucion, inicioProceso, wsDestino.Name, "AZ", _
        "Para cada grupo G-P calcula la diferencia entre la suma de AX y la suma de U, la divide por la cantidad de filas y no permite resultados negativos.", _
        wsDestino.Name & "!G4:G" & ultimaFilaDestino & ", P4:P" & ultimaFilaDestino & ", " & _
        "AX4:AX" & ultimaFilaDestino & " y U4:U" & ultimaFilaDestino & ".", _
        "Por grupo: AZ=MAX(0;(SUMA(AX)-SUMA(U))/CONTAR(filas del grupo)); el mismo valor se escribe en todas las filas del G-P. Destino: " & _
        wsDestino.Name & "!AZ4:AZ" & ultimaFilaDestino & ".", _
        cantidadFilas, duracionSegundos

    If filaDetalle <> CANTIDAD_REGISTROS Then

        Err.Raise _
            vbObjectError + 1001, _
            "RegistrarLogCalculo", _
            "La cantidad de registros del log no coincide."

    End If

    'La última corrida siempre comienza en la fila 2.
    wsLog.Cells(2, 1).Resize( _
        CANTIDAD_REGISTROS, 10).Value2 = detalle

End Sub


'==========================================================
' LOG DE ERROR
'
' Borra el log anterior y deja solamente el error
' correspondiente a la última corrida.
'==========================================================

Private Sub RegistrarLogError( _
    ByVal wb As Workbook, _
    ByVal nombreHoja As String, _
    ByVal inicioProceso As Date, _
    ByVal duracionSegundos As Double, _
    ByVal numeroError As Long, _
    ByVal descripcionError As String)

    Dim wsLog As Worksheet
    Dim detalle(1 To 1, 1 To 10) As Variant

    Set wsLog = ObtenerHojaLog(wb)

    PrepararHojaLog wsLog
    LimpiarLogAnterior wsLog

    detalle(1, 1) = _
        Format$(inicioProceso, "yyyymmdd_hhnnss")

    detalle(1, 2) = inicioProceso
    detalle(1, 3) = "ERROR"
    detalle(1, 4) = nombreHoja
    detalle(1, 5) = "GENERAL"

    detalle(1, 6) = _
        "La ejecución fue interrumpida antes de completar los cálculos."

    detalle(1, 7) = _
        "Error " & numeroError & ": " & descripcionError

    detalle(1, 8) = _
        "Revisar el mensaje de error y compilar el proyecto VBA."

    detalle(1, 9) = 0
    detalle(1, 10) = duracionSegundos

    wsLog.Cells(2, 1).Resize(1, 10).Value2 = detalle

End Sub


'==========================================================
' BORRAR LOG ANTERIOR
'
' Conserva los encabezados de la fila 1.
'==========================================================

Private Sub LimpiarLogAnterior( _
    ByVal wsLog As Worksheet)

    Dim ultimaFilaLog As Long

    ultimaFilaLog = wsLog.Cells( _
        wsLog.Rows.Count, "A").End(xlUp).Row

    If ultimaFilaLog >= 2 Then

        wsLog.Range( _
            "A2:J" & ultimaFilaLog).ClearContents

    End If

End Sub


'==========================================================
' AGREGAR REGISTRO AL ARREGLO DEL LOG
'==========================================================

Private Sub AgregarDetalleLog( _
    ByRef detalle As Variant, _
    ByRef filaDetalle As Long, _
    ByVal idEjecucion As String, _
    ByVal fechaHora As Date, _
    ByVal nombreHoja As String, _
    ByVal columna As String, _
    ByVal calculo As String, _
    ByVal fuentes As String, _
    ByVal metodo As String, _
    ByVal cantidadFilas As Long, _
    ByVal duracionSegundos As Double)

    filaDetalle = filaDetalle + 1

    detalle(filaDetalle, 1) = idEjecucion
    detalle(filaDetalle, 2) = fechaHora
    detalle(filaDetalle, 3) = "OK"
    detalle(filaDetalle, 4) = nombreHoja
    detalle(filaDetalle, 5) = columna
    detalle(filaDetalle, 6) = calculo
    detalle(filaDetalle, 7) = fuentes
    detalle(filaDetalle, 8) = metodo
    detalle(filaDetalle, 9) = cantidadFilas
    detalle(filaDetalle, 10) = duracionSegundos

End Sub


'==========================================================
' OBTENER O CREAR HOJA LOG
'==========================================================

Private Function ObtenerHojaLog( _
    ByVal wb As Workbook) As Worksheet

    Dim wsLog As Worksheet

    If ExisteHoja(HOJA_LOG, wb) Then

        Set wsLog = wb.Worksheets(HOJA_LOG)

    Else

        Set wsLog = wb.Worksheets.Add( _
            After:=wb.Worksheets(wb.Worksheets.Count))

        wsLog.Name = HOJA_LOG

    End If

    Set ObtenerHojaLog = wsLog

End Function


'==========================================================
' PREPARAR HOJA LOG
'==========================================================

Private Sub PrepararHojaLog( _
    ByVal wsLog As Worksheet)

    Dim encabezados(1 To 1, 1 To 10) As Variant

    encabezados(1, 1) = "ID Ejecución"
    encabezados(1, 2) = "Fecha y hora"
    encabezados(1, 3) = "Estado"
    encabezados(1, 4) = "Hoja calculada"
    encabezados(1, 5) = "Columna destino"
    encabezados(1, 6) = "Qué se calcula"
    encabezados(1, 7) = "Origen: hojas y rangos"
    encabezados(1, 8) = "Proceso aplicado y destino"
    encabezados(1, 9) = "Filas procesadas"
    encabezados(1, 10) = "Duración total (s)"

    With wsLog

        If .AutoFilterMode Then
            .AutoFilterMode = False
        End If

        .Range("A1:J1").Value2 = encabezados
        .Range("A1:J1").Font.Bold = True
        .Range("A1:J1").AutoFilter

        .Columns("A").ColumnWidth = 20
        .Columns("B").ColumnWidth = 20
        .Columns("C").ColumnWidth = 12
        .Columns("D").ColumnWidth = 22
        .Columns("E").ColumnWidth = 15
        .Columns("F").ColumnWidth = 75
        .Columns("G").ColumnWidth = 85
        .Columns("H").ColumnWidth = 95
        .Columns("I").ColumnWidth = 18
        .Columns("J").ColumnWidth = 20

        .Columns("B").NumberFormat = _
            "dd-mm-yyyy hh:mm:ss"

        .Columns("J").NumberFormat = "0.000"
        .Columns("F:H").WrapText = True
        .Rows(1).WrapText = True
        .Rows(1).VerticalAlignment = xlCenter
        .Range("A:J").VerticalAlignment = xlTop

    End With

End Sub


'==========================================================
' CLAVES
'==========================================================

Private Function CrearClave2( _
    ByVal valor1 As Variant, _
    ByVal valor2 As Variant) As String

    CrearClave2 = _
        NormalizarValor(valor1) & SEP & _
        NormalizarValor(valor2)

End Function


Private Function CrearClave4( _
    ByVal valor1 As Variant, _
    ByVal valor2 As Variant, _
    ByVal valor3 As Variant, _
    ByVal valor4 As Variant) As String

    CrearClave4 = _
        NormalizarValor(valor1) & SEP & _
        NormalizarValor(valor2) & SEP & _
        NormalizarValor(valor3) & SEP & _
        NormalizarValor(valor4)

End Function


'==========================================================
' NORMALIZAR VALOR
'==========================================================

Private Function NormalizarValor( _
    ByVal valor As Variant) As String

    If IsError(valor) Then

        NormalizarValor = "#ERROR"

    ElseIf IsNull(valor) Or IsEmpty(valor) Then

        NormalizarValor = vbNullString

    Else

        NormalizarValor = _
            UCase$(Trim$(CStr(valor)))

    End If

End Function


'==========================================================
' CONVERTIR A NÚMERO
'==========================================================

Private Function NumeroSeguro( _
    ByVal valor As Variant) As Double

    If IsError(valor) Then

        NumeroSeguro = 0

    ElseIf IsNull(valor) Or IsEmpty(valor) Then

        NumeroSeguro = 0

    ElseIf IsNumeric(valor) Then

        NumeroSeguro = CDbl(valor)

    Else

        NumeroSeguro = 0

    End If

End Function


'==========================================================
' CONVERTIR A TEXTO
'==========================================================

Private Function TextoSeguro( _
    ByVal valor As Variant) As String

    If IsError(valor) Then

        TextoSeguro = vbNullString

    ElseIf IsNull(valor) Or IsEmpty(valor) Then

        TextoSeguro = vbNullString

    Else

        TextoSeguro = CStr(valor)

    End If

End Function


'==========================================================
' VALIDAR NÚMERO
'==========================================================

Private Function EsNumeroValido( _
    ByVal valor As Variant) As Boolean

    If IsError(valor) Then

        EsNumeroValido = False

    ElseIf IsNull(valor) Or IsEmpty(valor) Then

        EsNumeroValido = False

    ElseIf Len(Trim$(CStr(valor))) = 0 Then

        EsNumeroValido = False

    Else

        EsNumeroValido = IsNumeric(valor)

    End If

End Function


'==========================================================
' OBTENER ÚLTIMA FILA
'==========================================================

Private Function ObtenerUltimaFila( _
    ByVal ws As Worksheet, _
    ByVal columnas As Variant, _
    ByVal filaMinima As Long) As Long

    Dim columna As Variant
    Dim filaActual As Long
    Dim filaMayor As Long

    filaMayor = 0

    For Each columna In columnas

        filaActual = ws.Cells( _
            ws.Rows.Count, _
            CStr(columna)).End(xlUp).Row

        If filaActual > filaMayor Then
            filaMayor = filaActual
        End If

    Next columna

    If filaMayor < filaMinima Then
        ObtenerUltimaFila = filaMinima - 1
    Else
        ObtenerUltimaFila = filaMayor
    End If

End Function


'==========================================================
' VALIDAR EXISTENCIA DE HOJA
'==========================================================

Private Function ExisteHoja( _
    ByVal nombreHoja As String, _
    ByVal wb As Workbook) As Boolean

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = wb.Worksheets(nombreHoja)
    On Error GoTo 0

    ExisteHoja = Not ws Is Nothing

    Set ws = Nothing

End Function


'==========================================================
' DURACIÓN CON CONTROL DE CAMBIO DE DÍA
'==========================================================

Private Function SegundosTranscurridos( _
    ByVal inicioTimer As Double) As Double

    Dim finTimer As Double

    finTimer = Timer

    If finTimer >= inicioTimer Then

        SegundosTranscurridos = _
            finTimer - inicioTimer

    Else

        SegundosTranscurridos = _
            86400# - inicioTimer + finTimer

    End If

End Function






```

### Módulo `ThisWorkbook`

```vb
Attribute VB_Name = "ThisWorkbook"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True

```

### Módulo `Hoja1`

```vb
Attribute VB_Name = "Hoja1"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True

```

### Módulo `Hoja2`

```vb
Attribute VB_Name = "Hoja2"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True

```

### Módulo `Hoja3`

```vb
Attribute VB_Name = "Hoja3"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True

```

### Módulo `Hoja4`

```vb
Attribute VB_Name = "Hoja4"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True

```

### Módulo `Hoja5`

```vb
Attribute VB_Name = "Hoja5"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True

```

### Módulo `Hoja6`

```vb
Attribute VB_Name = "Hoja6"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True

```

### Módulo `Hoja7`

```vb
Attribute VB_Name = "Hoja7"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True

```

### Módulo `Hoja8`

```vb
Attribute VB_Name = "Hoja8"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True

```

### Módulo `Hoja9`

```vb
Attribute VB_Name = "Hoja9"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True

```

### Módulo `Hoja10`

```vb
Attribute VB_Name = "Hoja10"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True

```

### Módulo `Hoja11`

```vb
Attribute VB_Name = "Hoja11"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True

```

### Módulo `Hoja12`

```vb
Attribute VB_Name = "Hoja12"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True

```

### Módulo `Hoja13`

```vb
Attribute VB_Name = "Hoja13"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True

```

### Módulo `Hoja14`

```vb
Attribute VB_Name = "Hoja14"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True

```

### Módulo `Hoja15`

```vb
Attribute VB_Name = "Hoja15"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True

```

### Módulo `Hoja16`

```vb
Attribute VB_Name = "Hoja16"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True

```

### Módulo `Hoja17`

```vb
Attribute VB_Name = "Hoja17"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True

```

### Módulo `Hoja18`

```vb
Attribute VB_Name = "Hoja18"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True

```

### Módulo `Hoja19`

```vb
Attribute VB_Name = "Hoja19"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True

```

### Módulo `Hoja20`

```vb
Attribute VB_Name = "Hoja20"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True

```

---

**Estado de esta versión:** primera radiografía completa del entregable final. Ya quedan identificadas varias fuentes externas y la lógica VBA principal; permanecen abiertos los orígenes de `Medidores!J`, `Prorrata SSCC`, `PRORRATA_RETIROS`, parte de parámetros maestros y el detalle funcional completo de `Calculo RE545`.

---

## Nueva trazabilidad confirmada a partir de scripts Python aportados por el usuario

### Entrada resuelta: `Medidas_SAE.xlsx`

**Estado:** CONFIRMADO.

El archivo `Medidas_SAE.xlsx`, que luego es consumido por la macro `Cargar_Medidas_SAE_En_Medidores`, es generado por el script:

- `2 generación claves Balance.py`

Ese script consolida archivos `medidas_batch_*.parquet`, filtra puntos de medida completos, cruza con `homol.parquet`, aplica el sentido de flujo, genera calendario cuarto-horario y horario, agrupa por `clave` y finalmente exporta:

- Archivo: `Medidas_SAE.xlsx`
- Hoja: `Medidas`
- Columnas exportadas: `Mes`, `Dia`, `Hora`, `Minutos`, `Hora Mes`, `Cuarto de Hora`, `clave`, `intervalo`, `Gen_Unidad`.

Por lo tanto, la cadena de trazabilidad confirmada es:

```text
Homologacion ClavesTF y PRMTE.xlsx
        ↓
0 diccionario prmte a claves_balance.py
        ↓
homol.parquet
        ↓
1 generación prmte.py
        ↓
API medidas del Coordinador
https://medidas.api.coordinador.cl/medidas/api/medidas/{periodo}/
        ↓
medidas_batch_*.parquet
        ↓
2 generación claves Balance.py
        ↓
Medidas_SAE.xlsx / hoja Medidas
        ↓
macro Cargar_Medidas_SAE_En_Medidores
        ↓
Medidores!A:I
```

### Detalle de scripts de esta cadena

#### `0 diccionario prmte a claves_balance.py`

Lee:

- `Homologacion ClavesTF y PRMTE.xlsx`
- Hoja `homol`

Y genera:

- `homol.parquet`

Este archivo contiene, entre otros campos, la relación necesaria para asociar punto de medida/canal con una `clave` de balance.

#### `1 generación prmte.py`

Lee:

- `homol.parquet`

Obtiene los `Punto de Medida` únicos y consulta la API de medidas del Coordinador para el período configurado. Consulta al menos los canales `idCanal=1` e `idCanal=3`.

Genera lotes:

- `medidas_batch_001.parquet`
- `medidas_batch_002.parquet`
- etc.

Además genera `puntos_fallidos.xlsx` como control de puntos que no pudieron ser recuperados.

#### `2 generación claves Balance.py`

Consolida los `medidas_batch_*.parquet`, valida completitud cuarto-horaria, deja medidas principales, calcula:

- `canalVal = canalVal1 + canalVal3`
- aplicación del signo mediante `Flujo`
- `Cuarto de Hora`
- `Hora Mes`
- `Hora`

Cruza directamente con `homol.parquet` para obtener `clave` y agrupa la generación como `Gen_Unidad`.

Genera:

- `1_inyecciones_pto_medidas.parquet`
- `2_generacion_PRMTE_clave_balance.parquet`
- `reporte_medidas_consolidadas.xlsx`
- `Medidas_SAE.xlsx`

### Script adicional: `3 Generacion Real.py`

**Estado respecto de las 5 entradas macro identificadas:** NO CONECTADO DIRECTAMENTE TODAVÍA.

Este script consulta otro endpoint del Coordinador:

```text
https://operacion.api.coordinador.cl/opreal-medidas/v1/bydate
```

para topologías específicas asociadas a BESS/SAE, selecciona el registro definitivo por mayor `idMeasure`, transforma medidas horarias a cuarto-horarias y genera:

- `medidas_15min_AAAAMM.xlsx`
- `log_inconsistencias_medidas.xlsx`

Con la evidencia disponible todavía no se puede afirmar que este archivo alimente directamente `Medidas_SAE.xlsx`, `cmg.xlsx`, `SSCC_Desempeño_*`, `3_REMUNERACIÓN_SUBASTAS_E_ID_*` u `OfertasSSCC`.

Debe mantenerse como **rama potencialmente relevante**, especialmente para datos de generación real/BESS, hasta identificar dónde se consume su salida.

### Estado actualizado de entradas cargadas por macro

1. `Medidas_SAE.xlsx` → **ORIGEN IDENTIFICADO Y CADENA PARCIALMENTE TRAZADA HASTA API + HOMOLOGACIÓN.**
2. `cmg.xlsx` → pendiente.
3. `SSCC_Desempeño_*` → pendiente.
4. `3_REMUNERACIÓN_SUBASTAS_E_ID_*` → pendiente.
5. `*OfertasSSCC*` → pendiente.

### Origen inicial aguas arriba confirmado

- `Homologacion ClavesTF y PRMTE.xlsx`, hoja `homol`.

**Estado:** origen inicial recibido.

El archivo es **entregado por otra persona del departamento** y, para efectos de esta ingeniería inversa, se considera un **origen inicial**. No se rastrea más atrás por ahora. Este Excel alimenta el script `0 diccionario prmte a claves_balance.py`, que genera `homol.parquet`, utilizado posteriormente en la construcción de `Medidas_SAE.xlsx`.

## Actualización de trazabilidad — CMg

### Estado
**Entrada macro resuelta:** `cmg.xlsx`

### Cadena identificada
```text
cmg2607_def_15minutal.csv
    ↓
Extrae CMG barras.py
    ↓
cmg.xlsx
    ↓
Macro Cargar_CMg_Desde_Archivo
    ↓
Hoja CMg del libro 11_PAGOS_BESS_2607_Definitivo.xlsm
    ↓
Calculo E Costos / Calculo RE545
```

### Evidencia del script `Extrae CMG barras.py`
El script toma como entrada el archivo `cmg2607_def_15minutal.csv`, ubicado en la misma carpeta del script, y genera `cmg.xlsx`.

El CSV se lee con separador `;` y codificación `latin1`. Se convierten `FECHA`, `HORA` y `MINUTO`; el campo `CMg[CLP/KWh]` se normaliza como numérico.

El script construye un índice global de `Cuarto de Hora` a partir de los bloques reales presentes en el CSV, contemplando que la cantidad de cuartos pueda variar por día.

Luego filtra únicamente las siguientes barras:
- `PVHFAMI_______033`
- `TOCOPILLA_____110`
- `DONHUMBERTO___033`
- `LA_CABANA_____220`
- `PVSOLDESIERTO_033`
- `PEQ___________220`
- `VICTOR_JARA___220`
- `FUTURO________033`
- `ANDES_SNG_____220`

Para cada combinación `FECHA + HORA + BARRA`, calcula el promedio horario de `CMg[CLP/KWh]`, manteniendo además los registros cuarto-horarios y el `Cuarto de Hora` calculado. La salida se guarda como `cmg.xlsx`.

### Punto aún pendiente aguas arriba
El origen de `cmg2607_def_15minutal.csv` **todavía no está identificado**. Por lo tanto, la entrada `cmg.xlsx` queda resuelta hasta ese CSV, pero aún falta rastrear qué proceso, sistema o descarga genera `cmg2607_def_15minutal.csv`.

### Estado actualizado de entradas externas cargadas por macros
1. `Medidas_SAE.xlsx` — **resuelta** hasta sus orígenes iniciales: API de Medidas del Coordinador + `Homologacion ClavesTF y PRMTE.xlsx` (entregado por otra persona del departamento).
2. `cmg.xlsx` — **resuelta parcialmente** hasta `cmg2607_def_15minutal.csv`; origen del CSV pendiente.
3. `SSCC_Desempeño_*` — pendiente.
4. `3_REMUNERACIÓN_SUBASTAS_E_ID_*` — pendiente.
5. `*OfertasSSCC*` — pendiente.

## Actualización de trazabilidad — confirmaciones del usuario

### Orígenes iniciales definidos por el usuario

- `cmg2607_def_15minutal.csv`: se recibe externamente. Para efectos de esta ingeniería inversa se considera **origen inicial** y no se rastrea más atrás por ahora.
- `SSCC_Desempeño_*`: se recibe externamente. Se considera **origen inicial** y no se rastrea más atrás por ahora.
- `3_REMUNERACIÓN_SUBASTAS_E_ID_*`: se recibe externamente. Se considera **origen inicial** y no se rastrea más atrás por ahora.

### Búsqueda de `*OfertasSSCC*`

La macro `Generar_Resumen_Ofertas_SSCC` llama a `OSSCC_BuscarArchivoOfertas(wbDestino.Path, wbDestino.FullName)`. Por tanto, la búsqueda se realiza en **la misma carpeta donde está guardado el libro `.xlsm` de balance** (`wbDestino.Path`).

Patrón utilizado:

```text
*OfertasSSCC*.*
```

La función acepta solamente extensiones `.xlsx`, `.xlsm`, `.xlsb` y `.xls`, ignora archivos temporales cuyo nombre comienza por `~$`, excluye el propio libro de destino y, si encuentra varios candidatos, selecciona el archivo con **fecha de modificación (`FileDateTime`) más reciente**.

Estado: el origen previo del archivo `*OfertasSSCC*` sigue pendiente; solo queda confirmado dónde y cómo lo busca el balance.


## Actualización de trazabilidad — Homologación ClavesTF y PRMTE

- **Archivo:** `Homologacion ClavesTF y PRMTE.xlsx`
- **Hoja utilizada:** `homol`
- **Estado:** **Origen inicial**
- **Procedencia:** archivo entregado por otra persona del departamento.
- **Tratamiento en esta trazabilidad:** no se investiga su generación aguas arriba por ahora; se toma como punto inicial conocido.
- **Uso posterior:** `0 diccionario prmte a claves_balance.py` lee la hoja `homol` y genera `homol.parquet`; este parquet se utiliza después para relacionar `Punto de Medida + Canal` con `clave` en la cadena que termina generando `Medidas_SAE.xlsx`.

Cadena consolidada:

```text
Homologacion ClavesTF y PRMTE.xlsx   [ORIGEN INICIAL — entregado por otra persona del departamento]
        ↓
0 diccionario prmte a claves_balance.py
        ↓
homol.parquet
        ↓
1 generación prmte.py + API de Medidas del Coordinador
        ↓
medidas_batch_*.parquet
        ↓
2 generación claves Balance.py
        ↓
Medidas_SAE.xlsx
        ↓
Macro del balance
        ↓
Medidores!A:I
```
