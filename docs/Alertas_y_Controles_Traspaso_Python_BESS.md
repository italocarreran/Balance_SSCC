# Catálogo de alertas y controles para el traspaso a Python — Balance BESS / SSCC

> **Objetivo:** preservar y reforzar en Python los controles que hoy existen en el libro Excel/VBA, incluyendo mensajes de macros, verificaciones auxiliares, errores visibles como `#N/D` y controles de consistencia que evitan que una omisión o un cruce incorrecto termine alterando pagos a empresas.
>
> **Estado:** documento vivo. Las secciones **25 a 27**, al final, dicen qué está
> implementado hoy en Python, qué del catálogo quedó desactualizado y qué controles
> se agregaron después de la primera revisión contra el código. Leer esas tres antes
> de planificar trabajo sobre este catálogo.
>
> **Principio:** un error no debe transformarse silenciosamente en `0`, vacío, `NaN` o en una fila descartada. Toda sustitución, omisión o falta de coincidencia con impacto potencial debe quedar registrada y, cuando corresponda, bloquear la generación del resultado final.

---

## 1. Clasificación de severidad propuesta

| Nivel | Nombre | Comportamiento Python propuesto |
|---|---|---|
| **CRÍTICA** | Bloqueante | Detener la ejecución antes de generar el entregable final. Se puede conservar una salida parcial de diagnóstico. |
| **ALTA** | Requiere revisión | Terminar los cálculos de diagnóstico, pero marcar la corrida como **NO APROBADA**. No permitir exportación final de pagos sin revisión explícita. |
| **MEDIA** | Advertencia | Continuar, registrar detalle completo y mostrar resumen al final. |
| **INFO** | Informativa | Registrar para trazabilidad; no implica error. |

La aplicación debe finalizar mostrando un estado inequívoco:

```text
APROBADA
APROBADA CON ADVERTENCIAS
NO APROBADA - REQUIERE REVISIÓN
FALLIDA - ERROR BLOQUEANTE
```

No utilizar solamente mensajes emergentes. Todas las alertas deben persistir en un archivo/log.

---

# 2. Formato mínimo de una alerta

Cada alerta registrada por Python debe contener, cuando aplique:

| Campo | Contenido |
|---|---|
| `id_alerta` | Código único, por ejemplo `MED-001` |
| `severidad` | CRÍTICA / ALTA / MEDIA / INFO |
| `etapa` | Medidores, SoC, Subastas, FD, Ofertas, RE545, etc. |
| `archivo` | Archivo origen relacionado |
| `hoja` | Hoja origen, si aplica |
| `fila_origen` | Fila del archivo origen |
| `fila_destino` | Fila de la tabla calculada, si aplica |
| `central` | Central / BESS |
| `configuracion` | Configuración, cuando corresponda |
| `fecha_hora` | Fecha/hora o período |
| `clave` | Clave de búsqueda que falló |
| `mensaje` | Descripción comprensible |
| `valor_encontrado` | Valor problemático |
| `valor_esperado` | Valor o condición esperada |
| `accion_python` | Detuvo / reemplazó por 0 / omitió / continuó |
| `origen_control` | `VBA EXISTENTE`, `ERROR EXCEL`, `CONTROL NUEVO` |

Salida recomendada:

```text
<CARPETA_BASE>/
└── Control/
    ├── Alertas.xlsx
    ├── Alertas.csv
    └── Resumen_Ejecucion.txt
```

La ruta definitiva puede cambiar, pero la existencia de un registro persistente debe considerarse obligatoria.

---

# 3. Controles de estructura y archivos de entrada

### `EST-001` — No existe carpeta `Medidas`
- **Severidad:** CRÍTICA.
- **Condición:** no existe `<BASE>/Medidas/`.
- **Acción:** detener.

### `EST-002` — No existe carpeta `Auxiliares`
- **Severidad:** CRÍTICA.

### `EST-003` — No existe carpeta `Ofertas`
- **Severidad:** CRÍTICA cuando la etapa requiera ofertas.

---

# 4. Controles de `Medidas_SAE.xlsx`

### `MED-001` — Archivo no encontrado
- **Severidad:** CRÍTICA.
- Falta `<BASE>/Medidas/Medidas_SAE.xlsx`.

### `MED-002` — Hoja `Medidas` no existe
- **Severidad:** CRÍTICA.
- **Origen VBA:** `Cargar_Medidas_SAE_En_Medidores`.

### `MED-003` — Hoja `Medidas` sin datos
- **Severidad:** CRÍTICA.
- **Origen VBA:** `Cargar_Medidas_SAE_En_Medidores`.

### `MED-004` — Columnas requeridas A:I incompletas
- **Severidad:** CRÍTICA.

### `MED-005` — Campos clave vacíos
- **Severidad:** ALTA.
- No descartar silenciosamente filas con campos utilizados para central/fecha/hora/cuarto/configuración.

### `MED-006` — Duplicado de clave temporal
- **Severidad:** ALTA.
- Clave recomendada: central/configuración + fecha + hora + cuarto de hora.

### `MED-007` — Período distinto al período de ejecución
- **Severidad:** CRÍTICA.

---

# 5. Controles del archivo `SOC_AAMM.xlsx`

### `SOC-001` — No existe archivo `SOC_AAMM`
- **Severidad:** CRÍTICA.

### `SOC-002` — Más de un `SOC_*.xlsx`
- **Severidad:** CRÍTICA.
- No escoger automáticamente “el más reciente”.

### `SOC-003` — El AAMM del nombre no coincide con la corrida
- **Severidad:** CRÍTICA.

### `SOC-004` — No se detectaron bloques de centrales
- **Severidad:** CRÍTICA.

### `SOC-005` — Central del archivo SoC no reconocida
- **Severidad:** ALTA.

### `SOC-006` — BESS esperado no aparece en SoC
- **Severidad:** ALTA.

### `SOC-007` — No se encuentra `Time Stamp` dentro del bloque
- **Severidad:** CRÍTICA para esa central.

### `SOC-008` — No se encuentra `Value` dentro del bloque
- **Severidad:** CRÍTICA para esa central.

### `SOC-009` — Más de un candidato `Time Stamp` o `Value`
- **Severidad:** ALTA.
- No elegir silenciosamente.

### `SOC-010` — Timestamp inválido
- **Severidad:** ALTA.

### `SOC-011` — Timestamp duplicado para la misma central
- **Severidad:** ALTA.

### `SOC-012` — Salto inesperado en la secuencia temporal
- **Severidad:** ALTA.

### `SOC-013` — Falta un cuarto de hora esperado
- **Severidad:** ALTA.

### `SOC-014` — SoC fuera de rango o no numérico
- **Severidad:** ALTA.
- El archivo usa proporción decimal; no transformar escala silenciosamente.

## Cruce SoC → Medidores

Clave confirmada:

```text
Central + Fecha + Hora + Cuarto de hora
```

### `SOC-015` — Fila de Medidores sin SoC
- **Severidad:** ALTA.
- Prohibido llenar con 0 sin alerta.

### `SOC-016` — Más de un SoC para una fila de Medidores
- **Severidad:** CRÍTICA.

### `SOC-017` — Registros SoC no utilizados
- **Severidad:** MEDIA.

### `SOC-018` — Cobertura SoC incompleta
- **Severidad:** ALTA.
- Recomendación: exigir 100 % salvo excepciones documentadas.

---

# 6. Controles de `Centrales.xlsx` / Diccionario / parámetros

### `AUX-001` — `Centrales.xlsx` no encontrado
- **Severidad:** CRÍTICA.

### `AUX-002` — Hoja `Resumen BESS` no encontrada
- **Severidad:** CRÍTICA.

### `AUX-003` — Hoja `Diccionario` no encontrada
- **Severidad:** CRÍTICA.
- **Origen VBA equivalente:** la lógica de Ofertas falla si no existe `Diccionario`.

### `AUX-004` — Nombre usado en Medidores no aparece en Diccionario
- **Severidad:** ALTA.
- **Origen VBA existente:** la macro de Ofertas advierte los nombres de `Medidores!G` ausentes en `Diccionario!E:G`; hoy los incorpora igualmente con oferta 0.
- **Python:** registrar todos los casos.

### `AUX-005` — Nombre con equivalencias ambiguas
- **Severidad:** ALTA.

### `AUX-006` — Parámetro maestro faltante para un BESS
- **Severidad:** CRÍTICA si participa en energía o dinero.

### `AUX-007` — Parámetro numérico contiene texto/error
- **Severidad:** CRÍTICA.

### `AUX-008` — Factor de `Resumen B:C` no encontrado
- **Severidad:** CRÍTICA.
- **Origen Excel/VBA:** AE/AF mantienen comportamiento `#N/D` cuando no existe la clave G.
- **Python:** no convertir ese caso en 0.

---

# 7. Controles de Ofertas SSCC

### `OFE-001` — No existe archivo `*OfertasSSCC*`
- **Severidad:** CRÍTICA.
- **Origen VBA:** `Generar_Resumen_Ofertas_SSCC`.

### `OFE-002` — Más de un archivo candidato
- **Severidad:** ALTA/CRÍTICA.
- Mostrar candidatos y regla de selección aplicada.

### `OFE-003` — Sin registros BESS/SAE/BAT con servicio `_RS`
- **Severidad:** CRÍTICA.
- **Origen VBA existente.**

### `OFE-004` — Año/mes no determinable
- **Severidad:** CRÍTICA.

### `OFE-005` — Resumen contiene más de un año o mes
- **Severidad:** CRÍTICA.
- **Origen VBA existente.**

### `OFE-006` — Servicio `_RS` sin 24 horas ofrecidas
- **Severidad:** MEDIA.
- Registrar central, día, servicio y cantidad de períodos detectados.

### `OFE-007` — Hora/período duplicado dentro de un servicio
- **Severidad:** MEDIA.
- La macro actual cuenta horas repetidas una sola vez; Python debe informar el duplicado.

### `OFE-008` — Período fuera de 1..24
- **Severidad:** ALTA.

### `OFE-009` — Nombre de Medidores no representado en Ofertas ni por equivalencia
- **Severidad:** ALTA.
- **Comportamiento actual:** se agrega con oferta 0 y se advierte.

### `OFE-010` — Día sin registro de oferta
- **Severidad:** MEDIA/ALTA.
- Registrar que el 0 proviene de ausencia de dato.

### `OFE-011` — Falta encabezado `Oferta completa`
- **Severidad:** CRÍTICA.
- **Origen VBA existente.**

### `OFE-012` — Resumen de ofertas vacío
- **Severidad:** CRÍTICA.

### `OFE-013` — Ningún nombre disponible para construir W:Y
- **Severidad:** CRÍTICA.

### `OFE-014` — Cantidad de filas W:Y inesperada
- **Severidad:** ALTA.
- Debe cumplirse, salvo justificación:

```text
cantidad de nombres × días del mes
```

---

# 8. Controles de ventana y oferta completa `AB:AE`

### `VEN-001` — Hora de inicio equivalente a `Medidores!S1` inválida
- **Severidad:** CRÍTICA.
- **Origen VBA existente:** debe estar entre 1 y 24.

### `VEN-002` — No existen datos G/L/R para resumir
- **Severidad:** CRÍTICA.

### `VEN-003` — No existe ninguna ventana numérica en L
- **Severidad:** CRÍTICA.
- **Origen VBA existente.**

### `VEN-004` — No se generaron grupos Central + Ventana
- **Severidad:** CRÍTICA.

### `VEN-005` — Oferta esperada distinta del diseño
- **Severidad:** ALTA.
- Para `S1=10`: ventana 0 = 36; normales = 96; última = 60.

### `VEN-006` — Ventana marcada incompleta
- **Severidad:** MEDIA.
- Registrar central, ventana, oferta acumulada, esperada y diferencia.

### `VEN-007` — Más de un resumen para la misma Central + Ventana
- **Severidad:** CRÍTICA.

---

# 9. Controles de Subastas

### `SUB-001` — Archivo `3_REMUNERACIÓN_SUBASTAS_E_ID_*` no encontrado
- **Severidad:** CRÍTICA.
- **Origen VBA existente.**

### `SUB-002` — Hoja `DB` no encontrada
- **Severidad:** CRÍTICA.

### `SUB-003` — Columnas necesarias de DB incompletas
- **Severidad:** CRÍTICA.
- Se necesitan al menos `B:L`, `P`, `V`, `Y`.

### `SUB-004` — No existen filas donde DB!K contenga BESS o SAE
- **Severidad:** CRÍTICA/ALTA.

### `SUB-005` — Clave equivalente a M no generable
- **Severidad:** CRÍTICA.
- `M = K&H&I`.

### `SUB-006` — Cruce equivalente a N no encuentra coincidencia
- **Severidad:** ALTA.
- **Excel actual:** `IFERROR(XLOOKUP(...),"")` esconde el fallo como vacío.
- **Python:** registrar cada fila sin match.

### `SUB-007` — Tipo de servicio no reconocido
- **Severidad:** ALTA.
- **Origen Excel:** la fórmula auxiliar puede devolver `REVISAR`.
- Cualquier `REVISAR` debe convertirse en alerta.

### `SUB-008` — DB!Y faltante
- **Severidad:** ALTA.
- Impacto: DB!Y → Subastas!P → `Calculo RE545!AI:AN`.

### `SUB-009` — DB!V faltante
- **Severidad:** ALTA.
- Impacto: DB!V → Subastas!Q → `Calculo RE545!AO:AQ`.

### `SUB-010` — Duplicados en claves usadas por cálculos posteriores
- **Severidad:** MEDIA/ALTA.
- Distinguir duplicación sumable esperada de doble contabilización accidental.

---

# 10. Controles de FD / SSCC Desempeño

### `FD-001` — Carpeta/archivo `SSCC_Desempeño_*` no encontrado
- **Severidad:** CRÍTICA.

### `FD-002` — Hoja `CPF Horario` no encontrada
- **Severidad:** CRÍTICA.
- **Origen VBA existente.**

### `FD-003` — Hoja `CSF Horario` no encontrada
- **Severidad:** CRÍTICA.
- **Origen VBA existente.**

### `FD-004` — Fórmula plantilla requerida no existe
- **Severidad:** CRÍTICA.
- **Origen VBA existente:** `ValidarPlantillaFormulas`.

### `FD-005` — Clave no encontrada en FD
- **Severidad:** **ALTA**.
- **Comportamiento actual:** continúa utilizando 0 y al final informa faltantes.
- **Python:** guardar **todos** los faltantes, no solo los primeros 15.
- Registrar sector FD, clave, fila destino y columnas destino.
- Recomendación: corrida `NO APROBADA` hasta revisión.

### `FD-007` — Unidad homologada ausente de todo un bloque de FD *(control nuevo)*
- **Severidad:** **ALTA**.
- **Qué separa:** `FD-005` es "a esta unidad le falta **una hora**"; `FD-007` es
  "esta unidad **no aparece nunca** en el bloque CPF/CSF de la hoja FD". Lo
  segundo no es un dato faltante, es una homologación que no cruza (o una
  central que no presta ese servicio), y repetirlo una vez por hora es lo que
  hacía inservible la hoja `Alertas`: en la primera corrida real fueron
  **2.232 alertas `FD-005` que eran 3 hechos**.
- **Python:** una alerta **por central y bloque**, con la unidad a la que está
  homologada, cuántas filas quedan en 0 y **la lista de unidades que sí trae
  ese bloque** (que es la pista para encontrar el nombre correcto). Cuando se
  emite `FD-007` para una unidad, NO se emiten además las `FD-005` de cada una
  de sus horas: son la misma frase repetida.
- **Origen:** `calcular_fd_prorrateado()`, con `unidades_bloque_fd()`.

### `FD-006` — Cantidad excesiva de claves FD faltantes
- **Severidad:** CRÍTICA configurable.
- Si supera umbral absoluto o porcentual definido por negocio, bloquear.

---

# 11. Controles de CMg

### `CMG-001` — Archivo `cmg.xlsx` no encontrado
- **Severidad:** CRÍTICA.

### `CMG-002` — Archivo sin hojas
- **Severidad:** CRÍTICA.
- **Origen VBA existente.**

### `CMG-003` — Sin datos desde la fila esperada
- **Severidad:** CRÍTICA.
- **Origen VBA existente.**

### `CMG-004` — Barra requerida no encontrada
- **Severidad:** ALTA.
- No reemplazar CMg faltante por 0 sin alerta.

### `CMG-005` — Más de un valor para fecha + hora + barra
- **Severidad:** ALTA.

### `CMG-006` — CMg no numérico / error
- **Severidad:** ALTA.

---

# 12. Controles de `Calculo E Costos`

### `ECO-001` — Falta una dependencia requerida
- **Severidad:** CRÍTICA.
- El VBA valida explícitamente `Calculo E Costos`, `Subastas`, `Resumen`, `Prorrata SSCC`, `Diccionario` y `FD`.
- En Python debe existir control equivalente sobre DataFrames/entradas.

### `ECO-002` — Umbral equivalente a `Resumen!H8` contiene error
- **Severidad:** CRÍTICA.
- **Origen VBA existente.**

### `ECO-003` — Umbral H8 no numérico
- **Severidad:** CRÍTICA.
- **Origen VBA existente.**

### `ECO-004` — No hay filas de entrada
- **Severidad:** CRÍTICA.

### `ECO-005` — `#N/D` equivalente
- **Severidad:** CRÍTICA.
- Condición: lookup obligatorio sin correspondencia.
- Aplicación conocida: factores usados por AE/AF.
- No aceptar `NaN → 0`.

### `ECO-006` — `#DIV/0!` equivalente
- **Severidad:** CRÍTICA/ALTA.

### `ECO-007` — `#VALUE!`, `#NUM!`, infinito o NaN inesperado
- **Severidad:** CRÍTICA.
- Escanear todos los campos numéricos calculados antes de aprobar.

### `ECO-008` — Cantidad de filas calculadas distinta de entrada
- **Severidad:** CRÍTICA.

---

# 13. Validación Medidores → E Costos / RE545

El libro actual contiene una validación explícita de conservación de energía.

### `TRA-001` — Energía presente en ambas hojas destino
- **Severidad:** CRÍTICA.
- **Origen VBA existente.**

### `TRA-002` — Energía indebida en `Calculo RE545`
- **Severidad:** CRÍTICA.

### `TRA-003` — Energía indebida en `Calculo E Costos`
- **Severidad:** CRÍTICA.

### `TRA-004` — Diferencia en ambas hojas
- **Severidad:** CRÍTICA.

### `TRA-005` — Diferencia solo en E Costos
- **Severidad:** CRÍTICA.

### `TRA-006` — Diferencia solo en RE545
- **Severidad:** CRÍTICA.

### `TRA-007` — Diferencia total de energía
- **Severidad:** CRÍTICA.
- Debe cumplirse dentro de tolerancia:

```text
Energía Medidores
=
Energía Calculo E Costos
+
Energía Calculo RE545
```

### `TRA-008` — Tolerancia superada
- **Severidad:** CRÍTICA.
- Guardar tolerancia utilizada y diferencia observada.

### `TRA-009` — Reporte detallado de diferencias
- **Control obligatorio.**
- Debe replicar al menos: fila origen, central, energía origen, T, esperado/real E Costos, esperado/real RE545, diferencias y observación.

---

# 14. Controles de `Calculo RE545`

### `RE-001` — Input de Subastas sin match
- **Severidad:** ALTA/CRÍTICA según columna.

### `RE-002` — Bloques AI:AN faltantes por ausencia de DB!Y
- **Severidad:** ALTA.

### `RE-003` — Bloques AO:AQ faltantes por ausencia de DB!V
- **Severidad:** ALTA.

### `RE-004` — Resultado AU no finito
- **Severidad:** CRÍTICA.
- AU combina bloques de Subastas mediante `SUMPRODUCT`.

### `RE-005` — Clave Central + Ventana no encontrada en resúmenes auxiliares
- **Severidad:** ALTA.

### `RE-006` — Resultado monetario no numérico
- **Severidad:** CRÍTICA.

---

# 15. Controles de secuencia temporal y CSV final

La macro actual de exportación crea `Verificacion_CSV` y registra los siguientes controles.

### `CSV-001` — Fecha distinta para mismo bloque
- **Severidad:** ALTA.

### `CSV-002` — Hora mayor a 24
- **Severidad:** ALTA.

### `CSV-003` — Cambio de hora +1 detectado
- **Severidad:** INFO/MEDIA.
- Actual: salto de 75 minutos.

### `CSV-004` — Cambio de hora -1 detectado
- **Severidad:** INFO/MEDIA.
- Actual: diferencia de -45 minutos.

### `CSV-005` — Fecha/hora duplicada
- **Severidad:** ALTA.

### `CSV-006` — Salto horario inesperado
- **Severidad:** ALTA.

### `CSV-007` — Día normal
- 96 bloques.
- **Severidad:** INFO; sin alerta de error.

### `CSV-008` — Día cambio +1
- 92 bloques.
- **Severidad:** INFO/MEDIA.

### `CSV-009` — Día cambio -1
- 100 bloques.
- **Severidad:** INFO/MEDIA.

### `CSV-010` — Día con bloques distintos de 96/92/100
- **Severidad:** ALTA.

### `CSV-011` — No hay registros con pago distinto de cero
- **Severidad:** ALTA.
- **Origen VBA existente.**

### `CSV-012` — Después de consolidar todos los pagos quedaron en cero
- **Severidad:** ALTA.
- **Origen VBA existente.**

---

# 16. Errores que Excel puede convertir silenciosamente en 0

El VBA actual contiene funciones que ante una celda con error, una celda vacía o texto no interpretable como número pueden devolver 0.

En Python no se debe asumir que esos casos equivalen a un cero real.

### `NUM-001` — Error convertido a 0
- **Severidad:** ALTA.
- Registrar valor original y origen.

### `NUM-002` — Texto no numérico convertido a 0
- **Severidad:** ALTA.

### `NUM-003` — Vacío convertido a 0
- **Severidad:** MEDIA/ALTA según campo.

Usar estados separados:

```text
cero_real
dato_faltante
dato_invalido
error_de_lookup
```

No colapsar todos en `0`.

---

# 17. Controles monetarios nuevos recomendados

Estos controles no necesariamente existen como macro explícita, pero son recomendables por el riesgo económico del proceso.

### `MON-001` — Total antes/después de cada etapa
- **Severidad:** ALTA si cambia sin explicación.
- Guardar totales por central, empresa, configuración, tipo de pago y total general.

### `MON-002` — Cambio monetario anormal por empresa/central
- **Severidad:** ALTA.
- Comparar contra ejecución Excel de validación y, más adelante, período anterior.

### `MON-003` — Central con pago pero sin empresa/propietario
- **Severidad:** CRÍTICA.

### `MON-004` — Empresa/central desconocida en maestros
- **Severidad:** ALTA.

### `MON-005` — Merge N:N no autorizado
- **Severidad:** CRÍTICA.
- Puede multiplicar pagos.

### `MON-006` — Pérdida de filas en un merge
- **Severidad:** ALTA/CRÍTICA.
- Informar cantidad y monto asociado.

### `MON-007` — Aumento inesperado de filas en un merge
- **Severidad:** CRÍTICA.

---

# 18. Controles de período

### `PER-001` — Medidas SAE y SOC de meses diferentes
- **Severidad:** CRÍTICA.

### `PER-002` — Ofertas de otro mes
- **Severidad:** CRÍTICA.

### `PER-003` — Subastas de otro mes
- **Severidad:** CRÍTICA.

### `PER-004` — SSCC Desempeño de otro mes
- **Severidad:** CRÍTICA.

### `PER-005` — CMg de otro mes
- **Severidad:** CRÍTICA.

Antes de calcular dinero, resolver un único:

```text
PERIODO_EJECUCION = AAMM
```

y comprobar que todas las entradas pertenecen al mismo período.

---

# 19. Integridad de merges / lookups

Para **cada** búsqueda o merge importante registrar:

```text
filas_entrada
filas_con_match
filas_sin_match
filas_con_match_multiple
porcentaje_cobertura
monto_asociado_a_sin_match
```

### `MRG-001` — Match faltante
- **Severidad:** ALTA o CRÍTICA.

### `MRG-002` — Match múltiple cuando se esperaba uno
- **Severidad:** CRÍTICA.

### `MRG-003` — Cobertura menor a 100 %
- **Severidad:** configurable, nunca silenciosa.

### `MRG-004` — Fila descartada por filtro
- **Severidad:** INFO/MEDIA.
- Registrar cantidad por causa.

---

# 20. Reporte final obligatorio de ejecución

Ejemplo:

```text
EJECUCIÓN BALANCE BESS 2607
===========================

Estado: NO APROBADA - REQUIERE REVISIÓN

Entradas:
[OK] Medidas_SAE.xlsx
[OK] SOC_2607.xlsx
[OK] Centrales.xlsx
[OK] OfertasSSCC
[OK] Subastas
[OK] FD
[OK] CMg

Alertas:
CRÍTICAS : 0
ALTAS     : 3
MEDIAS    : 7
INFO      : 4

Principales observaciones:
- 2 claves no encontradas en FD y sustituidas por 0.
- 1 central no encontrada en Diccionario.
- 4 ventanas con oferta incompleta.

Conciliaciones:
Medidores                  : X MWh
Calculo E Costos           : Y MWh
Calculo RE545              : Z MWh
Diferencia                 : 0.000000 MWh

Total monetario general    : $...
```

---

# 21. Reglas de aprobación recomendadas

## Corrida `APROBADA`

Solo cuando:
- no hay alertas CRÍTICAS;
- no hay alertas ALTAS pendientes;
- conciliación energética está dentro de tolerancia;
- no hay `NaN`, `Inf` o equivalentes a errores Excel en columnas críticas;
- todos los merges críticos cumplen cardinalidad;
- período consistente en todas las fuentes.

## Corrida `NO APROBADA`

Si ocurre cualquiera de:
- falta de archivo/hoja requerida;
- lookup crítico sin match;
- duplicación N:N no autorizada;
- diferencia energética;
- parámetro maestro faltante;
- error equivalente a `#N/D`;
- faltantes FD sin revisión;
- cruce SoC ambiguo;
- período inconsistente;
- dato monetario no numérico.

---

# 22. Filosofía para el traspaso

La versión Python no debe ser solamente más rápida que Excel. Debe ser **más auditable**.

En Excel existen tres tipos de protección que hoy ayudan al usuario:

1. mensajes programados en VBA;
2. errores visibles como `#N/D`;
3. anomalías que se detectan visualmente al revisar las hojas.

Python elimina parte de esa visibilidad si todo queda convertido automáticamente en DataFrames y `NaN`.

Por eso:

> **Cada error que Excel hace visible debe transformarse en una alerta explícita de Python.**

Y además:

> **Cada error que Excel hoy oculta mediante `IFERROR`, vacío o sustitución por 0 debe evaluarse para decidir si merece una alerta adicional.**

La meta no es conseguir una ejecución “sin errores técnicos”; la meta es conseguir una ejecución en que el usuario pueda confiar antes de realizar o publicar pagos.

---

# 23. Priorización para implementación

## Fase 1 — obligatoria desde la primera versión Python

1. archivos/carpetas requeridos;
2. período consistente;
3. SoC sin match / match múltiple;
4. nombres sin homologación;
5. Ofertas faltantes / períodos inconsistentes;
6. Subastas sin match;
7. faltantes FD;
8. `#N/D`/NaN/Inf equivalentes;
9. validación Medidores → E Costos + RE545;
10. cardinalidad de merges;
11. conciliación monetaria básica;
12. reporte persistente de alertas.

## Fase 2 — refuerzo

- comparación contra período anterior;
- detección de outliers monetarios por empresa/central;
- umbrales históricos;
- dashboard de alertas;
- historial de ejecuciones y aprobación/revisión.

---

# 24. Fuente de los controles

## Controles existentes / comportamiento Excel-VBA

Derivados de `Trazabilidad_11_PAGOS_BESS_2607_Definitivo.md`, incluyendo:

- validaciones de archivos y hojas;
- advertencias de Diccionario;
- validaciones de Ofertas;
- validaciones de ventanas;
- faltantes FD;
- validación energética Medidores → E Costos / RE545;
- controles de secuencia temporal y bloques del CSV;
- comportamiento `#N/D` por claves no encontradas;
- mensajes de error de Subastas, CMg, Medidas y SSCC Desempeño.

## Controles nuevos propuestos

Agregados para evitar que Python oculte problemas:

- cardinalidad de merges;
- cobertura de cruces;
- filas perdidas/duplicadas;
- período único entre fuentes;
- `NaN/Inf`;
- distinción entre cero real y dato faltante;
- conciliaciones monetarias;
- aprobación formal de la corrida.

---

## Estado del documento

**Documento inicial de controles para incorporar en el traspaso a Python.**

Debe mantenerse como documento vivo: cada vez que durante la réplica del Excel aparezca una nueva alerta, fórmula con `IFERROR`, sustitución por 0 o validación manual utilizada por el equipo, debe agregarse a este catálogo antes de cerrar la migración.

---

# 25. Estado de implementación en Python

Revisión del catálogo contra el código, hecha el 2026-09-14.

## Ya implementado

| Control | Dónde |
|---|---|
| `EST-*`, `MED-001…003`, `AUX-001…003`, `OFE-001…005`, `CMG-001…003`, `FD-001…003`, `VEN-001/003`, `SUB-002/003` | `ErrorEntrada` en `nucleo/rutas.py`, `nucleo/lectura.py`, `nucleo/hojas_entrada.py` y `nucleo/proceso.py`. Son bloqueantes: detienen la corrida, como pide el catálogo. |
| `SOC-002` / `SOC-003` | `buscar_soc()`: varios `SOC_*.xlsx` del período → error. **No** elige el más reciente, exactamente como pide el catálogo (el resto de los archivos sí lo hace, porque las macros originales lo hacían). |
| `PER-001` (Medidas vs período), `PER-002` (Ofertas de otro mes) | `construir_medidores()`. |
| `PER-001` en la etapa de pagos | `generar_pagos_bess()`: si la hoja `Medidores` del consolidado trae más de un mes, CRÍTICA. |
| `AUX-004` (nombre ausente del `Diccionario`) | `calcular_fd_prorrateado()` → alerta `DIC-001`. |
| `AUX-006` / `AUX-008` (parámetro maestro faltante) | `MAE-001` (barra), `MAE-002` (Pmax), `MAE-004` (capacidad), `MAE-005` (eficiencia), en las dos hojas de cálculo. |
| `AUX-007` (parámetro numérico inválido) | `MAE-003`: Pmax presente pero en 0 o no numérico — deja `AE`/`AF` vacías. |
| `CMG-004` (barra/cuarto sin CMg) | `construir_calculo_e_costos()` y `construir_calculo_re545()`. |
| `FD-005` (clave no encontrada en FD) | `calcular_fd_prorrateado()`. **Se guardan todas**, no una muestra: el tope de 15 es solo de la línea que va a la pantalla. Excepción deliberada: si la unidad no está en NINGUNA fila del bloque, eso va como un solo `FD-007` y no como una `FD-005` por hora. |
| `FD-007` (unidad ausente de todo un bloque de FD) | `calcular_fd_prorrateado()` + `unidades_bloque_fd()`. Una por central y bloque, con la lista de unidades disponibles. |
| `TRA-001`, `TRA-005`, `TRA-006`, `TRA-007`, `TRA-008`, `TRA-009` | `nucleo/conciliacion.py`. Tolerancia fijada: relativa `1e-9`, piso absoluto `1e-6`; las dos quedan escritas en la hoja `Ejecucion` junto con la diferencia observada. |
| §1 (estado de la corrida), §2 (formato de alerta), §20 (reporte final) | `nucleo/alertas.py` + hojas `Alertas` y `Ejecucion` de `Pagos_BESS.xlsx`. |
| §19 parcial (cobertura de cruces) | Cada alerta de cruce trae la central/clave que falló; el conteo sale del propio registro. |

## Controles nuevos, no previstos en el catálogo

| Id | Qué detecta |
|---|---|
| `PRO-001` | Central sin **ninguna** hora en la Prorrata SSCC: la homologación contra `Subastas!Configuración` no está cruzando y se cae todo el prorrateo de SSCC de esa central. |
| `PRO-002` | Filas sueltas sin prorrata para su `central+Hora Mes` (MEDIA: una hora sin SSCC puede ser legítima). |
| `SUB-011` | Central que no aparece en ninguna fila de `Subastas`: sus reservas `AC:AT` y `AU` quedan en 0 y se paga como si no hubiera tenido reservas. |
| `PAG-001` | Una hoja del libro de pagos quedó vacía porque no se pidió recalcularla y no había versión anterior. |

## Pendiente

- Las hojas `Alertas` / `Ejecucion` están en `Pagos_BESS.xlsx`. `Consolidado_entradas.xlsx` sigue con su hoja `Log` de texto libre: falta unificarlas.
- `FALLIDA - ERROR BLOQUEANTE` no se escribe en ningún archivo: si una entrada falta, la corrida se detiene antes de escribir y el motivo queda solo en la ventana. Habría que escribir un libro de diagnóstico aunque el cálculo no termine.
- Todo `§15` (CSV) y `§17` (`MON-*`) sigue sin implementar. Ver la sección 26 sobre el `§15`.

---

# 26. Correcciones al catálogo

Cosas del catálogo que no coinciden con el estado real del proyecto.

## 26.1 `§9` (Subastas) describe un origen que ya no se usa

`SUB-001`, `SUB-002` y `SUB-003` hablan del archivo
`3_REMUNERACIÓN_SUBASTAS_E_ID_*` y de la hoja `DB` de la planilla 3. Ese
intermediario ya se cortó: la hoja `Subastas` se arma directamente desde los
Access `OfertasSSCCAdj*.accdb`. Los tres controles hay que reescribirlos sobre
el origen nuevo, que además necesita los suyos:

- `.accdb` del período no encontrado;
- driver ODBC de Access ausente en el PC (hoy es un `ErrorSubastas` genérico);
- más de un `.accdb` candidato para el mismo período.

`SUB-008` y `SUB-009` (DB!Y y DB!V faltantes) siguen siendo válidos como
concepto, pero su origen ahora son las salidas `fma_*` y el `SSCC_Desempeño_*`,
no columnas de la planilla 3.

## 26.2 `§15` (CSV) es de una etapa que todavía no existe

Los controles `CSV-001…012` se refieren a la macro de exportación a CSV y a su
hoja `Verificacion_CSV`. En Python no hay exportación a CSV todavía. Conviene
marcar la sección como **etapa futura** para que no se mezcle con el trabajo
pendiente de lo que sí existe.

## 26.3 El riesgo de inflar las severidades

`SOC-012` (salto en la secuencia temporal) y `SOC-013` (falta un cuarto de hora)
están como ALTA. En un mes real de SCADA los huecos son habituales. Si un ALTA
se dispara en todas las corridas, el equipo aprende a ignorar los ALTA — y ahí
se pierde el valor de todo el catálogo, incluidos los controles que sí importan.

Recomendación: que arranquen en MEDIA con un umbral, y subirlos a ALTA recién
cuando los datos reales muestren que son raros. El modo de falla de un sistema
de alertas no es quedarse corto: es cansar.

## 26.4 La severidad fija no distingue si la fila paga

Un faltante en una fila con energía 0 y sin reserva no vale lo mismo que uno en
una fila que paga. El propio `§19` ya insinúa la salida
(`monto_asociado_a_sin_match`): conviene usar el monto asociado para decidir la
severidad, en vez de una tabla rígida por tipo de control.

---

# 27. Controles que faltaban

Vacíos detectados al contrastar el catálogo con el código.

### `RUN-001` — Manifiesto de entradas
- **Severidad:** INFO (pero **obligatorio**).
- Sin esto, `APROBADA` no es reproducible: dentro de tres meses nadie puede
  demostrar **cuál** `SOC_2607.xlsx` se usó.
- De cada archivo leído: nombre, ruta, tamaño, fecha de modificación y `sha256`
  del contenido. Dos corridas con el mismo manifiesto tuvieron las mismas
  entradas; si un total cambió y el manifiesto no, el cambio está en el código.
- **Implementado** en `nucleo/manifiesto.py`; se escribe en la hoja `Ejecucion`.

### `RUN-002` — El estado tiene que viajar con el entregable
- **Severidad:** CRÍTICA de diseño.
- Un `Resumen_Ejecucion.txt` en una carpeta `Control/` no acompaña al
  `Pagos_BESS.xlsx` que alguien manda por correo. El estado va estampado en el
  libro mismo.
- **Implementado**: hoja `Ejecucion` dentro de `Pagos_BESS.xlsx`.

### `RUN-003` — Corridas parciales
- **Severidad:** ALTA.
- El programa tiene un botón por hoja. Si se recalcula solo `Calculo RE545`, el
  estado de la corrida **no habla** de `Calculo E Costos`, que viene de una
  corrida anterior.
- **Implementado**: `Ejecucion` escribe siempre `hojas_recalculadas`, y `PAG-001`
  alerta cuando una hoja del libro quedó vacía.

### `CTX-001` — Contaminación entre corridas de períodos distintos
- **Severidad:** CRÍTICA.
- Las salidas se escriben hoja por hoja y las que no se regeneran se conservan
  tal cual. Es posible terminar con `Medidores` de 2607 y `Subastas` de 2606 en
  el mismo libro, sin un solo error.
- **Parcialmente implementado**: `PER-001` detecta que la hoja `Medidores` traiga
  más de un mes. Falta que **cada hoja** estampe su período y el hash de sus
  entradas, y alertar cuando dos hojas del mismo libro no coincidan.

### `DST-001` — La hora repetida del cambio de horario
- **Severidad:** CRÍTICA.
- El catálogo cubre los bloques 92/100 en el CSV (`CSV-008`/`CSV-009`) pero no el
  problema de fondo: el día de retroceso tiene **dos veces la misma hora local**,
  así que `Central+Fecha+Hora+Cuarto` es ambigua para el cruce SoC→Medidores y
  `Barra+Cuarto`→CMg. `SOC-016` ("más de un SoC para una fila") se va a disparar
  en masa ese día.
- Hay que **definir la regla** (¿primera ocurrencia? ¿segunda?) antes de que
  aparezca en producción, y documentarla acá.

### `UNI-001` — Unidades y orden de magnitud
- **Severidad:** ALTA.
- El código tiene factores `/4*1000` (MW → kWh por cuarto de hora). Un control de
  rango por central (energía del cuarto ≤ `Pmax/4`, SoC entre 0 y capacidad)
  atrapa un error de escala que ninguna alerta de lookup puede ver. El catálogo
  solo toca escala en `SOC-014`.

### `SUP-001` — Los supuestos hardcodeados sobre la planilla origen
- **Severidad:** ALTA.
- Ya aparecieron dos trampas reales al replicar el Excel: `AR:AT`
  (`CPF(+)/CSF(+)/CTF(+)` del bloque FMA) son la **constante 1** y no un `SUMIFS`,
  y el ranking de `BK/BL/BS`. Son supuestos sobre la planilla origen, y hoy están
  escritos en el código sin nada que los verifique.
- Hace falta una aserción que se dispare cuando el archivo real deje de cumplir
  el supuesto. Si no, el día que cambien la planilla nadie se entera.

### `LOG-001` — Volumen de alertas
- **Severidad:** de diseño.
- "Registrar todos los casos" sobre ~100.000 filas × varios cruces puede dar
  cientos de miles de registros. Sin una regla, el catálogo es inaplicable o el
  log queda inusable.
- **Regla adoptada:** el detalle completo va al archivo (hoja `Alertas`, una fila
  por caso, **sin tope**); a la pantalla va un resumen agregado por tipo de
  control con hasta 15 ejemplos. Es lo que hace `_avisar_claves_sin_mapeo()`.

### `NUM-004` — Tolerancia y redondeo, escritos
- **Severidad:** de diseño.
- `TRA-007`/`TRA-008` piden conciliar "dentro de tolerancia" sin fijar el número.
  Sumando ~100.000 `float64` la igualdad exacta no sirve.
- **Regla adoptada:** relativa `1e-9`, piso absoluto `1e-6`, las dos escritas en
  la hoja `Ejecucion` de cada corrida. Falta todavía definir el redondeo del
  dinero.
