# BITACORA.md — registro de sesiones

Solo se agrega. Nunca se edita ni borra una entrada vieja. La única
excepción es la sección "Pendientes abiertos", que sí se edita porque es un
estado, no un historial.

---

## Pendientes abiertos

- Crear casos de prueba con datos reales y comparar la salida Python
  contra `11_PAGOS_BESS_2607_Definitivo.xlsm` (Medidores, Ofertas SSCC,
  CMg, FD, Subastas, Calculo E Costos) (plan §13 punto 10, §20.3). Todo se
  validó hasta ahora solo con casos sintéticos.
- Confirmar si la columna `V` (clave de homologación día+central, interna
  a `calcular_r`) necesita persistirse en una hoja propia para poder
  auditarla fila a fila contra la planilla 11, o si alcanza con auditar
  "Ofertas SSCC por Dia" + `Diccionario!E:F:G` a mano.
- ~~Bloqueante: confirmar la estructura real de columnas de `Subastas`~~
  — **resuelto**: el usuario entregó un `Libro1.xlsx` con la macro real y
  encabezados+fórmulas de la hoja `subastas` real. Confirmó que faltaba la
  columna `Concepto` (B) — ver entrada de esta sesión ("Corrección grande:
  `NOMBRES_SUBASTAS` estaba mal desde el principio"). `NOMBRES_SUBASTAS`
  ya está corregido y validado con los valores exactos de ese archivo.
- ~~Confirmar con el usuario qué son [Configuración, Ciclo, Clave, SUBIDA,
  BAJADA]~~ — resuelto en la misma entrega: viven en `Subastas!S:W` (antes
  se documentaba como `R:V`, un error menor de una letra). Coincide
  exactamente con lo que `construir_dic_umbrales_subastas()` ya calculaba
  en Python — sigue sirviendo como dato de validación cruzada si el
  usuario comparte valores reales de esa tabla para comparar.
- ~~Revisar con el usuario los valores reales de `Subastas!Sub_Baj`~~ —
  resuelto: `Sub_Baj` SÍ tiene `BAJADA`/`SUBIDA` reales, el dato nunca
  estuvo mal. Lo que estaba mal era `NOMBRES_SUBASTAS`: el nombre
  `"Sub_Baj"` estaba pegado a la posición equivocada (una posición antes
  de donde realmente vive), así que el código, sin saberlo, terminaba
  filtrando sobre los valores de `Control` (`CSF`/`CTF`/`CPF`, sin
  dirección) en vez de sobre `BAJADA`/`SUBIDA` reales. Eso explica
  exactamente el crash de la sesión anterior (0 filas tras el filtro
  `.isin(["BAJADA","SUBIDA"])` — nunca iba a haber match comparando contra
  texto que dice "CSF"). Falta re-confirmar con un caso real que ahora sí
  aparezcan filas `BAJADA`/`SUBIDA` con la corrección aplicada.
- ~~Validar contra un caso real que `Subastas!Control` tenga exactamente
  los valores `CPF`/`CSF`~~ — aclarado: `Control` (real) tiene el tipo SIN
  dirección (`CSF`/`CTF`/`CPF`, confirmado con el archivo real), y
  `construir_dic_prorrata()` ya buscaba por substring "cpf"/"csf" — sigue
  funcionando igual, sin cambios de código. Lo que SÍ cambió es que
  `construir_dic_reservas_subastas()` (reservas por subasta de `Calculo
  RE545`) usaba por error `Control` cuando necesitaba `Concepto` (las
  etiquetas completas `CPF(-)`/`CSF(+)`/etc que usa la fórmula real de
  `Subastas!$B:$B`) — corregido en esta sesión.
- Confirmar con más de un archivo de SoC real que el patrón "fila de
  nombre limpio + fila de ruta SCADA apiladas" (resuelto esta sesión,
  `detectar_fila_nombres()`) es estable. Con el único archivo real visto
  hasta ahora, los nombres de la fila 2 ya vienen idénticos a
  `Medidas_SAE.xlsx` — sospechar que el `Diccionario` quizás ni haga falta
  para el SoC, pero falta confirmarlo con otro período/archivo.
- Correr una vez el botón "Actualizar" de `Medidas_SAE.xlsx` contra la API
  real: confirmar la forma de la respuesta y que el `intervalo` de las dos
  APIs sea el inicio del cuarto de hora en las dos (de eso depende el cruce
  contra el calendario compartido).
- ~~Confirmar si la columna `Canal` de la hoja `Gen real` tiene que
  significar algo~~ — **resuelto**: sí. La API de operación real no expone
  canales, pero sí devuelve **MWh**, y todo el balance trabaja en **kWh**.
  Esa columna es ahora la **unidad** de la medida (`MWh`/`kWh`, vacía =
  MWh) y es lo que decide el factor de conversión.
- Confirmar si la columna `Flujo` de la hoja `Gen real` hace falta o si todas
  las centrales van con 1 (el script original no aplicaba signo).
- Abrir la ventana en Windows con la unidad `F:` conectada y confirmar que
  el botón "Traer FD" encuentra el FD en
  `…\Indicadores Publicar\<Vn>\03 Desempeño para publicar` y que los links
  de "Origen: ..." abren el explorador en esa carpeta (lo unico de esta
  sesion que no se puede probar sin la unidad).
- Abrir la ventana en Windows y confirmar los anchos de columna del
  diagrama (`ANCHO_ESTRUCTURA` 420 px, `ANCHO_ESTADO` 86 px,
  `ANCHO_ACCION` 104 px) con las fuentes reales de Windows (Consolas /
  Segoe UI): en Linux, medido con `xvfb`, las columnas quedan rectas y los
  botones todos de 96x22, pero las fuentes no son exactamente las mismas.
- Correr un caso REAL con la planilla combinada (`Balance_BESS.xlsx`) y
  confirmar que actualizar una hoja sola no toca ninguna de las otras
  nueve. Está cubierto por pruebas con casos sintéticos, no con un libro
  real de ~27.000 filas.
- Correr `traer_csv_cmg()`/`generar_cmg()` una vez contra el CSV real de
  `T:\CMgReales 15MIN`
  para confirmar que las barras de `Resumen BESS!Barra inyección` están
  escritas exactamente igual que la columna `BARRA` del CSV (con el relleno
  de guiones bajos, ej. `TOCOPILLA_____110`) y que el CSV real trae las 7
  columnas que dejan `BARRA` en D y `Cuarto de Hora` en H (ver
  `_validar_layout_cmg`).
- Validar la Prorrata SSCC normalizada (corregida esta sesión) contra más
  filas de la planilla 11 real — solo se confirmaron 2 casos puntuales
  (los que el usuario reportó), aunque el mecanismo (normalizar por fila)
  quedó confirmado con certeza matemática, no es una inferencia.
- ~~Investigar el reporte del usuario "En el R545 tengo diferencias igual
  parten en AK:AN"~~ — resuelto esta sesión: no era un problema de
  columnas AK:AN en sí (esas coinciden perfectamente comparando por
  posición/orden, no por letra de Excel — la salida nunca reprodujo la
  letra real, ver más abajo), sino que `AR:AT` (`CPF(+)`/`CSF(+)`/`CTF(+)`
  del bloque "FMA", el tercero de los tres bloques de reservas) se
  calculaban con el mismo `SUMIFS` que el resto, cuando en el archivo real
  son la **constante 1** en las 26.784 filas, no una fórmula. Ver entrada
  de esta sesión.
- Revisar si `construir_dic_mapeo_diccionario()` (columnas A:B, usada por
  el FD homologado de `Calculo E Costos!AM:AR`) y `_mapas_homologacion_
  fge()` (columnas E:F:G, usada por `Medidores!V`) también deberían usar
  `_bloques_columnas_diccionario()` en vez de índices de columna fijos
  (`0,1` y `4,5,6`) — hoy coinciden con los bloques reales por casualidad
  en el único `Diccionario` visto, pero si el orden de los bloques
  cambiara en otro archivo (o se agregara un bloque nuevo a la izquierda)
  esas posiciones fijas se romperían. No se tocó porque HOY funcionan
  bien y cambiarlas sin otro caso real de por medio sería un cambio
  especulativo.
- Correr un caso real completo (`Consolidado_entradas.xlsx` +
  `Pagos_BESS.xlsx`) con la corrección de `Subastas` aplicada, para
  confirmar que ahora sí aparecen filas `L=1` (participa en subasta) y
  que las Prorratas/reservas de RE545 no quedan todas en 0. Es el
  siguiente paso natural después de esta sesión.
- `Calculo E Costos`, `Calculo RE545`, `PRORRATA_RETIROS` y `Resumen` están
  implementadas. Queda pendiente el CSV (`Verificacion_CSV`) y validar las
  dos hojas nuevas contra el libro real completo de un período.
- Confirmar contra un caso real cuál de las columnas de `Subastas` suma
  cada bloque de reservas de `Calculo RE545` (`AC:AH`, `AI:AN`, `AO:AT`).
  Se siguió la fórmula (posición `O`/`P`/`Q`), pero los nombres reales de
  esas columnas (`FD`, `FMA`, sin nombre) están corridos una columna
  respecto de los títulos de grupo de RE545 (Subastas/FD/FMA). Ver plan
  §26.3.
- Decidir si la columna `Energía SSCC` de la hoja `Subastas` de
  `Consolidado_entradas.xlsx` tiene que quedar escrita ahí. El cálculo ya
  no es un pendiente (`calcular_subastas_energia_sscc()`), pero se hace
  del lado de `Pagos_BESS.xlsx`, que se genera después y en otro archivo;
  escribirla en la hoja `Subastas` implicaría que el botón "Generar
  Pagos_BESS" modifique el archivo de la otra ventana. Es una decisión de
  presentación, no de cálculo.
- Revisar `Subastas!M` ("Ciclo"). Aplicando el mismo corrimiento de una
  columna que el usuario confirmó (y que se usó para `L`, `AW` y
  `Subastas!N`), la fórmula original `=K&H&I` sería
  `Configuración & Dia & Hora_dia`, no `Propietario & Hora_dia &
  Hora_mes` como está hoy. Hoy no afecta ningún cálculo (ninguna otra
  columna consume `Ciclo`), por eso no se tocó sin preguntar.
- ~~Confirmar si la carpeta `Subastas/` (creada en su momento, no existe
  en la planilla original) es el nombre/ubicación que se quiere
  mantener~~ — resuelto: se queda, y además ahora tiene adentro la
  subcarpeta `DB subastas/` (nombre elegido por el usuario) con los
  Access, que son el origen real.
- **Confirmar de dónde sale `Energía SSCC`** (`Subastas!O`, antes
  `DB!P`): hoy se toma `CANTIDAD PONDERADA MW` del Access
  (`COLUMNA_ENERGIA_SSCC_ACCDB`), que es la `Quantity2`. La otra
  candidata es `CANTIDAD MW` (`Quantity`). Con los datos reales vistos
  no se pueden distinguir (las dos coinciden en casi todas las filas,
  porque `entradas_sscc.py` completa la ponderada vacía con la cruda).
  Es un cambio de una línea si el usuario confirma la otra.
- ~~`FMA` de `Subastas` (`Q`) pendiente~~ — **resuelto** esta sesión:
  sale de las tres salidas de FMA (`fma_cpf_*`, `fma_csf_*`,
  `fma_cft_*`) que ahora se guardan en `FD y FMA/`, siguiendo el
  documento de trazabilidad del usuario. Falta validarlo contra un
  `DB!V` real.
- ~~`FD` de `Subastas` (`P`) pendiente~~ — **resuelto**: sale del
  `SSCC_Desempeño_*`, hoja por familia, cruzando Control + Unidad +
  Hora_mes (`calcular_fd_subastas()`). Falta validarlo contra un `DB!Y`
  real.
- ~~Vector de Participación CSF~~ — **resuelto**: es el "Indicador de
  participación" que se deriva de la columna `Respuesta CSF` de la hoja
  `CSF Horario` del mismo archivo. Ya multiplica al FMA de las filas
  CSF. Falta validarlo contra un `DB!AC` real.
- **Agregar/confirmar el bloque `FD` de la hoja `Diccionario`**:
  `Configuración` → unidad como la nombra el `SSCC_Desempeño_*`. Ese
  bloque ya existía (es el que usa `Calculo E Costos!AM:AR`) y ahora lo
  usa también `Subastas!FD`; hay que revisar que cubra todas las
  centrales que aparecen en subastas, porque las que falten se buscan
  con el nombre tal cual y quedan avisadas en el log.
- **Agregar el bloque `FMA CPF` a la hoja `Diccionario` de
  `Centrales.xlsx`**: dos columnas, `Configuración` → central como la
  nombra `fma_cpf` (ej. `SAE-DEL-DESIERTO` → `BESS DEL DESIERTO`). Es
  la equivalencia que en la planilla 3 vivía en `FMA_CPF!AD:AE` y que
  el documento de trazabilidad recomienda mudar acá. Mientras no esté,
  el FMA de CPF se busca con el nombre tal cual y se avisa en el log.
- **Confirmar el comportamiento de `CTF(+)`**: la fórmula real de
  `DB!V` tiene dos bloques consecutivos para `CTF(+)` y, por cómo están
  anidados los `SI`, el segundo nunca se evalúa — así que `CTF(+)`
  busca solo en la primera tabla mientras que `CTF(-)` busca en las
  dos. Se replicó ese comportamiento tal cual (el documento lo pide
  expresamente) y quedó en una constante,
  `CTF_MAS_BUSCA_EN_LAS_DOS_TABLAS = False`, para poder comparar
  después contra la versión corregida. Con centrales BESS es casi
  seguro que no cambia nada (no hacen CTF).
- **Definir el ajuste de `Hora_mes` por cambio de hora.** La fórmula
  real de la planilla 3 es `=($F9-1)*24+$G9+IF(F9>$F$2,1,0)`, donde
  `SUBASTAS!F2` es el día del mes en que cambia la hora. Hoy
  `calcular_hora_mes_subastas()` tiene ese ajuste como parámetro
  (`dia_cambio_hora`) y lo deja **apagado**, que es lo correcto en los
  10 meses del año sin cambio de hora, pero falta definir de dónde sale
  ese día en los otros dos (el `archivo_de_configuracion.yaml` del
  script viejo tiene un `cambio_de_hora: -1` en las variables
  mensuales, pero no está claro si es un día o un signo). Mientras
  tanto, el log avisa si algún día del mes no trae 24 horas.
- ~~Confirmar dos rarezas de la propuesta nueva del `Diccionario`~~ —
  resueltas por el usuario: el `FMA_CPF` cruzado entre Andes 3 y Andes 4
  era un error suyo y ya lo corrigió en su archivo; que Nuevo Quillagua 2
  y Víctor Jara no tengan `FMA_CPF` **está bien**, y pidió que se tolere
  pero se advierta igual (hecho: el aviso ahora dice cuántas centrales
  son, cuáles, y que su FMA CPF queda en 0).
- **Confirmar la homologación FD de `SAE-CRCA-PFV-NUEVO-QUILLAGUA-2`**: el
  `Diccionario` la mapea a sí misma y la hoja `FD` no tiene ninguna unidad
  con ese nombre. Pero sí trae una unidad que **ninguna** central reclama:
  `BESS PFV MARIA ELENA`. Muy probablemente sean la misma central (la
  barra de Nuevo Quillagua 2 es `PEQ___________220`). Si lo es, hay que
  escribirlo en la columna `FD` del `Diccionario` y desaparecen 1.488 de
  las 2.232 alertas.
- **Confirmar que `SAE-CRCA-PFV-ANDES4` no presta CSF**: su unidad
  (`BESS PFV ANDES SOLAR IV`) está en el bloque CPF de la hoja `FD` pero
  no en el CSF, así que su FD CSF queda en 0 (las 744 alertas restantes).
  Si es correcto, no hay nada que arreglar.
- Medir contra un caso real cuánto baja la generación del FMA CPF con el
  índice de carpeta nuevo (antes: >6 minutos, el usuario la tuvo que
  cortar). En sintético se comprobó que el árbol se recorre 1 vez en vez
  de 31 por versión, pero falta el número real sobre la unidad de red.
- **FMA CSF en cero para algunas centrales** (reporte del usuario: "el FMA
  de CSF es horario y sistémico, es raro que algunas centrales queden en
  cero"). NO es la homologación: se midió contra el
  `Consolidado_entradas.xlsx` real que mandó el usuario y las 9
  `Configuración` de `Subastas` ya cruzaban 9/9 contra el bloque `FD` del
  `Diccionario` viejo. El sospechoso que queda es el **Vector de
  Participación CSF** (`_participacion_csf()` en
  `Script/Fd/Desempeno_Horario.py`), que multiplica al FMA base y vale
  **0** cuando el `SSCC_Desempeño_*` dice "No Participó" para esa unidad en
  esa hora (y su alternativa TG/TV tampoco). Si eso es correcto, el cero
  es el resultado esperado y no hay nada que arreglar. Falta que el
  usuario tome una central+hora con FMA CSF en cero y la mire en la hoja
  `CSF Horario` del `SSCC_Desempeño_*` para confirmarlo.

- Confirmar el nombre definitivo de `Pagos_BESS.xlsx` (provisorio, elegido
  por el usuario como "pagos_bess o algo así por ahora").
- Probar la ventana nueva (diagrama + botones "Generar") con una carpeta
  base real: solo se probó por ahora con `python -m py_compile` (no hay
  entorno grafico en esta sesión para abrir la ventana) y con pruebas
  sintéticas de la lógica de árbol (`_prefijos_arbol`) y de generación
  parcial (`generar_consolidado`, `generar_pagos_bess`) por separado.
- Evaluar si `guardar_config()` necesita escritura atómica (ver
  `METODOLOGIA.md` §7).

---

## 2026-09-10 — Organización inicial del repositorio

Se recibieron los archivos de la etapa Medidores (`Balance_BESS.py`,
`nucleo.py`), el plan de traspaso a Python y una plantilla de metodología.
El repositorio estaba prácticamente vacío (solo un `README.md` de una
línea). Se ordenó siguiendo la plantilla de metodología recibida:

- `Balance_BESS.py` y `nucleo.py` agregados en la raíz del repositorio, sin
  cambios respecto de lo recibido.
- Plan de traspaso agregado como documento de dominio en
  `docs/Plan_Traspaso_Python_Balance_BESS.md`.
- `METODOLOGIA.md` adaptado con la información real del proyecto
  (secciones 1, 2, 4, 5, 7 y 8 completadas; ya no es la plantilla genérica).
- Creados `MAPA.md` (un bloque por script), `BITACORA.md` (este archivo),
  `REGLAS.md` (checklist de inicio/cierre de sesión), `requirements.txt`
  (`pandas`, `openpyxl`) y `.gitignore` (`config.json`, `__pycache__/`,
  salidas `.xlsx` de casos concretos).
- `README.md` actualizado con instalación, uso y tabla de navegación hacia
  el resto de los documentos.

Estado dejado: solo está implementada la etapa Medidores, y dentro de ella
solo las columnas A:J, L, N y O. Las columnas K, M, P, Q, R, S, T quedan
como `pd.NA` (pendientes, ver arriba). No hay implementación de etapas
posteriores del balance ni de `OfertasSSCC`. No se corrió el proceso contra
un caso real en esta sesión (no había datos de un caso disponibles); queda
pendiente crear casos de prueba para validar contra la planilla 11.

---

## 2026-09-10 (2) — Plan actualizado: AAMM manual y especificación cerrada de columnas

Se recibió una versión actualizada del plan de traspaso
(`docs/Plan_Traspaso_Python_Balance_BESS.md`, secciones 16–18 nuevas) más un
pedido explícito: el archivo de SoC no se llama literalmente
`SOC_AAMM.xlsx`; el período AAMM debe ingresarlo el usuario en un recuadro
de la ventana, no inferirse del nombre de un archivo.

Cambios en `nucleo.py`:

- `buscar_soc(medidas_dir, aamm)` ahora recibe el AAMM como parámetro
  (validado por la nueva `validar_aamm()`) y busca cualquier `.xlsx` dentro
  de `Medidas/` cuyo nombre contenga "SOC" y ese AAMM, en vez de exigir el
  patrón literal `SOC_AAMM.xlsx`. Se probó con un nombre deliberadamente no
  literal (`"resumen soc julio 2607 v2.xlsx"`) y detecta correctamente.
- `revisar_estructura(carpeta_base, aamm)` y `ejecutar(carpeta_base, aamm, ...)`
  ahora reciben el AAMM en vez de extraerlo del archivo de SoC. El checklist
  agrega una fila "Periodo (AAMM)" que bloquea Ejecutar si no son 4 dígitos.
- Se agregó `CARPETA_OFERTAS = "Ofertas"` y `ofertas_dir` en
  `resolver_rutas()`, con una fila no bloqueante en el checklist (ubicación
  ya definida por el plan §16.2; falta patrón de archivo y lectura).
- `LETRA_A_CAMPO` se extendió de A a AE siguiendo la tabla cerrada del plan
  (§16.3). Se corrigió el orden final de columnas para que use
  `list(LETRA_A_CAMPO.values())` en vez de `sorted(LETRA_A_CAMPO)`: con
  claves de dos letras ("AA", "AB"...) el orden alfabético de texto ya no
  coincide con el orden real de columnas de Excel.
- Se implementó `K` = copia de `L` fila a fila (`Copia_Ventana`).
- Se separaron `COLUMNAS_VACIAS` (M, P, Q, U, Z, AA — diseño confirmado,
  no pendiente) de `COLUMNAS_PENDIENTES_OFERTAS` (R, S, T, V, W, X, Y, AB,
  AC, AD, AE — dependen de las macros de Ofertas SSCC).

Cambios en `Balance_BESS.py`:

- Nuevo campo "Periodo del caso (AAMM)" en la ventana (`Entry` limitado a
  4 dígitos), con su propio `StringVar` persistido en `config.json`.
  Cambiar el AAMM (Enter o al perder foco) vuelve a correr la validación
  de estructura.
- `revisar()` y la llamada a `nucleo.ejecutar()` pasan el AAMM ingresado.

**No se implementaron** `R, S, T, V, W, X, Y, AB, AC, AD, AE`: el plan
actualizado (§17, §18) exige que las macros de Ofertas SSCC se repliquen
fielmente, pero esta sesión solo tiene el nombre de esas macros y qué
columna produce cada una — no su código VBA. Implementarlas sin eso sería
adivinar la lógica, lo que el propio plan prohíbe (§18: "no reemplazar
lógica conocida por placeholders" — y en este caso la lógica todavía no es
conocida por el asistente). Quedan documentadas en
`docs/Plan_Traspaso_Python_Balance_BESS.md` §19.2 y en "Pendientes
abiertos" arriba.

Verificación: se armó un caso sintético en el scratchpad (carpeta con
`Medidas_SAE.xlsx`, un archivo de SoC con nombre no literal, y
`Centrales.xlsx`) y se corrió `nucleo.ejecutar()` de punta a punta: genera
31 columnas en el orden correcto, `K` coincide con `L`, y las columnas
vacías/pendientes quedan como `NaN`. No se probó contra un caso real ni
contra la planilla 11 (sigue sin datos reales disponibles en el entorno).

---

## 2026-09-10 (3) — Ofertas SSCC implementado a partir del código VBA fuente

Se recibió `Trazabilidad_11_PAGOS_BESS_2607_Definitivo.md` con el código VBA completo de
`Generar_Resumen_Ofertas_SSCC` y `Resumir_Medidores_Central_Ventana_Oferta_Completa`, y las
fórmulas de Excel de `Medidores!K,L,N,O,R,S,T,V`. Esto resuelve el pendiente principal de la
sesión anterior.

**Hallazgo importante antes de programar:** las fórmulas muestran que `V` (y por macro, `W, X,
Y, AB, AC, AD, AE`) NO son columnas por fila de `Medidores` — son tablas auxiliares de otro
largo (central × día, central × ventana) que solo viven en esas letras de columna porque ahí
había espacio libre en la planilla. Mantenerlas como columnas `pd.NA` del mismo largo que A:U
(como se hizo la sesión anterior, cuando eran genuinamente "pendientes") dejó de tener sentido
una vez que se pueden calcular: ahora se escriben como hojas propias de `Hoja_Medidas.xlsx`
("Resumen Ofertas SSCC", "Ofertas SSCC por Dia", "Resumen Ventana Oferta"). Esto está
documentado con más detalle en `docs/Plan_Traspaso_Python_Balance_BESS.md` §20.

Cambios en `nucleo.py`:

- `LETRA_A_CAMPO` ahora va de A a U (se sacaron V, W, X, Y, AB, AC, AD, AE). `COLUMNAS_VACIAS`
  quedó en M, P, Q, U. Se eliminó `COLUMNAS_PENDIENTES_OFERTAS` (ya no queda nada pendiente).
- Nuevas funciones que replican la macro `Generar_Resumen_Ofertas_SSCC`:
  `construir_resumen_ofertas_sscc()` y sus auxiliares privados (`_contiene_bess_o_sae`,
  `_servicio_termina_en_rs`, `_es_respuesta_si`, `_normalizar_periodo`, etc., prefijo `_` como
  las de detección de SoC).
- `cargar_resumen_en_medidores()` replica `OSSCC_CargarResumenEnMedidores` (equivalente a
  `Medidores!W:Y`), incluyendo la homologación vía `Diccionario!E:F:G` y el aviso de nombres de
  `Medidores!clave` no encontrados en el diccionario.
- `calcular_r()` replica la fórmula de `R` (`VLOOKUP` contra la tabla `V:Y`), usando
  `_mapas_homologacion_fge()`/`_homologar_fge()` para la homologación específica vía
  `Diccionario!F/G→E` que usa la fórmula de `V` — un mapeo DISTINTO del que usa
  `construir_homologacion()` para el SoC (una es posicional por columna, la otra trata toda la
  fila como equivalencias simétricas). Si no hay match, `R` queda `NaN` y se registra un aviso
  en vez de fallar (no debería pasar si `Diccionario` está completo, pero no se asume).
- `calcular_s()` replica la fórmula de `S`, vectorizada por "corridas" de `Ventana` constante
  (no se reinicia por central, igual que la fórmula original).
- `construir_resumen_ventana_oferta()` replica `Resumir_Medidores_Central_Ventana_Oferta_
  Completa` (equivalente a `Medidores!AB:AE`).
- `calcular_t()` replica la fórmula de `T` (`1 - Completa`), vía merge contra el resumen
  anterior.
- `buscar_archivo_ofertas()` busca el archivo `*OfertasSSCC*` más reciente en `Ofertas/` — a
  diferencia del SoC, si hay más de uno SÍ se elige automáticamente por fecha de modificación
  (así lo hace la macro `OSSCC_BuscarArchivoOfertas` original).
- `revisar_estructura()`: la fila de Ofertas SSCC pasó de `pendiente` (no bloqueaba) a `falta`
  (bloquea Ejecutar) si no hay carpeta `Ofertas/` o no hay archivo `*OfertasSSCC*` — ahora es
  obligatoria (plan §17-18).
- `construir_medidores()` y `ejecutar()` quedaron con nuevos parámetros (`ruta_ofertas`,
  `diccionario`) y devuelven además las tres tablas auxiliares. `escribir_salida()` las escribe
  como hojas nuevas.

**Trampa encontrada y corregida en la misma sesión:** el primer intento de
`buscar_archivo_ofertas()` comparaba contra el literal `"ofertasscc"` (dos "s" seguidas), pero
`"OfertasSSCC".lower()` da tres "s" seguidas ("Ofertas" + "SSCC"). El patrón ahora se deriva en
tiempo de ejecución con `"OfertasSSCC".lower()` (constante `PATRON_NOMBRE_OFERTAS`) en vez de
transcribirlo a mano, para no repetir el error. Ver `METODOLOGIA.md` §7 si se agrega ahí.

**Verificación:** caso sintético con 1 central, 2 días, `Hora` en convención 1-24 (como usa la
planilla real, confirmado porque `(INICIO_VENTANA-1)*4=36` y `(25-INICIO_VENTANA)*4=60` dieron
exactamente los conteos de filas de las ventanas de inicio y de cierre del caso de prueba) y un
archivo de Ofertas SSCC con oferta completa las 24 horas ambos días: las 3 ventanas (inicio,
normal, última) quedaron `Completa=1` con 36/96/60 filas respectivamente, y R=1/T=0 en las 192
filas de `Medidores`. Un segundo caso con una hora "No" en vez de "Sí" marcó correctamente
`Oferta completa=0`. No se probó contra un caso real ni contra la hoja `Medidores` de la
planilla 11 (sigue pendiente, ver arriba).

---

## 2026-09-10 (4) — El archivo de SoC puede ser .csv, no solo .xlsx

El usuario aclaró que `SOC_AAMM` (el archivo de SoC) puede llegar como `.csv`, manteniendo la
misma estructura de bloques horizontales por central (`Status | Questionable | Time Stamp |
Value`) que ya soportaba `extraer_soc()`. Antes de tocar código se preguntó explícitamente por
la estructura del CSV (¿tabla larga o mantiene los bloques?) para no adivinar mal un formato
que ya funcionaba: el usuario confirmó que mantiene los bloques, solo cambia el contenedor.

Cambios en `nucleo.py`:

- `EXTENSIONES_SOC = {".xlsx", ".csv"}` (antes `buscar_soc()` exigía `.xlsx` a secas).
- Nueva función `leer_soc_crudo(ruta_soc)`: elige `pd.read_csv(ruta_soc, header=None)` o
  `pd.read_excel(ruta_soc, sheet_name=0, header=None)` según la extensión.
  `extraer_soc()` la usa en vez de llamar a `pd.read_excel` directo.
  `detectar_fila_nombres()`/`detectar_bloques()` no cambiaron: ya trabajan sobre el DataFrame
  resultante sin que les importe de dónde vino.

Probado con un CSV sintético (mismo layout de bloque que ya se probaba en Excel, separador
coma, UTF-8): `extraer_soc()` lo lee igual que un `.xlsx` equivalente, sin incidencias.

No se tocó el formato del archivo de OfertasSSCC (sigue siendo Excel, como confirma el propio
VBA que filtra por extensiones `.xlsx/.xlsm/.xlsb/.xls`) ni el de `Medidas_SAE.xlsx`.

---

## 2026-09-10 (5) — Corrección: el archivo de SoC no era .csv, era otro archivo

El usuario subió un `SOC_2607.csv` real para probar el soporte agregado en la entrada anterior,
y falló: `extraer_soc()` no encontró los encabezados `Time Stamp`/`Value`. Al inspeccionar el
archivo, resultó ser una tabla larga con columnas `Fecha_Hora, CONFIGURACION, Central, Pago,
Tipo_pago, Bloque_15min` — un archivo de pagos/liquidación, sin ninguna columna de SoC, que
solo coincidía por casualidad con el patrón de nombre "SOC"+AAMM.

Se preguntó al usuario antes de tocar más código (no se adaptó el parser a esta estructura
nueva a ciegas). Confirmó: era el archivo equivocado, "dejalo como antes es un xlsx" — el
archivo de SoC real siempre es `.xlsx`, con la estructura de bloques horizontales que
`extraer_soc()` ya soportaba desde el principio.

Se revirtió el soporte de `.csv` para el SoC agregado en la entrada anterior:

- `nucleo.buscar_soc()` vuelve a exigir `.xlsx` exclusivamente.
- Se eliminó `nucleo.leer_soc_crudo()`; `extraer_soc()` vuelve a llamar `pd.read_excel`
  directamente, como antes de esa sesión.
- Se revirtieron las menciones a "SoC puede ser .csv" en `README.md`, `MAPA.md` y
  `METODOLOGIA.md` §5 (documentos de estado actual, no de historial).
- Se agregó una trampa en `METODOLOGIA.md` §7: un archivo que matchea el patrón de nombre
  "SOC"+AAMM pero no tiene la estructura de bloques esperada no es el archivo de SoC, aunque
  comparta el nombre — no adaptar el parser a ciegas, preguntar primero.
- `docs/Plan_Traspaso_Python_Balance_BESS.md` sección 21 se corrigió en el mismo lugar (no se
  duplicó una sección nueva) para reflejar que el archivo sigue siendo siempre `.xlsx`.

Esta entrada de bitácora no borra ni edita la entrada anterior (2026-09-10 (4)) según la regla
de solo-agregar; queda como registro de que esa sesión partió de una premisa equivocada y esta
la corrigió.

---

## 2026-09-10 (6) — Hojas auxiliares de Ofertas SSCC: se unen y se recorta una

Pedido del usuario: las tablas "Ofertas SSCC por Dia" y "Resumen Ventana Oferta" deben quedar
en la misma hoja, y "Resumen Ofertas SSCC" es auxiliar — mejor no crearla como hoja.

Cambios en `nucleo.py`:

- `construir_resumen_ofertas_sscc()` sigue calculándose igual dentro de `construir_medidores()`
  (sin cambios en su lógica), pero deja de devolverse/persistirse: es un paso intermedio que
  solo hace falta en memoria para construir la tabla equivalente a `Medidores!W:Y`.
  `construir_medidores()` ahora devuelve `(df_medidores, avisos, df_wxy, df_resumen_ventana)`
  (antes devolvía también `df_resumen_ofertas`).
- Nueva función `_escribir_tabla_con_titulo(writer, hoja, df, titulo, fila_inicio)`: escribe un
  título en negrita y la tabla debajo, dentro de una hoja dada, y devuelve la fila donde debería
  empezar el siguiente bloque (para poder apilar varias tablas en la misma hoja).
- `escribir_salida()` ya no recibe `df_resumen_ofertas`; ahora escribe `df_wxy` y
  `df_resumen_ventana` una debajo de la otra en una sola hoja nueva, `HOJA_OFERTAS_SSCC =
  "Ofertas SSCC"`, cada una con su título.
- `Hoja_Medidas.xlsx` queda con tres hojas: `Medidores`, `Ofertas SSCC`, `Log` (antes tenía
  cinco).

Probado con el mismo caso sintético de sesiones anteriores: la hoja combinada queda con el
título+tabla de "Ofertas SSCC por dia" (31 filas) seguido de una fila en blanco y el
título+tabla de "Resumen ventana oferta" (3 filas), en el orden y con los valores esperados.

---

## 2026-09-10 (7) — Hojas CMg, FD, Subastas + renombre del archivo de salida

Pedido del usuario: agregar las hojas CMg, FD y Subastas, replicando las macros
`Cargar_CMg_Desde_Archivo`, `Cargar_SSCC_Desempeno_En_FD` y `Cargar_Remuneracion_Subastas_Rapido`
del documento de trazabilidad VBA, y renombrar el archivo de salida de `Hoja_Medidas.xlsx` a
`Consolidado_entradas.xlsx`.

Se leyó el código VBA completo de las tres macros y de sus funciones auxiliares
(`UltimaFilaEntreColumnasCMg`, `FiltrarFilasBESSoSAE`/`EsBESSoSAE`, `AjustarBloqueFormulas`,
`BuscarArchivoSSCCMasReciente`, `AbrirConexionExcelSubastas`/`AjustarFormulasMNSubastasRapido`/
`BuscarArchivoSubastasMasRecienteRapido`), más la sección de fórmulas del libro (`5.2 FD`,
`5.3 Subastas`) del documento de trazabilidad, ya usado en sesiones anteriores para Ofertas SSCC.

**Hallazgo estructural (igual patrón que V:Y/AB:AE de Medidores, pero por columnas):** en `FD`,
el bloque CSF (A:M, viene de `CSF Horario`) y el bloque CPF (Q:AE, viene de `CPF Horario`) son
dos tablas independientes de distinto largo que comparten la hoja en rangos de columnas
distintos, no de filas. Se escriben lado a lado (`escribir_salida()` usa `startcol` en
`df.to_excel()`), cada una con su propio número de filas.

Cambios en `nucleo.py`:

- Nuevas carpetas/archivos: `CARPETA_CMG="Cmg"` (+ `ARCHIVO_CMG="cmg.xlsx"`, nombre literal fijo,
  a diferencia de todos los demás archivos externos del proyecto), `CARPETA_SSCC_DESEMPENO=
  "SSCC_Desempeño"`, `CARPETA_SUBASTAS="Subastas"`.
- `_buscar_archivo_excel_mas_reciente()`: generaliza `buscar_archivo_ofertas()` (antes solo
  servía para Ofertas) para reusarse también en `buscar_archivo_sscc_desempeno()` y
  `buscar_archivo_subastas()`, con un flag `desde_inicio` porque Ofertas busca el patrón en
  cualquier posición del nombre, mientras que SSCC_Desempeño y Subastas exigen que el nombre
  *empiece* con el patrón (así lo hacen sus macros originales, con `Dir("patron*.*")`).
  `EXTENSIONES_OFERTAS` se renombra a `EXTENSIONES_EXCEL` (ya no es solo de Ofertas).
- `leer_cmg()`: replica `Cargar_CMg_Desde_Archivo` — hoja `CMg` o la primera si no existe,
  columnas A:I con su encabezado real (no se inventan nombres), ordenadas por columna D
  ascendente y luego H ascendente.
- `construir_fd()` + `_construir_bloque_fd_csf()`/`_construir_bloque_fd_cpf()`: replican
  `Cargar_SSCC_Desempeno_En_FD` y las fórmulas de FD (`A=str(B)&F`, `B=(DAY(D)-1)*24+E+1+
  IF(DAY(D)>100,1,0)`, `C=DAY(D)`, `K=J`, `L=K`, `M=B` para el bloque CSF; análogo con
  Q,R,S,T,U,V,AA,AC,AD,AE para el bloque CPF). El término `IF(DAY(fecha)>100,...)` se conserva
  tal cual aunque nunca sea cierto para un día real (fiel a la fórmula original, no se "limpia").
  Filtro BESS/SAE **sin** BAT (`_contiene_bess_o_sae_sin_bat()`, distinta de la de Ofertas SSCC
  que sí incluye BAT — son dos filtros reales distintos, no se fusionaron).
- `construir_subastas()`: replica `Cargar_Remuneracion_Subastas_Rapido` (que en VBA usa ADO/SQL
  contra la hoja `DB`; en Python se lee directo con pandas aplicando el mismo filtro/selección,
  sin necesitar ADO). Arma B:L (copia), M (fórmula `=K&H&I`), O/P/Q (copias de P/Y/V). `N` queda
  `pd.NA` documentada como pendiente: su fórmula real depende de `'Calculo E Costos'`, una hoja
  de una etapa posterior sin implementar — no se adivina.
- `revisar_estructura()`, `ejecutar()`, `escribir_salida()`: extendidos para validar, leer y
  escribir las tres hojas nuevas. Las tres entradas son obligatorias (bloquean Ejecutar si
  faltan), igual criterio que Ofertas SSCC.
- `ARCHIVO_SALIDA` cambia de `"Hoja_Medidas.xlsx"` a `"Consolidado_entradas.xlsx"`.

**Decisión de arquitectura (a confirmar con el usuario, ver "Pendientes abiertos"):** la macro
original de Subastas buscaba su archivo directamente en la carpeta del `.xlsm`, sin subcarpeta.
Se le creó una carpeta propia `Subastas/` para ser consistente con el resto de las entradas
externas de este proyecto (cada una con su carpeta bajo la carpeta base del caso), no porque la
planilla original lo hiciera así.

**Verificación:** caso sintético con datos para las 6 entradas (Medidas_SAE, SoC, Centrales,
OfertasSSCC, cmg.xlsx, SSCC_Desempeño_\*, 3_REMUNERACIÓN_SUBASTAS_E_ID_\*). Se verificaron a
mano los valores esperados de las fórmulas de FD (A, B, C, K, L, M del bloque CSF; Q, R, S, AC,
AD, AE del bloque CPF) y de Subastas (M, O, P, Q) contra los datos de entrada armados a
propósito, y coincidieron exactamente. `Consolidado_entradas.xlsx` quedó con las 6 hojas
esperadas: `Medidores`, `Ofertas SSCC`, `CMg`, `FD`, `Subastas`, `Log`. No se probó contra un
caso real ni contra la planilla 11.

---

## 2026-09-10 (8) — Encabezados reales de FD/Subastas + Ofertas SSCC lado a lado

El usuario entregó un Excel (`Libro1.xlsx`) con los encabezados reales de `FD` y `Subastas`, y
pidió que las dos tablas de la hoja "Ofertas SSCC" (sesión anterior, apiladas verticalmente)
queden una al lado de la otra en vez de una debajo de la otra.

**Encabezados FD/Subastas (plan §24):** confirman exactamente las fórmulas ya implementadas la
sesión anterior — nada cambió en los VALORES calculados, solo se reemplazaron los nombres de
columna por letra (`"A"`, `"B"`, ...) por los nombres reales (`NOMBRES_FD_CSF`, `NOMBRES_FD_CPF`,
`NOMBRES_SUBASTAS` en `nucleo.py`). Detalle interesante: los encabezados reales confirman que
`FD!M` y `FD!AE` repiten literalmente el nombre "Hora Mes" de `FD!B`/`FD!R` — coincide con que
también repiten su valor (`M=B`, `AE=R`), así que no hay contradicción. Como Python no permite
indexar sin ambigüedad un DataFrame con nombres de columna duplicados, el renombre se aplica con
`set_axis()` al final de cada función (`_construir_bloque_fd_csf`/`_construir_bloque_fd_cpf`),
después de terminar todos los cálculos con nombres de letra únicos — nunca antes.

También se detectó que `Subastas!K` (la columna que se filtra por BESS/SAE) se llama
"Propietario" en la vida real, y `Subastas!M` ("Ciclo") es literalmente
`Propietario & Hora_dia & Hora_mes` — nombres mucho más claros que "K", "H", "I". `Subastas!Q`
no tiene encabezado en el archivo real: se dejó sin nombre (columna `""`), no se le inventó uno.

**Ofertas SSCC lado a lado:** `_escribir_tabla_con_titulo()` ahora acepta `columna_inicio`
además de `fila_inicio`, y devuelve `(fila_siguiente, columna_siguiente)` en vez de solo la
fila — cada llamada usa el que corresponda según cómo se estén acomodando los bloques
(`escribir_salida()` ahora encadena por columna para esta hoja en particular).

Probado con el mismo caso sintético de sesiones anteriores: los encabezados de `FD` y
`Subastas` en el archivo generado coinciden letra por letra con los del `Libro1.xlsx` entregado,
y la hoja "Ofertas SSCC" quedó con "Ofertas SSCC por dia" en A1:C... y "Resumen ventana oferta"
arrancando 2 columnas después (F1:I...), ambas en la fila 1.

---

## 2026-09-11 — `Calculo E Costos`: etapa base (H + CMg + traspaso de Medidores)

El usuario pidió crear la hoja `Calculo E Costos` ("Ecostos"), cuya lógica sale de la macro
`Actualizar_Calculos_Columnas` (columna H de barras por fórmula, CMg asignado por
`Asignar_CMg_a_Calculos_Turbo`), en un archivo **separado** de `Consolidado_entradas.xlsx`
("pagos_bess o algo así por ahora"). El archivo de encabezados reales que el usuario intentó
enviar para esta hoja llegó dos veces sin la hoja `Ecostos` (solo traía `FD`/`Subastas`, ya
confirmadas la sesión anterior) — queda pendiente, ver "Pendientes abiertos".

`Actualizar_Calculos_Columnas` resultó ser una macro de ~1500 líneas con dependencias profundas
(Subastas, Resumen, Diccionario, FD, y una "Prorrata SSCC" que en el documento de trazabilidad
figuraba como fuente externa pendiente). Antes de traducir todo de una vez, se preguntó al
usuario: (a) si la Prorrata SSCC era un archivo externo disponible, y (b) cómo priorizar el
trabajo dado el tamaño de la macro. Respuestas: la Prorrata SSCC **no es un archivo externo**,
es una tabla dinámica derivable de `Subastas` (`Filas: Configuración, Hora_mes` / `Columnas:
Control` / `Valores: Cuenta de Sub_Baj`) — no es un bloqueador real. Y la prioridad elegida fue
explícitamente **"por etapas: primero H + CMg + traspaso de Medidores"**.

Esta sesión implementa exactamente esa primera etapa (plan §25):

- `_normaliza_cuarto()`: replica `NormalizaCuarto` (numérico → texto del entero redondeado,
  otro → texto recortado, vacío/error → `""`).
- `construir_dic_cmg(df_cmg)`: replica el armado del diccionario de
  `Asignar_CMg_a_Calculos_Turbo` a partir de `df_cmg` (columnas por posición, ya sin renombrar
  desde `leer_cmg()`: D=Barra, F=valor a asignar en Q, H=Cuarto de Hora). Clave
  `UCase(Barra)+"|"+NormalizaCuarto(CuartoHora)`; ante clave repetida gana la primera fila.
- `construir_mapa_barra(resumen_bess)`: resuelve `H` (antes fórmula
  `=VLOOKUP(G,Resumen!B:G,6,FALSE)`) homologando **por nombre de columna**
  (`Nombre activo`/`Barra inyección` de `Resumen BESS`) en vez de por posición, porque
  `Centrales.xlsx` no reproduce el layout `Resumen!B:G` del libro original. `leer_centrales()`
  ya devolvía `resumen` pero `ejecutar()` lo descartaba (`_, diccionario = ...`); ahora se
  captura y se usa.
- `construir_calculo_e_costos(df_medidores, mapa_barra, dic_cmg, registrar=print)`: arma A:G
  (con D↔E invertidas, igual que `Traspasar_Medidores_A_Calculos_Rapido`), `Barra` (H),
  `Energia_Positiva`/`Energia_Negativa` (I/J, energía de `Medidores!Gen_Unidad` separada por
  signo, solo si `Ventana_No_Completa=1`), `SoC` (K, copia de `Medidores!SoC`),
  `Copia_Ventana` (P, copia de `Medidores!Copia_Ventana`) y `CMg` (Q, homologado por Barra +
  Cuarto de Hora). Nombres de columna: placeholders, ver "Pendientes abiertos".
- `escribir_pagos_bess(ruta_salida, df_ecostos, registrar=print)`: escribe `Pagos_BESS.xlsx`
  (nueva constante `ARCHIVO_SALIDA_PAGOS`) con la única hoja `Calculo E Costos`
  (`HOJA_CALCULO_ECOSTOS`).
- `ejecutar()`: encadena las funciones anteriores después de escribir
  `Consolidado_entradas.xlsx`, y ahora también escribe `Pagos_BESS.xlsx`.

Explícitamente fuera de esta etapa (decisión del usuario): el resto de columnas de
`Actualizar_Calculos_Columnas` (L, M, N, O, R, S, T, U, W, X, Y, AB:AF, AG:AX, AZ) y toda la
hoja `Calculo RE545`.

**Verificación:** dos scripts de prueba sintéticos (sin persistir en el repo, borrados al
cerrar la sesión, según la convención). El primero prueba cada función nueva contra
`DataFrame`s armados a mano (incluye el caso "clave CMg repetida gana la primera fila" y el
caso "Ventana_No_Completa≠1 → energía en 0 en esta hoja"). El segundo repite las pruebas clave
pasando por un `.xlsx` real (vía `leer_centrales()`/`leer_cmg()`) para confirmar que los tipos
que devuelve `pandas`/`openpyxl` al leer un archivo real (no un `DataFrame` construido a mano)
no rompen `_normaliza_cuarto()` ni la homologación por nombre. Todos los casos coincidieron con
lo esperado a mano. No se probó contra un caso real ni contra la planilla 11.

---

## 2026-09-11 (2) — Fix: `Resumen BESS` con título arriba de los encabezados

Al correr contra un `Centrales.xlsx` real (primera vez que el usuario probó la etapa `Calculo
E Costos` fuera de un caso sintético) salió:

```
La hoja 'Resumen BESS' de Centrales.xlsx debe tener una columna de nombre de central
('Nombre activo') y una de barra de inyeccion ('Barra inyección'). Columnas encontradas:
['Cuadro N° 1: Resumen BESS', 'Unnamed: 1', ..., 'Unnamed: 8']
```

Causa: `leer_centrales()` leía `Resumen BESS` con `pd.read_excel(header=0)`, asumiendo que la
fila 1 ya traía los encabezados. El archivo real trae un título fusionado
("`Cuadro N° 1: Resumen BESS`") en esa fila, y los encabezados reales van una fila más abajo —
mismo problema que ya se había resuelto para el SoC (`detectar_fila_nombres()`), pero acá
todavía no se había aplicado el mismo criterio.

**Fix:** nueva función `_leer_resumen_bess(ruta, nombre_hoja)` que lee la hoja cruda
(`header=None`) y busca, en las primeras 15 filas, la que contiene textos que matchean
`'nombre'+'activ'` y `'barra'` (normalizados) — nunca una posición fija. `leer_centrales()` la
usa para `Resumen BESS`; `Diccionario` no se toca porque ya se leía con `header=None` sin
asumir fila fija. `construir_mapa_barra()` no necesitó cambios: ya buscaba por nombre de
columna, no por posición.

**Verificación:** test sintético con una hoja de 4 filas (título fusionado, encabezados,
2 centrales) reproduciendo exactamente la estructura del error real; `leer_centrales()` +
`construir_mapa_barra()` devolvieron el mapa esperado. Sin persistir en el repo (convención de
pruebas).

---

## 2026-09-11 (3) — Ventana nueva: diagrama de carpetas + "Generar" por salida

El usuario adjuntó `Revisor_Reliquidacion.py` (otro proyecto suyo) como referencia de cómo
quiere que se vea la ventana: un diagrama de texto de la estructura de carpetas/archivos
(prefijos `├──`/`└──`/`│`, monoespaciada) con un botón por fila cuando corresponde, en vez del
checklist plano que había hasta ahora. Pedido concreto: al lado de `Consolidado_entradas.xlsx`
un botón "Generar" que abra una ventana con una casilla por entrada (Medidas, CMg, Ofertas,
etc.) para elegir qué recalcular, y lo mismo para `Pagos_BESS.xlsx` ("ajustamos detalles
después").

Antes de tocar código se preguntó al usuario (`AskUserQuestion`) el punto más consecuente: qué
pasa con una entrada destildada al apretar Actualizar. Eligió explícitamente **"se preserva lo
que ya había"** (frente a "se regenera todo igual" o "queda vacío"), y confirmó el layout
general (carpeta+AAMM igual que hoy, árbol debajo — reemplazando al checklist plano, ya que el
árbol muestra el mismo estado OK/FALTA/PENDIENTE).

**`nucleo.py` — regeneración parcial (cambio de arquitectura, no solo de UI):**

- `_copiar_hoja_existente(wb_origen, nombre_hoja, wb_destino)`: copia una hoja completa (solo
  valores, sin fórmulas ni formato) de un workbook `openpyxl` a otro. Es el mecanismo real de
  "preservar": nunca se intenta reconstruir `Ofertas SSCC` o `FD` (bloques de distinto largo,
  con títulos y `startcol`) a partir de un DataFrame leído de vuelta — se copia la hoja física
  tal cual, evitando reinventar su estructura.
- `escribir_salida()` gana `ruta_existente` y `hojas_regenerar` (`None` = comportamiento
  clásico, regenera las 5 hojas). Cuando `hojas_regenerar` es un set, las hojas fuera de ese set
  se preservan vía `_copiar_hoja_existente()`; si no existían antes, quedan vacías y se registra
  un aviso (log de la corrida + fila del `Log`) en vez de fallar en silencio.
- `SECCIONES_CONSOLIDADO`: agrupa las 4 casillas de la ventana con las hojas que produce cada
  una. Decisión de diseño: **no hay una casilla por archivo de entrada**, sino una casilla
  `"medidores"` que junta Medidas_SAE + SoC + Centrales(Diccionario) + OfertasSSCC, porque
  `construir_medidores()` los necesita siempre los 4 juntos — tildar solo "Ofertas" y dejar
  "Medidas" destildada no permitiría recalcular nada coherente. `"cmg"`, `"fd"`, `"subastas"`
  quedan independientes porque cada uno sale de una sola función/archivo.
- `generar_consolidado(carpeta_base, aamm, secciones_activas, registrar, progreso)`: reemplaza
  a la vieja `ejecutar()`. Valida (y exige) los archivos de entrada **solo para las secciones
  tildadas** — si `"medidores"` no está tildada, no hace falta tener Medidas_SAE/SoC/Centrales/
  Ofertas presentes ni siquiera un AAMM válido.
- `generar_pagos_bess(carpeta_base, registrar, progreso)`: separado de `generar_consolidado()`
  (ya no hace todo un `ejecutar()` monolítico). Lee `Medidores` de `Consolidado_entradas.xlsx`
  ya generado (`pd.read_excel`, no se recalcula) en vez de recibir `df_medidores` en memoria
  como antes — refleja que ahora son dos flujos independientes disparados por botones distintos.
  Perdió el parámetro `aamm` (no lo usaba: todo sale de `Medidores`, que ya trae Mes/Dia/Hora).
- La vieja `ejecutar(carpeta_base, aamm, ...)` se **eliminó** (no solo se dejó como wrapper):
  nada la llama ya que la ventana dispara `generar_consolidado`/`generar_pagos_bess` por
  separado, y mantenerla como código muerto no aportaba nada.

**`Balance_BESS.py` — reescritura de la ventana:**

- `_profundidad_fila()`, `_es_ultimo_en_su_nivel()`, `_prefijos_arbol()`: arman los prefijos
  tipo árbol a partir de la lista plana `(etiqueta, estado, detalle)` que ya devolvía
  `revisar_estructura()`, sin pedirle a `nucleo.py` que sepa de árboles/interfaz. La profundidad
  se deduce del TEXTO de la etiqueta (`"algo/"` = carpeta de primer nivel, `"  hoja ..."` =
  nieto, el resto = archivo hijo de la carpeta anterior) en vez de que `nucleo.py` devuelva un
  campo de profundidad — mantiene a `nucleo.py` sin conceptos de UI.
- El panel "Entradas detectadas" (`pintar_checklist`, filas planas) se reemplaza por "Estructura
  del caso" (`pintar_arbol`, con los prefijos de arriba). Al final del árbol se agregan a mano
  las dos filas de salida (`Consolidado_entradas.xlsx`, `Pagos_BESS.xlsx`), cada una con su
  botón "Generar...".
- Se eliminó el botón único "Ejecutar": cada salida se dispara desde su propia ventana
  (`abrir_ventana_generar_consolidado`, `abrir_ventana_generar_pagos`), con casillas (la primera)
  o solo una explicación (la segunda, sin casillas todavía). `lanzar_generacion()` es el helper
  compartido: corre la función de `nucleo` en un hilo, y hace que `registrar`/`progreso` escriban
  en el log/barra de la ventana PRINCIPAL (no hay log propio por ventana secundaria), para no
  duplicar esos widgets.
- Al terminar una generación se vuelve a llamar `revisar()` para refrescar el árbol (por si el
  archivo de salida cambió de OK/FALTA), y se cierra la ventana "Generar" correspondiente.

**Verificación:** sin entorno gráfico disponible en esta sesión (no hay `tkinter` instalado acá,
solo se pudo correr `python -m py_compile` sobre `Balance_BESS.py`/`nucleo.py`). Se probó por
separado, con scripts sintéticos borrados al cerrar la sesión:
- La lógica de árbol (`_profundidad_fila`/`_prefijos_arbol`, copiadas fuera de `Balance_BESS.py`
  para poder importarlas sin `tkinter`) contra la lista real que devuelve
  `revisar_estructura()` en una carpeta de prueba: el árbol impreso en consola tiene la forma
  esperada.
- `escribir_salida()` con `hojas_regenerar`: preserva hojas no tildadas desde `ruta_existente`,
  regenera las tildadas, y avisa (sin fallar) cuando una hoja a preservar no existía todavía.
- `generar_consolidado()`: no exige Medidas/Ofertas/Centrales si `"medidores"` no está tildada;
  corre de punta a punta con solo `"cmg"` tildada; rechaza secciones desconocidas y el caso de no
  tildar nada.
- `generar_pagos_bess()`: exige `Consolidado_entradas.xlsx` con `Medidores` antes de correr, y
  de punta a punta con datos reales (vía `Centrales.xlsx`/`cmg.xlsx` reales) da los mismos
  resultados que la implementación anterior.

**No probado:** la ventana real (no hay `tkinter` en este entorno) — falta que el usuario la
abra y confirme que el diagrama se ve como esperaba y que los botones "Generar" funcionan en la
práctica.

---

## 2026-09-11 (4) — `Calculo E Costos`, etapa 2: L, N, O, R, S, T, U, W, X, Y, AB, AC, AD

El usuario pidió terminar el cálculo de Ecostos ("necesito que termines con el calculo de
Ecostos"). Antes de implementar se leyó completa la macro `Actualizar_Calculos_Columnas`
(módulo `J_Calculo_Ecostos`, ~1500 líneas) y sus funciones auxiliares (`CrearDiccionarioSubastas`,
`CrearDiccionarioProrrata`, las funciones de ordenamiento por bloques, `CalcularAsignacionEnergia`,
etc.), y se encontró un problema real: buena parte de las columnas restantes (`M`, `AE`, `AF` y
todo `AG:AZ`) depende de una hoja `Resumen` del libro original — **distinta** de
`Centrales.xlsx!Resumen BESS` — con una tabla central→factor y un umbral único (`H8`), y de un
umbral de subida/bajada por central+ventana. Ninguna de las dos cosas está mapeada en la
migración. Se le preguntó al usuario 3 veces con distinto nivel de detalle técnico (la primera
con letras de columna VBA, que no se entendió — "no t entendí" / "no entiendo dime el valoer");
la pregunta se simplificó a lenguaje llano y tampoco se resolvió del todo, pero el usuario dio
una pista concreta que sí resolvió una pieza clave (ver abajo). **Las otras dos (hoja `Resumen`
y umbral de subida/bajada) siguen abiertas** — ver "Pendientes abiertos".

**El problema de `L` y cómo se resolvió:** la fórmula real de `L` compara
central+mes+día+hora de `Calculo E Costos` contra `Subastas`, filtrando por un "tipo"
BAJADA/SUBIDA. El código VBA documentaba esa clave por posición (`Subastas!D` = tipo,
`G,H,I,K` = el resto de la clave), pero esas letras no coinciden con los encabezados reales de
`Subastas` ya confirmados (`D` es `Fecha`, no un texto BAJADA/SUBIDA) — se le pidió al usuario
que revisara la macro `Traspasar_Medidores_A_Calculos_Rapido` por si la respuesta estaba ahí (no
estaba: se confirmó con `grep` que esa macro no menciona nada de Subastas/BAJADA/SUBIDA/Resumen).
El usuario confirmó directamente: **"Es la columna C de la hoja subastas que ya generamos"** —
`Subastas!Sub_Baj`, tal como sugería el propio nombre de la columna. Con eso se reconstruyó la
clave equivalente por NOMBRE de columna real (no por posición): tipo=`Sub_Baj`,
central=`Configuración` (esta última **inferida**, no confirmada letra por letra: es el mismo
campo que usa la tabla dinámica Prorrata SSCC como identificador de central, y da una
alineación semántica limpia con Mes/Dia/Hora_dia; se descartó `Propietario` porque es el campo
que se usa para filtrar por BESS/SAE, no para identificar una central puntual). Queda como
pendiente validar esto contra un caso real (ver "Pendientes abiertos").

**Implementado en `nucleo.py`** (todo lo que NO depende de la hoja `Resumen`):

- `_normaliza_valor_vba(valor)`: replica `NormalizarValor` (mayúsculas+recorte; números sin
  decimales de más, igual que `CStr` en VBA) para armar claves compuestas comparables entre
  `Calculo E Costos` y `Subastas`.
- `calcular_l(df_ecostos, df_subastas)`: ver arriba.
- `calcular_n_o(df_ecostos)`: por grupo (central+ventana), suma acumulada de energía positiva/
  negativa (solo filas `L=1`) ordenando por `Cuarto de Hora` descendente, repartida a todas las
  filas que comparten ese `Cuarto de Hora`.
- `calcular_r_ecostos(df_ecostos)`: ranking por grupo, `CMg` descendente + `Cuarto de Hora`
  descendente, con empates compartiendo el ranking de inicio del bloque ("competition
  ranking"). **Nombre con sufijo `_ecostos` a propósito** — ver el error de abajo.
- `calcular_s_t_u`, `calcular_w_x`, `calcular_y_ab_ac_ad`: ver plan §25.7 para el detalle de
  cada una (`W` tiene una excepción fiel al original: la primera fila de todo el archivo no
  reinicia a 1, toma el valor de `Hora`).
- `completar_calculo_e_costos_grupos(df_ecostos, df_subastas, registrar=print)`: combina todas
  las anteriores. `generar_pagos_bess()` ahora también lee la hoja `Subastas` de
  `Consolidado_entradas.xlsx` (además de `Medidores`) y la llama después de
  `construir_calculo_e_costos()`.

**Error encontrado y corregido durante esta sesión (antes de comitear):** la primera versión de
`calcular_r_ecostos` se llamaba simplemente `calcular_r()`, pisando en silencio a la función que
YA existía con ese nombre para Medidores (`Oferta_Completa_Dia`, lógica no relacionada). Python
no avisa de la redefinición; recién se manifestó como un `TypeError` de argumentos al correr el
test de punta a punta (`construir_medidores()` llama a `calcular_r()` en tiempo de ejecución, y
para ese momento el nombre ya apuntaba a la versión nueva de 1 argumento). Se corrigió
renombrando a `calcular_r_ecostos` y se agregó una trampa en `METODOLOGIA.md` §7: antes de
agregar una función nueva, `grep` para confirmar que el nombre no existe ya.

También se corrigió, en el mismo pase de pruebas, un bug real en `calcular_w_x`: la primera
versión calculaba `W` como "posición dentro del bloque" en vez de "contador que suma 1 desde el
valor de la fila anterior", así que la excepción de la primera fila (`W` = `Hora` en vez de 1)
no se arrastraba al resto de su bloque. Se corrigió sumando el offset (`Hora - 1`) a todas las
filas del primer bloque.

**Verificación:** un test sintético con 2 grupos (central+ventana) y valores elegidos a mano
para poder calcular cada columna manualmente de antemano (incluye empates en `CMg`+`Cuarto de
Hora` para `R`, una fila con `Energia_Positiva=0` para probar el "no calificaI" de `Y`/`AB`, y
un cambio de `Copia_Ventana` a mitad de archivo para probar que `W` es realmente global y no por
grupo) — todas las columnas coincidieron con el cálculo a mano. Un segundo test corrió
`generar_pagos_bess()` de punta a punta con `Centrales.xlsx`/`cmg.xlsx`/`Consolidado_entradas.xlsx`
(con hojas `Medidores` y `Subastas`) reales, confirmando que las 13 columnas nuevas aparecen en
`Pagos_BESS.xlsx` con los tipos y valores esperados. Sin persistir en el repo (convención de
pruebas). No se probó contra un caso real ni contra la planilla 11 — sigue pendiente, y ahora es
más urgente por la inferencia de `L` sin confirmar.

---

## 2026-09-11 (5) — `Calculo E Costos`: M, AE, AF + nombres reales de toda la hoja

El usuario adjuntó `Centrales.xlsx` (real) y un `Libro1.xlsx` con 4 hojas: `FD`, `subastas`,
`E COSTOS` y `Resumen` — respondiendo al bloqueo de la sesión anterior ("Que necesitas de la
hoja original de resumen porque eso es todo lo que hay... puede que los indices esten
diferente").

**Hallazgo clave**: la hoja `Resumen` del libro original y `Centrales.xlsx!Resumen BESS` son
**la misma tabla** — mismos 9 encabezados en el mismo orden (`Nombre activo`, `Pmax (MW)`,
`Horas para descarga forzada`, `Capacidad (MWh)`, `Energía mínima`, `Barra inyección`, `%
Energía sobre mínima (indicador nuevo ciclo)`, `Ciclos max diarios`, `Eficiencia`). No hacía
falta una hoja nueva: el "factor" de `AE`/`AF` es la columna `Pmax (MW)`, y el "umbral" de `M`
es el valor de `% Energía sobre mínima...` en la primera fila de datos (que en el archivo real
cae justo en `H8`, de ahí la referencia fija del VBA original).

**Implementado en `nucleo.py`:**

- `construir_dic_resumen_factor(resumen_bess)`: arma central→`Pmax (MW)` y el umbral de SoC
  mínimo, a partir de la MISMA hoja que ya usa `construir_mapa_barra()` (misma búsqueda por
  nombre de columna normalizado, no por posición).
- `calcular_m(df_ecostos, umbral_soc_minimo)`: `1` si `SoC > umbral`, si no `0`.
- `_calcular_asignacion_energia(bloque, energia_maxima, factor)` + `calcular_ae_af(df_ecostos, dic_factor)`:
  replica `CalcularAsignacionEnergia` — reparte el máximo de `N`/`O` del grupo en bloques de 15
  minutos según el orden `W` y el factor de la central; sin factor o factor=0, queda en blanco
  (`pd.NA`) en vez de fabricar un error de Excel. `Int()` de VBA se replica con `math.floor`
  (redondea hacia abajo incluso en negativos, distinto de truncar hacia cero).
- `completar_calculo_e_costos_grupos()` ahora también agrega `M`, `AE`, `AF`, y al final
  renombra TODAS las columnas con `NOMBRES_CALCULO_E_COSTOS` (nombres reales, de la hoja
  `E COSTOS` de `Libro1.xlsx` — fila 3 tiene el encabezado real de cada columna). Mismo patrón
  que `NOMBRES_FD_CSF`/`NOMBRES_SUBASTAS`: todo el cálculo interno sigue usando los nombres/
  letras de siempre, el rename es el último paso antes de escribir.
- `generar_pagos_bess()` ahora también arma `dic_factor`/`umbral_soc_minimo` (de la misma
  lectura de `Centrales.xlsx` que ya hacía para `mapa_barra`) y se los pasa a
  `completar_calculo_e_costos_grupos()`.

**Confirmación indirecta de la columna `L`**: el archivo de encabezados reales muestra que
`Calculo E Costos!G` se llama literalmente `Configuracion` — el mismo nombre de campo que
`Subastas!Configuración`, que es lo que se venía usando (inferido) para homologar centrales en
`calcular_l()`. Sube la confianza en esa elección, aunque sigue sin confirmarse fila por fila.

**Lo que sigue pendiente (`AG:AZ`)**, según la misma hoja `E COSTOS`:
- `AG:AL` ("Prorratas") necesitan la tabla dinámica Prorrata SSCC, todavía no construida.
- `AM:AR` ("FD") necesitan una categoría `CTF` que no existe en nuestra hoja `FD` (solo tiene
  CSF/CPF) — origen sin identificar.
- `AW` (Descuento FD) necesita el umbral de subida/bajada de Subastas: el archivo de
  encabezados reales muestra `Subastas!R:V` (`Configuración`, `Ciclo`, `Clave`, `SUBIDA`,
  `BAJADA`) como una tabla de resumen aparte, pero sin filas de datos de ejemplo, y no coincide
  con `Subastas!U:W` que usaba el código VBA — la posición real sigue sin confirmarse.

**Verificación:** tests sintéticos (sin persistir en el repo) para `construir_dic_resumen_factor`
(incluye el caso "primera fila con nombre válido" cuando hay filas sin central), `calcular_m`,
`_calcular_asignacion_energia` (bloque completo y bloque parcial con fracción), `calcular_ae_af`
(incluye el caso "central sin factor -> NA"), y que `completar_calculo_e_costos_grupos()`
devuelve exactamente las columnas de `NOMBRES_CALCULO_E_COSTOS` en ese orden. Se repitió también
el test de regresión completo de la etapa 2 (2 grupos, empates en ranking, etc. de la sesión
anterior) contra los nombres reales, sin cambios en los valores esperados. Un test de punta a
punta corrió `generar_pagos_bess()` completo con archivos reales de `Centrales.xlsx`/`cmg.xlsx`/
`Consolidado_entradas.xlsx`, confirmando que las 28 columnas de `Pagos_BESS.xlsx` salen con los
nombres reales y los tipos esperados. No se probó contra un caso real ni contra la planilla 11.

---

## 2026-09-11 (6) — `Calculo E Costos`, etapa 3: `AG:AV` (Prorratas, FD homologado, costo ponderado)

El usuario aportó dos confirmaciones cortas que destrabaron esta etapa: **"creo que no tiene
ctf no está en los FD y en la hoja de los ecostos sale con 0"** y **"Las subastas que te mande
vs las del original están corridas una columna, la primera en el original está vacía"**.

La segunda explica retroactivamente la discrepancia de letras encontrada la sesión anterior
para la columna `L`: aplicando ese corrimiento de una columna a lo que documentaba el VBA
(`Subastas!D`=tipo, `G,H,I,K`=clave), se obtiene exactamente `Sub_Baj` + `Configuración+Mes+
Dia+Hora_dia` — lo mismo que ya se había implementado por inferencia, ahora con una explicación
clara. La primera confirma que `CTF` (columnas `AI`, `AL`, `AO`, `AR`) no necesita ningún
origen: en el VBA original están hardcodeadas en 0 (`salidaAGAX(i,3)=0`, etc.), nunca dependen
de un diccionario — coincide exactamente con lo que el usuario reportó ver en el archivo real.

**Implementado en `nucleo.py`** (todo lo que no depende del umbral de subida/bajada, que sigue
sin resolverse — ver "Pendientes abiertos"):

- `construir_prorrata_sscc(df_subastas)`: arma la tabla dinámica Prorrata SSCC con
  `pandas.pivot_table` directamente desde `Subastas` (`index=[Configuración, Hora_mes],
  columns=Control, values=Sub_Baj, aggfunc=count`) — confirmado hace varias sesiones que NO es
  un archivo externo, pero recién ahora se construye en Python.
- `construir_dic_prorrata()`: busca entre las columnas que deja el pivot la que contenga "cpf"
  y la que contenga "csf" en el nombre (inferido, no confirmado que `Control` tenga esos dos
  valores exactos — avisa si no las encuentra, no falla).
- `calcular_prorratas()`: homologa por central+`Hora Mes` → `(AG, AH)`. `AJ=AG`, `AK=AH`
  (duplicados a propósito, así lo hace el VBA original: `salidaAGAX(i,4)=valorAG`). `AI=AL=0`.
- `construir_dic_mapeo_diccionario()`: TERCERA lectura de la hoja `Diccionario` (columna A→B,
  primera coincidencia gana) — distinta de `construir_homologacion()` y de
  `_mapas_homologacion_fge()`, documentado como trampa nueva en `METODOLOGIA.md` §7.
- `_calcular_bloque()` + `construir_dic_fd_bloque()` + `calcular_fd_prorrateado()`: homologan la
  central contra `Diccionario`, arman una clave "bloque de 4 + central homologada" (uno para
  descarga usando `Y`, otro para carga usando `AC`) y buscan esa clave en `FD!CSF(±)`/`CPF(±)`
  → `(AM, AN, AP, AQ)`. Central no encontrada en `Diccionario` → blanco (`pd.NA`); central
  encontrada pero sin match en `FD` → 0 (fiel al original, que solo registra un aviso). `AO=AR=0`.
- `_calcular_costo_ponderado()` + `calcular_as_at()`: replica `CalcularCostoPonderado` → `(AS,
  AT)`, combinando las Prorratas, el FD homologado y `AE`/`AF`.
- `calcular_au_av()`: promedio de `AB`/`AD` por grupo (central+ventana), activado solo si la
  suma GLOBAL de energía por ventana (TODAS las centrales que comparten esa `Copia_Ventana`, sin
  agrupar por central — una agrupación distinta de la de `N/O/R/Y/AB/AC/AD`) supera ±10 → `(AU,
  AV)`.
- `completar_calculo_e_costos_grupos()` ahora recibe también `diccionario`, `df_fd_csf` y
  `df_fd_cpf`, y agrega las 16 columnas nuevas antes del rename final.
- `generar_pagos_bess()` ahora también busca y lee el archivo `SSCC_Desempeño_*` (con
  `buscar_archivo_sscc_desempeno()` + `construir_fd()`, igual que ya hacía `generar_consolidado()`
  para la sección `"fd"`) y mantiene `diccionario` de `leer_centrales()` en vez de descartarlo.

**`NOMBRES_CALCULO_E_COSTOS` gana valores DUPLICADOS a propósito**: `AG:AL` ("Prorratas") y
`AM:AR` ("FD") comparten los mismos 6 nombres cortos (`CPF(-)`, `CSF(-)`, `CTF(-)`, `CPF(+)`,
`CSF(+)`, `CTF(+)`) porque así están en el archivo real (se distinguen por un encabezado de
grupo en las filas 1-2 que no se replica en nuestro esquema de una sola fila de encabezado) —
mismo criterio que el `"Hora Mes"` duplicado de `FD`. Documentado como trampa en
`METODOLOGIA.md` §7 (indexar por ese nombre después del rename da una `Series` ambigua).

**Verificación:** tests sintéticos (sin persistir en el repo) para cada función nueva por
separado con valores calculados a mano (incluye el caso "central sin match en Diccionario ->
blanco", "central con match pero sin FD -> 0", `CalcularCostoPonderado` con energía/precio en
blanco, y `AU`/`AV` con una tercera central en el mismo `Copia_Ventana` para probar que la suma
global cruza centrales). Un test llamó a `completar_calculo_e_costos_grupos()` completo y
confirmó que el orden final de columnas coincide exactamente con `NOMBRES_CALCULO_E_COSTOS`. Un
test de punta a punta corrió `generar_pagos_bess()` con un archivo `SSCC_Desempeño_*` sintético
pero con la estructura real que espera `construir_fd()` (datos desde la fila 12, filtro BESS/SAE
en columna D) — corrió sin errores; la única discrepancia fue que `pd.read_excel` renombra
columnas duplicadas al releer (`"CPF(-)"` → `"CPF(-).1"`), un comportamiento conocido de pandas
al leer, no un problema de lo que se escribió (confirmado escribiendo y releyendo un `DataFrame`
con columnas duplicadas de prueba). No se probó contra un caso real ni contra la planilla 11.

---

## 2026-09-11 (7) — `Calculo E Costos`, etapa 4: `AW`, `AX`, `AZ` (+ `Subastas!N`) — hoja completa

El usuario pidió terminar lo que quedaba de Ecostos ("Termina lo pendiente de E Costos porfa. Lo
que sigue pendiente (AW, AX, AZ)"). Al empezar la sesión el repositorio **no tenía** el documento
de trazabilidad del `.xlsm`: las sesiones anteriores lo habían leído como adjunto, pero lo único
que quedó escrito fue *qué* faltaba, no el código VBA ni las fórmulas de las tres columnas. Sin
eso había que adivinar la lógica, que es justo lo que prohíben `REGLAS.md` y el plan §18, así que
se pidió el documento antes de tocar código. El usuario lo entregó y pidió dejarlo en el repo:
ahora vive en **`docs/Trazabilidad_11_PAGOS_BESS_2607_Definitivo.md`** (331 KB, incluye el VBA
completo y todas las fórmulas del libro), y `README.md` lo agrega a la tabla de navegación. Es la
causa raíz de este bloqueo: no vuelve a pasar.

**El "bloqueante" resultó no existir.** Lo que estaba anotado como una dependencia circular
(`AW` → umbral → `COUNTIFS` sobre `Subastas!N` → `Calculo E Costos`) se deshizo al leer la
fórmula real de `Subastas!N` (sección 5.3 del documento):

```
=IFERROR(XLOOKUP(1,('Calculo E Costos'!$D$2:$D$50000=J3)*('Calculo E Costos'!$G$2:$G$50000=K3),
        'Calculo E Costos'!$P$2:$P$50000,""),"")
```

`Calculo E Costos!P` es `Ciclo de Carga del mes` (= `Copia_Ventana`), que viene de `Medidores` y
está disponible desde la etapa base: **no depende de ninguna columna calculada**. Y de paso: la
columna `Energía SSCC` no es una energía, es un número de ciclo. El nombre engañaba.

**Implementado en `nucleo.py`** (sección nueva "etapa 4", después de `calcular_au_av`):

- `calcular_subastas_energia_sscc(df_subastas, df_ecostos)`: replica el XLOOKUP de arriba
  homologando por NOMBRE de columna (`Hora_mes` + `Configuración` de nuestra hoja `Subastas`
  contra `Hora Mes` + `clave` de Ecostos), con el mismo corrimiento de una columna ya confirmado
  para `L`. Primera coincidencia gana (XLOOKUP sin modo de búsqueda); sin coincidencia, `""`.
- `construir_dic_umbrales_subastas(df_subastas, energia_sscc)`: replica la tabla `Subastas!U:W`
  del libro original (`U = S&"&"&T`, `V`/`W` = `COUNTIFS` por SUBIDA/BAJADA). **No era un archivo
  externo ni una hoja aparte**: se deriva de `Subastas` + `Subastas!N` contando filas por
  central+ciclo+tipo, exactamente como la Prorrata SSCC. Diferencia deliberada con el original:
  las combinaciones salen de los datos en vez de una lista fija escrita a mano — es equivalente,
  porque una combinación sin datos daría umbral 0 y con umbral 0 ninguna fila pasa el filtro
  `W <= umbral*4` (`W` arranca en 1), o sea `AW = 0` de las dos formas.
- `_clave_central_ciclo()`: la clave `central&ciclo`, que en el original se arma distinto de cada
  lado (concatenación de Excel en `Subastas!U`, `TextoSeguro(G)&"&"&TextoSeguro(P)` en el VBA) y
  tiene que dar lo mismo. Usa `_valor_clave()` para que un `3` y un `3.0` den los dos `"3"`.
- `calcular_aw_ax(df_ecostos, dic_umbrales)`: replica `AW` y `AX` del bloque "AU, AV, AW, AX Y AZ".
  **Ojo con un detalle que parece un error de tipeo y no lo es:** los promedios van cruzados —
  `AB` se promedia con el umbral de **BAJADA** y `AD` con el de **SUBIDA**. Así está en el VBA
  original, se replicó tal cual.
- `calcular_az(df_ecostos)`: `MAX(0, (SUMA(AX) - SUMA(U)) / cantidad de filas)` por grupo
  central+ventana, mismo valor en todas las filas del grupo.
- `NOMBRES_CALCULO_E_COSTOS` suma `AW` → `Descuento FD`, `AX` → `Total`, `AZ` → `Monto a
  compensar`. **`Total` queda duplicado a propósito** (es también el nombre real de `U`), igual
  que los `CPF(-)`/`CSF(-)`/... repetidos entre `AG:AL` y `AM:AR`: en el archivo real se
  distinguen por el encabezado de grupo de las filas 1-2, que este esquema de una sola fila de
  encabezado no replica. Para llegar sin ambigüedad a una de esas columnas hay que ir por
  posición, no por nombre (los tests de esta sesión lo hacen así).

`AY` **no** se calcula: la macro salta de `AX` a `AZ` y el documento no muestra ninguna región de
fórmulas para `AY4:AY26787` (sección 5.4). Se documentó explícitamente para que no parezca un
olvido.

**Verificación** (tests sintéticos, sin persistir en el repo, según la convención):

- `Subastas!N`: cruce por Hora_mes+central, incluido el caso "no hay fila que cruce → `''`" y el
  caso "gana la primera coincidencia".
- Umbrales: los conteos SUBIDA/BAJADA salieron exactamente los esperados a mano, incluido un tipo
  distinto de SUBIDA/BAJADA (no cuenta), una fila sin ciclo (no entra en ninguna clave) y la
  normalización `5.0` → `"5"` en la clave.
- `AW`: un grupo armado a propósito con `W = [1, 4, 5, 9]` y umbrales (subida 2, bajada 1), donde
  los filtros `W<=4` y `W<=8` dejan subconjuntos distintos — promedios 15 y 4 calculados a mano,
  y las 4 filas coincidieron. Casos borde: grupo sin umbrales → 0; umbral 0/0 → 0; `AE` en blanco
  → 0 **solo en esa fila**; `cantidadADW = 0` → 0 en **todo** el grupo.
- `AX = AU + AV - AW` fila a fila, y `AZ` constante por grupo, nunca negativo (probado con una
  suma de `U` enorme) y dividido por la cantidad de filas del grupo.
- Corrida completa de `completar_calculo_e_costos_grupos()` con datos sintéticos: devuelve las 48
  columnas de `NOMBRES_CALCULO_E_COSTOS` en orden, las tres nuevas numéricas en todas las filas.

No se probó contra un caso real ni contra la planilla 11 — sigue siendo el pendiente principal.

---

## 2026-09-11 (8) — `Calculo RE545`, etapa base (`A:V`)

El usuario pidió seguir con `Calculo RE545` y entregó `Calculo_RE545_reducido_para_IA.xlsx`
(hoja real recortada: fila 3 = nombres de columna, filas 1-2 = títulos de grupo, hoja
`Mapa_Formulas` con todas las familias de fórmulas, columna `CF` con la fila original). Se copió
al repo como `docs/Calculo_RE545_reducido_para_IA.xlsx` — mismo criterio que el documento de
trazabilidad: si es la fuente de una decisión, vive en el repo, no en un adjunto de sesión.

Avisó además que quedaba poco contexto y que se iba a dormir, así que esta entrada se escribe
con el detalle suficiente para que la próxima sesión continúe sin preguntarle nada.

**Hallazgo estructural:** `Calculo RE545` es casi toda **fórmulas en la hoja**, no valores
escritos por macro (al revés que `Calculo E Costos`, donde la macro J escribe casi todo). Las
únicas columnas que escribe el VBA son las del traspaso y `Q`/`R`.

**Cómo se reparten las dos hojas** (`Traspasar_Medidores_A_Calculos_Rapido` recorre `Medidores`
una sola vez y escribe en las dos): `A:G`, `K` y `P` van **iguales a las dos**; lo que se reparte
es la energía, según `Medidores!T` (`Ventana_No_Completa`): `= 1` → E Costos; cualquier otro
número, vacío, no numérico o error → RE545. Solo RE545 recibe además `T` (= `Medidores!L`,
"Ventana de valorizacion") y `R` (`CMg!I`, "CMg Promedio", porque `CompletarDestinoTurbo` se
llama con `escribirR:=True` para esta hoja y `False` para la otra).

**Cambios en `nucleo.py`:**

- `construir_dic_cmg()` ahora guarda **el par** `(CMg!F, CMg!I)` en vez de solo `CMg!F` — son los
  dos elementos del `Array()` del diccionario VBA. Se agregó `_buscar_cmg(dic_cmg, barra,
  cuarto_hora)` como helper compartido (replica `CompletarDestinoTurbo`), y
  `construir_calculo_e_costos()` ahora toma `[0]` de ese par. **Si algo se rompe en CMg, mirar
  acá primero**: es un cambio en una estructura que ya usaban las dos hojas.
- `construir_calculo_re545()`: etapa base (traspaso + Barra + Q/R + T).
- `construir_dic_resumen_eficiencia()`: central → `Eficiencia` de `Resumen BESS` (la 9na columna
  de las 9 de esa tabla = el `VLOOKUP(...,9,0)` de `V`). Es una **tercera** lectura de esa hoja,
  distinta de `construir_mapa_barra()` y `construir_dic_resumen_factor()` — no fusionar.
- `calcular_s_re545()`: el "ranking cmg" de esta hoja. **No es el mismo** que el de E Costos:
  agrupa por central + `T` (ventana de valorización) y ordena por `CMg Promedio` + `Hora`, no por
  `CMg` + ciclo.
- `calcular_u_v_re545()`: `U` (`EiniT`) = `SoC % × Pmax (MW) × 1000`; `V` (`EalmT`) = la carga
  total del grupo central+ventana cambiada de signo, por la `Eficiencia`. Central que no está en
  `Resumen BESS` → las dos en blanco (equivale al `#N/A` del VLOOKUP).
- `completar_calculo_re545()`: agrega `L`, `M`, `N`, `O`, `S`, `U`, `V` y renombra a
  `NOMBRES_CALCULO_RE545`. `L`, `M`, `N` y `O` son **literalmente las mismas fórmulas** que en E
  Costos, así que se reusan `calcular_l()`, `calcular_m()` y `calcular_n_o()` (verificado contra
  la fórmula real de RE545, que es la versión explícita de lo que esa función ya hacía).
- `escribir_pagos_bess()` acepta un `df_re545` opcional y escribe la hoja `Calculo RE545`
  (`HOJA_CALCULO_RE545`) en `Pagos_BESS.xlsx`; `generar_pagos_bess()` la arma y la pasa.

**Trampa nueva, importante:** `R`, `S`, `T`, `U` y `V` existen en las dos hojas y **significan
cosas distintas** en cada una (`U` es "Total" en E Costos y "EiniT" en RE545). Nunca reusar una
función de una hoja en la otra sin leer antes la fórmula real de la columna.

**Verificación** (tests sintéticos, sin persistir en el repo): el reparto de energía entre las dos
hojas (incluido el caso `Ventana_No_Completa` vacío → RE545) y que la hoja hermana sigue dando lo
complementario; `T` = `Medidores!L`; `Q`/`R` del par de CMg; `S` con un grupo armado a propósito
con empates de `CMg Promedio` resueltos por `Hora` (valores calculados a mano: 1.5, 1.25, 1.0,
1.75, 1.0); `U`/`V` con una central sin ficha en `Resumen BESS` → blanco; y la corrida completa
de `completar_calculo_re545()` devolviendo las 22 columnas con los nombres reales. Se corrieron
también los tests de la etapa 4 de E Costos como regresión (el cambio de `construir_dic_cmg`
podía romperlos): pasan.

**Lo que sigue** (detalle por bloque en el plan §26.3): `AC:AU`, `AW:BG` y `BI:CE`.

---

## 2026-09-11 (9) — `Calculo RE545`, etapa 2: `AC:AU` (reservas por subasta)

Segunda etapa de RE545, en la misma sesión (el usuario pidió avanzar sin consultarlo).

`AC:AU` son tres bloques de 6 columnas con los **mismos 6 encabezados** (`CPF(-)`, `CSF(-)`,
`CTF(-)`, `CPF(+)`, `CSF(+)`, `CTF(+)`), que se distinguen por el título de grupo de la fila 2:
"Subastas" (`AC:AH`), "FD" (`AI:AN`) y "FMA" (`AO:AT`). Los tres son el mismo `SUMIFS` contra
`Subastas` cambiando solo la columna sumada, y `AU` = `SUMPRODUCT` de los tres bloques `/4*1000`.

- `construir_dic_reservas_subastas(df_subastas)`: arma los tres diccionarios
  `(central, hora del mes, tipo) -> suma`. Criterios por NOMBRE (`Configuración`, `Hora_mes`,
  `Control`), columnas sumadas por POSICIÓN (`O`, `P`, `Q`).
- `calcular_reservas_re545(df_re545, dics)`: devuelve las 18 columnas + `AU`.
- `NOMBRES_CALCULO_RE545` suma los 18 nombres repetidos + `SUMA Reservas*FMA*FD`.

**Decisión documentada (pendiente de validar, no bloquea):** los nombres reales de `Subastas`
llaman `FD` a `O` y `FMA` a `P` — corridos una columna respecto de los títulos de grupo de RE545,
que dicen Subastas/FD/FMA para `O`/`P`/`Q`. Es el mismo corrimiento de una columna que el usuario
ya había descrito para el archivo de Subastas. Se siguió **la fórmula** (posición), no el nombre,
porque la fórmula es la fuente primaria. Si al validar con datos reales los tres bloques salen
corridos entre sí, esto es lo primero que hay que revisar (plan §26.3).

**Verificación:** tests sintéticos con dos filas de `Subastas` que comparten central+hora+tipo
(para probar que el `SUMIFS` suma y no pisa), un tipo que no existe en los datos (→ 0, no blanco),
una fila de RE545 cuya hora del mes no cruza con nada (→ los 18 en 0 y `AU` en 0) y `AU`
calculado a mano (`(8x4x1 + 7x4x0.25)/4*1000 = 9.750`). Regresión completa de E Costos (etapas
2-4) y de la etapa base de RE545: pasan.

---

## 2026-09-11 (10) — `Calculo RE545`, etapa 3: `AW:BG` (resumen por central + ventana)

Tercera etapa de RE545 en la misma sesión.

**Hallazgo estructural:** `AW:BG` **no son más columnas del bloque principal**: son una tabla
aparte de 288 filas (9 centrales × 32 ventanas) contra las 26.787 del bloque principal,
compartiendo la hoja. Tercer caso del mismo patrón en este proyecto (ya había pasado con los
bloques CSF/CPF de `FD` y con las dos tablas de `Ofertas SSCC`). Se escribe al lado del bloque
principal con una columna en blanco de separación (`escribir_pagos_bess()` ahora acepta
`df_resumen_re545` y usa `startcol`).

**De dónde sale `AY` ("Oferta Completa"), que no es fórmula ni la escribe ninguna macro:** es la
columna `Completa` de `construir_resumen_ventana_oferta()` — la misma tabla central+ventana que ya
alimenta `Medidores!T` (`T = 1 - Completa`). Coinciden el nombre, la clave, el dominio (0/1) y el
sentido, y el archivo real lo confirma: la central cuya última ventana queda incompleta tiene
`AY = 0` justo ahí. No hubo que inventar nada ni pedir un archivo nuevo: `generar_pagos_bess()`
reconstruye ese resumen desde la hoja `Medidores` ya generada.

**Otro dato que se resolvió de paso:** `Medidores!$S$1` (que usa `BV`) es la **hora de inicio de
ventana**, el mismo dato que la constante `INICIO_VENTANA` — se deduce de la fórmula de
`Medidores!L`, que incrementa la ventana justo cuando la hora es igual a `S1`.

**Cambios en `nucleo.py`:** `NOMBRES_RESUMEN_RE545`, `calcular_bv_re545()`,
`construir_resumen_ventanas_re545()`. Además `completar_calculo_re545()` ya **no** renombra (el
renombre pasó a `renombrar_calculo_re545()`, que se llama al final): la tabla resumen necesita el
DataFrame con los nombres internos.

**Verificación:** tests sintéticos con grupos armados a propósito — `AZ`/`BA` tomando la PRIMERA
fila del grupo y no la última (el `AGGREGATE(15,6,...,1)` del original), `BB` como suma de `AU`,
`BF` sumando solo las filas de la hora anterior al inicio de ventana, `BC` calculado a mano
(`MIN(MAX(MIN(120,50),20),40) = 40`), el caso "oferta incompleta → `BC = 0`" y el caso "ventana 31
→ `BC = 0`". Más un test de escritura real del `.xlsx` confirmando que las dos tablas quedan lado
a lado con una columna en blanco entre medio. Regresión de E Costos y de las etapas 1-2 de RE545:
pasan.

**Pendiente:** `BI:CE` (Componentes 1 y 2) y, con eso, `BD`/`BE` del resumen.

---

## 2026-09-11 (11) — `Calculo RE545`, etapa 4: `BI:CE` — hoja completa

Última etapa de RE545 en la misma sesión. Con esto quedan **completas las dos hojas de cálculo**
del libro (`Calculo E Costos` y `Calculo RE545`).

**Corrección de un error propio de las etapas anteriores de esta misma sesión** (queda anotado
porque cambia valores ya commiteados): `VLOOKUP(G, Resumen!$B$8:$J$26, 4, 0)` **no es `Pmax
(MW)`**, es **`Capacidad (MWh)`**. El orden real de las 9 columnas de `Resumen BESS` es `Nombre
activo`, `Pmax (MW)`, `Horas para descarga forzada`, `Capacidad (MWh)`, `Energía mínima`, `Barra
inyección`, `% Energía sobre mínima`, `Ciclos max diarios`, `Eficiencia` — y es consistente con
que `H` use el índice 6 para `Barra inyección`. Afectaba a `U` (`EiniT`) y a `BC` (`Edisp_T`), que
se habían implementado con `Pmax`. Se agregó `construir_dic_resumen_capacidad()` y se corrigieron
las dos. Tabla de referencia de qué índice usa cada columna, en el plan §26.7. **`AE`/`AF` de
E Costos siguen bien con `Pmax`** (ahí el VBA usa `Resumen!B:C`, o sea el índice 2), igual que `BN`
de RE545.

**Implementado (`calcular_componentes_re545()` + `completar_checks_resumen_re545()`):** `BI`
(`Orden`, con el salto de 4 filas del original), `BJ` (`Periodo`), `BK`/`BL` (sumas por
central+orden+ventana+periodo), `BM` (`LARGE(IF(...))` matricial: el k-ésimo `CMg` más grande
entre todas las filas con el mismo `BK`, con `k = Periodo/15 + 1`), `BN` (`Edisp_Asig`), `BO`,
`BQ`, `BR`, `BS` (`INDEX/MATCH` matricial), `BT`, `BU`, `BV`, `BW`, `BX`, `BY`, `BZ`, `CA`, `CC` y
`CE` (`Monto a compensar`). Y con eso, `BD`/`BE` del resumen `AW:BG`.

**Las dos recursiones, que es lo único que no se puede vectorizar:** `BN` necesita los `BN`
anteriores de su grupo (se recorre de arriba hacia abajo) y `BU` necesita los `BU` **posteriores**
(se recorre de abajo hacia arriba, porque `BT` mira las filas siguientes). `BY` necesita los `BZ`
anteriores, pero `BZ` no depende de `BY`, así que ahí alcanza con calcular `BZ` primero.

**Verificación:** un caso sintético de 8 filas (2 horas × 4 bloques) con todos los valores
elegidos para poder calcular a mano: `BI` = 1,1,1,1,2,2,2,2; `BM` = 80,70,60,50 repetido; `BN`
cortándose al llegar al `Edisp_T` del grupo (150, 150, 100, 0, ...); `BT`/`BU` con la recursión
hacia arriba (400,400,400,400,300,200,100,0 y 0,0,0,0,100,100,100,100); `BY` como acumulado de
`BZ`; y `CE` en sus dos ramas — la de `MAX(...,0)` (da 0) y una positiva calculada a mano
(`(28.500 - 26.000)/8 = 312,5` por fila). Más el caso de inyección negativa (`BZ = 0`, `BY = BX`,
`BU = 0` por `BS = 0`) y los dos checks del resumen. Regresión de E Costos (etapas 2-4) y de las
etapas 1-3 de RE545: pasan.

**Lo que sigue** (nuevo frente, ninguna analizada todavía): las hojas de salida que consumen estas
dos — `PRORRATA_RETIROS`, `Compensacion total`, `Resumen` y el CSV.

---

## 2026-09-11 (12) — Casillas por hoja en la ventana "Generar" de `Pagos_BESS.xlsx`

El usuario pidió cerrar el pendiente que había quedado anotado ("ajustamos detalles después"):
la ventana "Generar" de `Pagos_BESS.xlsx` era todo o nada (una sola hoja, sin casillas), a
diferencia de la de `Consolidado_entradas.xlsx` (`SECCIONES_CONSOLIDADO`). Ahora que la hoja
tiene dos salidas (`Calculo E Costos` y `Calculo RE545`, completas desde la sesión anterior),
pedido explícito: **dos casillas, una por hoja**.

**Cambios en `nucleo.py`** (mismo patrón que `SECCIONES_CONSOLIDADO`/`generar_consolidado`):

- `SECCIONES_PAGOS`: tupla con `("ecostos", "Calculo E Costos", descripción, ("Calculo E
  Costos",))` y `("re545", "Calculo RE545", descripción, ("Calculo RE545",))`.
- `escribir_pagos_bess()` gana `ruta_existente` y `hojas_regenerar` (antes solo tenía
  `df_ecostos`/`df_re545`/`df_resumen_re545`/`registrar`). Con `hojas_regenerar=None` se
  comporta exactamente igual que antes (retrocompatible: los tests de sesiones anteriores que
  la llaman posicionalmente sin estos parámetros nuevos siguen funcionando tal cual). Con un
  `set`, la hoja que NO está en el set se copia tal cual desde `ruta_existente` en vez de
  escribirse desde el DataFrame — reusa `_copiar_hoja_existente()`, la misma función que ya
  usaba `escribir_salida()` para `Consolidado_entradas.xlsx`. Si no hay versión anterior para
  preservar, la hoja queda vacía y se registra un aviso (mismo criterio, sin duplicar código).
- `generar_pagos_bess()` gana el parámetro obligatorio `secciones_activas` (antes no lo tenía;
  es un cambio incompatible a propósito, como ya había pasado con `generar_consolidado()`).
  Valida secciones desconocidas y "ninguna tildada" antes de tocar ningún archivo. Las lecturas
  compartidas (`Medidores`, `Subastas`, `Centrales.xlsx`, `cmg.xlsx`) se hacen siempre que haga
  falta alguna sección; lo que se condiciona es el CÁLCULO de cada hoja:
  - `dic_eficiencia`/`dic_capacidad` (que solo usa RE545: `V`, `U`/`BC`/`BN`) solo se arman si
    `"re545"` está tildada.
  - El archivo `SSCC_Desempeño_*` (que solo usa `Calculo E Costos`, para `AM:AR`) solo se exige
    y se lee si `"ecostos"` está tildada. **Confirmado con un caso real armado a propósito**:
    tildar solo `"re545"` corre sin pedir ese archivo aunque no exista en la carpeta del caso;
    tildar solo `"ecostos"` sí lo exige y falla con un mensaje claro si falta.
  - `df_ecostos`/`df_re545`/`df_resumen_re545` quedan en `None` si su sección no está tildada, y
    así se le pasan a `escribir_pagos_bess()` junto con el `hojas_regenerar` correspondiente.

**Cambios en `Balance_BESS.py`:** `abrir_ventana_generar_pagos()` reescrita para recorrer
`nucleo.SECCIONES_PAGOS` igual que `abrir_ventana_generar_consolidado()` recorre
`SECCIONES_CONSOLIDADO` — un `LabelFrame` con casilla + descripción por sección, valida que haya
al menos una tildada, y pasa `secciones_activas` a `generar_pagos_bess()`.

**Verificación:** tests sintéticos de `escribir_pagos_bess()` (primera corrida escribe las dos
hojas; segunda corrida con solo `"ecostos"` recalcula esa hoja y preserva RE545 tal cual estaba,
sin vaciarlo; tercera corrida con solo `"re545"` al revés; caso sin versión anterior para
preservar → hoja vacía + aviso) y de `generar_pagos_bess()` (secciones vacías/desconocidas →
`ErrorEntrada`). Además un caso **con archivos reales** armado a propósito (`Centrales.xlsx`,
`cmg.xlsx`, `Consolidado_entradas.xlsx` con `Medidores`/`Subastas`) para confirmar en la práctica
que "solo RE545" no pide `SSCC_Desempeño_*` y que "solo E Costos" sí, y que una corrida que falla
a mitad de camino (por archivo faltante) no toca el `Pagos_BESS.xlsx` ya existente. Regresión
completa de las sesiones anteriores (Calculo E Costos etapas 2-4, Calculo RE545 etapas 1-4):
pasa. No se probó la ventana tkinter en sí (sin entorno gráfico en esta sesión, como siempre).

---

## 2026-09-11 (13) — Confirmado: `Configuración` (no `Propietario`) para homologar centrales

El usuario confirmó explícitamente: **"si se usa configuración"**. Cierra la última inferencia
sin confirmar de las que quedaban documentadas — la homologación de central en `calcular_l()`
(y, por extensión, en todo lo que reusa el mismo criterio: `AG:AL`/`AM:AR` de Prorratas/FD
homologado, `AW` con la tabla de umbrales, `calcular_reservas_re545()` de `Calculo RE545`, y
`calcular_subastas_energia_sscc()`) usa `Subastas!Configuración`, **no** `Subastas!Propietario`.
No se tocó código: la implementación ya usaba `Configuración` desde que se resolvió por
inferencia razonada (mismo nombre de campo que `Calculo E Costos!G`, alineación semántica con
Mes/Dia/Hora_dia); esta sesión solo pasa esa elección de "inferida, pendiente de confirmar" a
"confirmada por el usuario" en toda la documentación (`docs/Plan_Traspaso_Python_Balance_BESS.md`
§25.6, `METODOLOGIA.md` §7, `BITACORA.md`).

**Sigue pendiente** (no es lo mismo que la homologación en sí): validar fila por fila contra un
caso real que la cantidad de filas con `L=1` sea razonable — eso confirma que el CRUCE funciona
bien con datos reales, más allá de que ya esté confirmado qué columna usar.

---

## 2026-09-11 (14) — Fix: crash con datos reales cuando `Subastas` no tiene filas BAJADA/SUBIDA

**Primera corrida contra datos reales** (el usuario corrió `generar_pagos_bess()` desde la
ventana, con `SSCC_Desempeño_Julio_2026_V2.xlsx` real: 6.696 filas CSF/CPF leídas bien). Se
cayó en `completar_calculo_e_costos_grupos()` → `calcular_l()` → `_construir_set_subastas_tipo()`
con:

```
ufunc 'add' did not contain a loop with signature matching types
(dtype('int64'), dtype('<U1')) -> None
```

**Causa raíz:** el filtro por `Subastas!Sub_Baj` en `{BAJADA, SUBIDA}` no encontró **ninguna
fila** en el archivo real del usuario — la tabla filtrada (`sub`) quedó con 0 filas. Con 0 filas,
`pandas.Series.map()` es un no-op que no llega a ejecutar la función: devuelve una Series vacía
con el **mismo dtype que tenía antes** de mapear, en vez del dtype que devolvería la función.
Como `Subastas!Mes` se lee de Excel como `int64` (es una columna numérica), el resultado de
`sub["Mes"].map(_normaliza_valor_vba)` quedó en `int64` en vez de texto, y al concatenarlo con
`"¦"` (separador de clave) para armar la clave compuesta, `numpy` no encuentra una operación
`int64 + texto` y lanza el error. **No se reprodujo con los tests sintéticos de sesiones
anteriores** porque ninguno armó a propósito el caso "cero filas después de filtrar" — siempre
había al menos una fila BAJADA o SUBIDA en los datos de prueba.

**Fix en `nucleo.py`:** nueva función `_columna_clave_vba(serie)` = `serie.map(_normaliza_valor_
vba).astype(str)` — el `.astype(str)` fuerza el dtype a texto **siempre**, esté vacía la Series o
no, corrigiendo el caso que `.map()` no cubre. Reemplaza el patrón `columna.map(_normaliza_valor_
vba)` en los 4 lugares donde participa de una concatenación con "+": `_construir_set_subastas_
tipo()` (el que crasheaba), `calcular_l()`, `calcular_prorratas()` y `calcular_subastas_energia_
sscc()`. Los otros dos usos de `.map(_normaliza_valor_vba)` en el archivo (`calcular_reservas_
re545()`) no se tocaron: ahí el resultado se usa como clave de tupla en un diccionario, no se
concatena con "+", así que no está expuesto a este bug.

**Verificación:** se reprodujo el error exacto con un `DataFrame` armado a propósito (columnas
`object`/`int64` "clásicas", no el dtype `str` nuevo de pandas 3.x que usa este sandbox por
defecto — hubo que forzar los dtypes explícitamente para reproducirlo, porque con el dtype nuevo
la concatenación no fallaba) y una fila cuyo `Sub_Baj` no es ni `BAJADA` ni `SUBIDA` (0 filas tras
filtrar). Confirmado que antes del fix reproduce el `ufunc 'add'` exacto y que después corre sin
error, devolviendo `L=0` en todas las filas (comportamiento correcto: si de verdad no hay ninguna
subasta que cruce, nadie participa). Regresión completa de las 14 sesiones anteriores: pasa.

**Hallazgo de fondo, no un bug de código, queda como pendiente:** que el filtro haya dado 0 filas
significa que en el caso real usado, **ninguna fila de `Subastas!Sub_Baj` normaliza exactamente a
`BAJADA` o `SUBIDA`**. Puede ser el período correcto (sin subastas ese mes) o puede ser que el
texto real use otra palabra/formato — hay que confirmarlo con el usuario (ver "Pendientes
abiertos"). Mientras tanto, con ese caso, `L`, `M`, `N`, `O` y las reservas por subasta de
`Calculo RE545` dan 0/vacío en todas las filas — no es un error, es el resultado correcto de la
fórmula real con ese filtro.

---

## 2026-09-11 (15) — Fix: SoC con nombres de central como ruta SCADA + hallazgo en `Subastas`

Segunda corrida contra datos reales del usuario, dos problemas nuevos.

### Fix: SoC — nombres de bloque como ruta SCADA completa

El log mostró: `Centrales en Medidas_SAE.xlsx sin bloque de SoC: [9 centrales limpias]` y
`Centrales en el SOC que no estan en Medidas_SAE.xlsx: [rutas tipo
'\\SRV-SCADA-AF1\SEN\Generación\SEN\03 Región II\SAE-Del Desierto|Nombre PCP/PID']` — cero
cruces. Causa: en el archivo real de SoC (exportación tipo PI), el nombre de cada bloque de
central **no es el nombre limpio**: es la ruta SCADA completa del punto, con un sufijo `|Nombre
PCP/PID` (o `|Nombre` en al menos un caso) pegado al final. `detectar_bloques()` toma ese texto
tal cual (`nombre_bess_origen = str(valor).strip()`, es lo correcto — no reinterpreta nada), pero
`extraer_soc()` buscaba ese texto LITERAL en el `Diccionario`, y ningún humano escribe esa ruta
completa a mano en una hoja de equivalencias — por eso el cruce daba siempre 0.

**Fix:** `_extraer_nombre_desde_ruta_scada(texto)` — si el texto tiene `\`, devuelve solo el
último tramo (después de la última `\`) sin el sufijo después de `|` (ej. de la ruta de arriba
saca `"SAE-Del Desierto"`). No es una reinterpretación de datos: es separar una estructura ya
presente en el archivo (ruta + sufijo), no adivinar a qué central corresponde. `extraer_soc()`
ahora prueba primero el texto literal (compatibilidad con cualquier SoC "limpio" sin ruta) y, si
no hay match, prueba de nuevo con el nombre extraído — el `Diccionario` puede tener cualquiera de
las dos formas. Si ninguna tiene match, usa el nombre **limpio** (no la ruta completa) como
`central`, solo para que avisos/incidencias sean legibles — sigue sin cruzar contra `Medidores`,
mismo comportamiento que antes.

**Importante, se lo dejo dicho al usuario en el chat:** este fix por sí solo **no alcanza** para
la mayoría de las 9 centrales. Comparando el nombre que queda tras extraerlo de la ruta contra el
nombre limpio real de `Medidas_SAE.xlsx`:

| Extraído de la ruta SoC | Real en Medidas_SAE.xlsx | ¿Cruza solo con el fix? |
|---|---|---|
| `SAE-Tocopilla` | `SAE-TOCOPILLA` | Sí (`normalizar()` ya ignora mayúsculas) — pero el merge final es por texto LITERAL, así que de todas formas necesita un `Diccionario` que devuelva el texto exacto `SAE-TOCOPILLA` |
| `SAE-Del Desierto` | `SAE-DEL-DESIERTO` | No (espacio vs guion) |
| `SAE-PE La Cabaña` | `SAE-CRCA-PE-LA-CABANA` | No (falta prefijo `CRCA-`) |
| `SAE-PFV Victor Jara` | `SAE-CRCA-PFV-VICTOR-JARA` | No (falta prefijo) |
| `SAE-PFV Andes Solar 4` | `SAE-CRCA-PFV-ANDES4` | No |
| `SAE-PFV Andes Solar III` | `SAE-CRCA-PFV-ANDES3` | No (`III` vs `3`) |
| `SAE-PFV Nuevo Quillagua II` | `SAE-CRCA-PFV-NUEVO-QUILLAGUA-2` | No (`II` vs `2`) |
| `SAE-PFV Don Humberto` | `SAE-CRCA-PFV-DON-HUMBERTO` | No (falta prefijo) |

Las 9 (todas, en la práctica) necesitan una fila en `Centrales.xlsx!Diccionario` con el nombre
limpio de `Medidas_SAE` y el nombre extraído de la ruta (o la ruta completa, cualquiera de las
dos funciona ahora) — eso es contenido del archivo del usuario, no algo que el código deba
adivinar (`Diccionario` es mantenimiento manual, ver `METODOLOGIA.md`).

**Caso sin resolver, no es una ruta SCADA:** `'07 Region RM'` apareció como un bloque completo
aparte, sin `\` ni `|`. No encaja con el patrón de los otros 8 — no se adivinó qué es (podría ser
un bloque real mal cortado, o una columna/central que no corresponde). Se le preguntó al usuario.

### Hallazgo (sin tocar código): posible columna faltante en `Subastas`

El usuario compartió una foto de la hoja `Subastas` real de la planilla 11. Una fila de datos
muestra, en celdas consecutivas: `CSF(-)` (que es exactamente el resultado de la fórmula real de
`Control`, sección 5.3 del documento de trazabilidad: `=IF(AND(C1="CSF",D1="SUBIDA"),"CSF(+)",
IF(AND(C1="CSF",D1="BAJADA"),"CSF(-)",...))`) seguido de `CSF` y de `BAJADA`. Esa fórmula necesita
DOS insumos (`C`=servicio SIN dirección, `D`=dirección) para armar el label de `Control` — pero
`NOMBRES_SUBASTAS` hoy solo tiene UNA columna entre `Control` y `Fecha` (`Sub_Baj`, que se asumía
que guardaba `SUBIDA`/`BAJADA` directamente, confirmado hace sesiones por el usuario: "Es la
columna C de la hoja subastas que ya generamos"). Si la foto es correcta, falta una columna
"Servicio" que nunca se mapeó, y **todo lo que sigue después de `Control` queda corrido una
posición**. Esto también explicaría el otro hallazgo de la sesión anterior (`Sub_Baj` dando 0
filas BAJADA/SUBIDA con datos reales — si el filtro estaba comparando contra la columna
equivocada, por supuesto no cruza nada).

**No se tocó código todavía.** Este archivo ya tuvo varias rondas de "corrimiento de columna" mal
resueltas por inferencia (ver sesiones anteriores); antes de tocar `NOMBRES_SUBASTAS` de nuevo
hace falta confirmación letra por letra del usuario, no otra inferencia visual sobre una captura
de pantalla. Se le pidió que confirme el contenido exacto de las primeras columnas de `Subastas`
en la planilla 11 real.

**Verificación:** tests sintéticos de `_extraer_nombre_desde_ruta_scada()` (ruta con los dos
sufijos vistos, texto sin ruta queda intacto) y de `extraer_soc()` de punta a punta (con y sin
`Diccionario`, y confirmando que `nombre_scada_original` sigue guardando la ruta completa sin
tocar, solo `central` cambia). Regresión completa: pasa.

---

## 2026-09-11 (16) — Corrección grande: `NOMBRES_SUBASTAS` estaba mal desde el principio

El usuario adjuntó `Libro1.xlsx` con la hoja `subastas` real (encabezados **y fórmulas**, no solo
nombres) y pidió: "revisa la macro `Cargar_Remuneracion_Subastas_Rapido` y arma según el adjunto
en el mismo formato de salida que tengo en la planilla 11". Se copió el archivo al repo como
`docs/Libro1_Subastas_real.xlsx` (mismo criterio que los otros adjuntos que son fuente de una
decisión). Cruzando ese archivo contra el código VBA de la macro (ya en el documento de
trazabilidad) se encontró la causa raíz de la confusión que venía arrastrándose desde hace varias
sesiones (el "corrimiento de columna" que nunca terminaba de cuadrar del todo).

**La prueba definitiva** es la fórmula real de `Subastas!B1`:

```
=IF(AND(C1="CSF",D1="SUBIDA"),"CSF(+)",IF(AND(C1="CSF",D1="BAJADA"),"CSF(-)",
  IF(AND(C1="CTF",D1="SUBIDA"),"CTF(+)",IF(AND(C1="CTF",D1="BAJADA"),"CTF(-)","REVISAR"))))
```

Esta fórmula arma la columna `B` a partir de DOS insumos: `C` (el tipo de servicio SIN dirección:
`CSF`/`CTF`/`CPF`) y `D` (la dirección: `SUBIDA`/`BAJADA`). `NOMBRES_SUBASTAS` **solo tenía una
columna ahí** (`B`="Control", asumiendo que guardaba el label completo como `"CSF(-)"`, y
`C`="Sub_Baj" asumiendo que ahí vivía directamente `SUBIDA`/`BAJADA`) — le faltaba contar una
columna entera: la real `B` = `Concepto` (el label completo, lo que veníamos llamando "Control"),
`C` = `Control` (el tipo SIN dirección, una columna que nunca se había mapeado) y recién `D` =
`Sub_Baj`. Confirmado letra por letra contra los datos reales del archivo (fila 3: `B3='CSF(-)'`,
`C3='CSF'`, `D3='BAJADA'`, `K3='SAE-CRCA-PFV-DON-HUMBERTO'`, `L3='EGP_CHILE'`).

**`NOMBRES_SUBASTAS` corregido, de punta a punta** (16 posiciones `B:Q`, todas confirmadas contra
el archivo real):

| Letra | Antes (mal) | Ahora (confirmado) |
|---|---|---|
| B | `Control` | **`Concepto`** |
| C | `Sub_Baj` | **`Control`** |
| D | `Fecha` | **`Sub_Baj`** |
| E | `Año` | **`Fecha`** |
| F | `Mes` | **`Año`** |
| G | `Dia` | **`Mes`** |
| H | `Hora_dia` | **`Dia`** |
| I | `Hora_mes` | **`Hora_dia`** |
| J | `Configuración` | **`Hora_mes`** |
| K | `Propietario` | **`Configuración`** |
| L | `Clave horaria` | **`Propietario`** |
| M | `Ciclo` | **`Clave horaria`** |
| N | `Energía SSCC` | **`Ciclo`** |
| O | `FD` | **`Energía SSCC`** |
| P | `FMA` | **`FD`** |
| Q | *(sin nombre)* | **`FMA`** |

Cada nombre "de antes" cayó exactamente una posición más adelante de donde debía — **no** era un
corrimiento uniforme de todo el archivo (la explicación que se venía dando, "la primera columna
del original está vacía"), era que faltaba UNA columna real (`Concepto`) al principio del bloque
de 11 que copia la macro (`DB!B:L → Subastas!B:L`, confirmado con el `MsgBox` de la macro:
`"DB B:L → Subastas B:L"`, `"DB P → Subastas O"`, `"DB Y → Subastas P"`, `"DB V → Subastas Q"`).
`A` (antes de `B`=Concepto) sigue genuinamente sin usar — eso sí estaba bien.

**Consecuencia feliz: casi nada de la lógica ya escrita estaba mal, solo los nombres.**
`construir_subastas()` arma los valores por **posición** (columna 0 del bloque B:L, columna 1,
etc.), y esas posiciones eran correctas — el bug estaba únicamente en qué nombre se le pegaba a
cada posición al final (`NOMBRES_SUBASTAS`). Por ejemplo, la fórmula de "Clave horaria"
(`=Configuración&Dia&Hora_dia`, confirmada con la fórmula real `M3=K3&H3&I3`) ya se calculaba
así en el código (usando las letras internas `K`,`H`,`I` como *posiciones*, que por construcción
coinciden con las letras reales de Excel en el bloque B:L) — solo estaba mal etiquetada como
`"Ciclo"`. Verificado con los valores exactos de dos filas del archivo real (test sintético, ver
abajo): el resultado coincide.

**Cambios de código:**

- `NOMBRES_SUBASTAS`: corregido (tabla de arriba).
- `calcular_subastas_energia_sscc()` → renombrada **`calcular_subastas_ciclo()`**: la columna
  `N` que esta función resuelve **no es "Energía SSCC", es "Ciclo"** — el nombre real de la
  columna que trae el XLOOKUP (`Calculo E Costos!P` = "Ciclo de Carga del mes") coincide
  exactamente con que sea un número de ciclo, no una energía. La lógica/fórmula que ya estaba
  implementada es la correcta — solo el nombre estaba equivocado. `construir_dic_umbrales_
  subastas()` renombra su parámetro `energia_sscc` → `ciclo_subastas` (mismo motivo).
- `construir_dic_reservas_subastas()` (reservas por subasta de `Calculo RE545`, `AC:AU`): usaba
  `df_subastas["Control"]` como criterio de tipo (`CPF(-)`/`CSF(+)`/etc) — **esto sí era un bug
  real, no solo un nombre**: la fórmula real usa `Subastas!$B:$B` como rango de coincidencia
  contra el encabezado de columna (`AC$3="CPF(-)"`), y `Subastas!B` es `Concepto` (el label
  completo), no `Control` (que ahora sabemos que es solo el tipo sin dirección — nunca iba a
  poder distinguir `(-)` de `(+)`). Corregido a `df_subastas["Concepto"]`.
- `construir_prorrata_sscc()`/`construir_dic_prorrata()` (Prorrata SSCC, `AG:AL`): **sin
  cambios** — el pivot ya usaba `columns="Control"`, y ahora que se sabe que `Control` real es
  el tipo sin dirección (`CSF`/`CTF`/`CPF`), sigue siendo exactamente lo que el pivot necesita
  (agrupa por tipo, no por tipo+dirección — coincide con que el VBA original duplique
  literalmente `AJ=AG`/`AK=AH`, sin distinguir dirección en absoluto). Era una decisión correcta
  desde el principio, ahora con más fundamento.

**Pieza extra que este archivo confirmó, sin que hiciera falta pedirla:** la tabla de umbrales
SUBIDA/BAJADA (`AW`, bloqueada durante varias sesiones hasta la "etapa 4") vive en
`Subastas!S:W` (`S`=Configuración, `T`=Ciclo, `U`=Clave, `V`=SUBIDA, `W`=BAJADA) — coincide
exactamente con lo que `construir_dic_umbrales_subastas()` ya calculaba en Python de forma
independiente (no la lee del archivo, la deriva). Buena señal cruzada de que esa parte del
cálculo está bien encaminada.

**Verificación:** nuevo test (`test_subastas_real.py`, no persistido) que arma dos filas de
`DB` con los valores EXACTOS de las filas 3 y 7 del `Libro1.xlsx` real (mismo texto, mismos
números) y confirma que `construir_subastas()` separa correctamente `Concepto`/`Control`/
`Sub_Baj`, ubica `Configuración`/`Propietario` en su lugar, calcula `Clave horaria` igual que la
fórmula real, deja `Ciclo` vacío (se calcula después, en `Pagos_BESS.xlsx`) y copia `Energía
SSCC`/`FD`/`FMA` en las posiciones correctas (`40.6`/`0.8533`/`0` y `74.5`/`1`/`0`, tal cual el
archivo). Se corrigieron los tests de sesiones anteriores que usaban los nombres viejos
(`test_etapa4.py`, `test_re545.py`) para que seteen `Concepto` además de `Control` donde
corresponde. Regresión completa (E Costos etapas 2-4, RE545 etapas 1-4, secciones de
`Pagos_BESS.xlsx`, SoC): pasa.

**No se tocó** (fuera de alcance de esta corrección, quedan igual): `FD`/`E COSTOS`/`Resumen` del
mismo `Libro1.xlsx` — se revisaron de pasada y coinciden con lo ya implementado (`NOMBRES_FD_CSF`,
`NOMBRES_FD_CPF`, `NOMBRES_CALCULO_E_COSTOS`), sin cambios.

**Todavía pendiente** (ver "Pendientes abiertos"): correr un caso real completo con esta
corrección aplicada, para confirmar que ahora sí aparecen filas `L=1` y que Prorratas/reservas de
RE545 no quedan en 0 (la falta de la columna `Concepto` explicaría, retroactivamente, por qué el
caso real de la sesión anterior daba 0 filas `BAJADA`/`SUBIDA`).

---

## 2026-09-11 (17) — Fix: fila de nombres del SoC (fila 2, no la 3) + selector Medidores/Ofertas separado

Dos pedidos cortos del usuario, con un archivo real (`SOC_2607.xlsx`) que resolvió el primero de
punta a punta.

### Fix: `detectar_fila_nombres()` se quedaba con la fila equivocada

El usuario avisó: "el soc sigue sin nada en la fila 2 está el nombre de la central porsi no la
3". Con el archivo real se confirmó la estructura exacta: fila 2 = nombre limpio de la central
(`"SAE-CRCA-PFV-DON-HUMBERTO"`, idéntico al de `Medidas_SAE.xlsx`), fila 3 = la ruta SCADA
completa (`"\\SRV-SCADA-AF2\SEN\Generación\...\SAE-PFV Don Humberto|Nombre"`, la misma pieza que
ya se había resuelto la sesión anterior con `_extraer_nombre_desde_ruta_scada`), fila 4 en blanco,
fila 5 los encabezados `Status/Questionable/Time Stamp/Value`. Dos filas útiles APILADAS (sin
blanco entre medio) antes del hueco en blanco que precede a los encabezados — `detectar_fila_
nombres()` subía desde los encabezados y se quedaba con la PRIMERA fila útil que encontraba
(fila 3, la ruta), sin darse cuenta de que había otra más arriba (fila 2, el nombre limpio).

**Fix:** en vez de devolver la primera fila útil encontrada subiendo, ahora se sigue subiendo
mientras las filas sigan siendo útiles (sin blanco de por medio) y se devuelve la MÁS ARRIBA de
ese bloque contiguo. Con una sola fila útil (el caso más común hasta ahora) el comportamiento es
idéntico a antes — retrocompatible.

**Verificación, con el archivo real completo:** `extraer_soc()` sobre `SOC_2607.xlsx` (sin ningún
`Diccionario`/homologación, `mapa_homologacion={}`) detecta las **9 centrales exactas** de
`Medidas_SAE.xlsx` (`SAE-CRCA-PE-LA-CABANA`, `SAE-CRCA-PFV-ANDES3`, ..., `SAE-TOCOPILLA`),
**26.793 filas**, **cero incidencias**. La misteriosa central suelta `'07 Region RM'` de la sesión
anterior queda explicada: es lo que la fila 3 (ruta SCADA) tiene para el bloque de `SAE-CRCA-
PFV-MANZANO` en particular — un dato incompleto que ya no se usa, porque ahora se lee la fila 2.
Con esto, y salvo que otro archivo real muestre lo contrario, probablemente **ya no hace falta
ningún `Diccionario` para homologar el SoC** — los nombres de fila 2 ya vienen idénticos a
`Medidas_SAE.xlsx`. Se agregó `test_soc_fila_nombres.py` (no persistido) con el caso real (2 filas
apiladas), el caso simple (1 fila, retrocompatibilidad) y un caso con 3 filas apiladas.

### Selector separado: `Medidores` y `Ofertas SSCC` como casillas independientes

Pedido: "separa medidas de ofertas el selector". Hasta ahora `SECCIONES_CONSOLIDADO` tenía una
sola sección `"medidores"` que escribía las dos hojas juntas (`("Medidores", "Ofertas SSCC")`),
porque `construir_medidores()` las arma en una sola pasada (`Medidores!R:S:T` depende de Ofertas
SSCC). Ahora son dos ids separados (`"medidores"` → hoja `Medidores`; `"ofertas_sscc"` → hoja
`Ofertas SSCC`), cada uno decidiendo solo si se REESCRIBE su propia hoja — pero **la lectura
combinada sigue siendo una sola**: alcanza con que cualquiera de las dos esté tildada para que se
lean los 4 archivos de entrada (`Medidas_SAE.xlsx`, SoC, `Centrales.xlsx`, OfertasSSCC) y se corra
`construir_medidores()`; lo que cambia es solo qué hoja(s) se escriben al final (`hojas_
regenerar`, mismo mecanismo de preservación que ya usaban `CMg`/`FD`/`Subastas`). No hacía falta
tocar `Balance_BESS.py`: la ventana ya recorre `SECCIONES_CONSOLIDADO` genéricamente, así que
ahora dibuja 5 casillas en vez de 4 sin ningún cambio de código — solo se agrandó la ventana
(`620x420` → `620x560`) para que entren.

**Verificación:** `test_selector_medidas_ofertas.py` (no persistido), con las funciones de
lectura/cálculo monkeypatcheadas para no depender de archivos Excel completos: confirma que
tildar solo `"ofertas_sscc"` dispara igual la lectura combinada (se llama a `construir_
medidores()`) pero `Medidores` queda preservado tal cual estaba, y viceversa con solo
`"medidores"` tildada. Regresión completa de las 16 sesiones anteriores: pasa. No se probó la
ventana tkinter en sí (sin entorno gráfico en esta sesión, como siempre).

---

## 2026-09-11 (18) — Diagnóstico mejorado: "sin bloque de SoC" ahora dice si es un problema de Diccionario

El usuario corrió de nuevo con el fix de la sesión anterior. Mejoró mucho (de 9 centrales sin
cruzar a 1): `Centrales en Medidas_SAE.xlsx sin bloque de SoC: ['SAE-CRCA-PFV-NUEVO-QUILLAGUA-2']`
— pero avisó que esa central sí está en su archivo de SoC.

**Diagnóstico:** con el `SOC_2607.xlsx` real que había compartido, `extraer_soc()` (aislado, sin
`Diccionario`) SÍ detecta esa central perfectamente — fila 2 trae literalmente
`'SAE-CRCA-PFV-NUEVO-QUILLAGUA-2'`, idéntico a `Medidas_SAE.xlsx`, sin ningún carácter raro (se
revisó con `repr()`, sin espacios/unicode ocultos). Sin `Diccionario`, esta central cruza sola.

**Hipótesis más probable, confirmada como técnicamente posible con una prueba:** el
`Centrales.xlsx!Diccionario` del usuario probablemente tiene una fila para esta central con el
nombre "feo" (estilo ruta SCADA, ej. `"SAE-PFV Nuevo Quillagua II"`) escrito ANTES que el nombre
limpio en esa misma fila. `construir_homologacion()` toma el PRIMER valor de cada fila como
"canónico" — si ese orden quedó así (probablemente porque el `Diccionario` se armó en una época
en que solo se conocía el nombre feo, antes de esta sesión), la homologación **rompe** un cruce
que la fila 2 del SoC ya resolvía sola: convierte el nombre limpio en el feo, y el feo no cruza
contra nada en `Medidas_SAE.xlsx`. Se armó una prueba sintética que reproduce exactamente este
mecanismo (`construir_homologacion()` con una fila `["SAE-PFV Nuevo Quillagua II",
"SAE-CRCA-PFV-NUEVO-QUILLAGUA-2"]` → el nombre limpio homologa hacia el feo).

**No se tocó la lógica de homologación** (cambiar cuál valor de la fila gana como "canónico"
afectaría potencialmente otras centrales que sí dependen del orden actual — cambio de más riesgo
del que amerita una hipótesis todavía sin confirmar con el archivo real del usuario). En cambio,
se mejoró el **diagnóstico**: el aviso "Centrales en Medidas_SAE.xlsx sin bloque de SoC" ahora
busca, para cada central faltante, si existe algún `nombre_scada_original` (el nombre crudo antes
de homologar) que normalice igual a esa central — si lo encuentra, lo dice explícitamente en el
aviso ("el SoC SÍ trae un bloque con nombre crudo [...] — revisar si Diccionario lo está
homologando a otro nombre"). Si el usuario corre de nuevo, el mensaje mismo va a confirmar o
descartar la hipótesis sin necesitar que comparta su `Diccionario`.

**Verificación:** test sintético (`test_soc_diagnostico.py`, no persistido) con la lógica exacta
del bloque nuevo: caso "hay candidato crudo" (muestra la pista) y caso "de verdad no hay bloque"
(mensaje simple, sin inventar pistas falsas). Regresión completa: pasa.

**Pendiente:** confirmar con el próximo aviso (o con el contenido real de `Diccionario`) si la
hipótesis es correcta. Si lo es, la corrección más simple sería reordenar esa fila del
`Diccionario` (poner el nombre limpio primero) — eso ya lo puede hacer el usuario directamente en
su archivo, sin esperar un cambio de código.

---

## 2026-09-11 (19) — Dos bugs reales confirmados y arreglados con la primera comparación fila a fila contra la planilla 11

El usuario compartió, por primera vez, una comparación directa: `Pagos_BESS.xlsx` generado por
Python con una hoja extra pegada a mano (`Ecostos planilla 11`) con los valores reales del
`.xlsm` original para las mismas filas, más su `Centrales.xlsx` real (dos veces) y confirmó que
el aviso mejorado de la sesión anterior efectivamente detectó la pista ("el SoC SÍ trae un
bloque..."). Con estos tres archivos se resolvieron dos bugs reales de una — la primera vez que
el proyecto se valida contra datos reales, no solo sintéticos, y encontró exactamente el tipo de
error que esa validación existe para atrapar.

### Fix 1: `construir_homologacion()` mezclaba tablas de equivalencia distintas

Se leyó el `Diccionario` real completo. Confirmó algo que el plan ya sabía en teoría (sección
4.2/B: "la hoja presenta bloques asociados a FD/Subastas/ofertas") pero que `construir_
homologacion()` nunca implementó correctamente: son **tablas independientes por columnas**
(`A:B`=FD, columnas `C:D` vacías, `E:F:G`=Subastas/ofertas), no una fila = todos los sinónimos de
una central. Para 6 de las 9 centrales las dos tablas coinciden fila a fila por casualidad del
orden en que se cargaron — pero para las últimas 3 (`VICTOR-JARA`, `ANDES4`, `NUEVO-QUILLAGUA-2`)
el orden de la tabla de la derecha está corrido una fila respecto de la izquierda. La función
vieja trataba la fila entera como un solo grupo de sinónimos, así que en esas 3 filas mezclaba
central de una tabla con la central de la fila vecina de la otra tabla — literalmente homologaba
`NUEVO-QUILLAGUA-2` hacia `ANDES4` (la central de la fila anterior en el bloque de la derecha).
Esto explica el aviso persistente que el usuario venía reportando (y que el diagnóstico de la
sesión anterior confirmó correctamente: "el SoC SÍ trae el bloque").

**Pista clave para encontrarlo:** las otras dos funciones que leen esta misma hoja
(`construir_dic_mapeo_diccionario()`, columnas `A:B`; `_mapas_homologacion_fge()`, columnas
`E:F:G`) **ya** usaban posiciones de columna fijas — nunca tuvieron este bug, porque ya estaban
diseñadas sabiendo que son tablas separadas. Solo `construir_homologacion()` (la más vieja de
las tres, escrita antes de que se entendiera bien la estructura de bloques) se había quedado con
el enfoque ingenuo de "toda la fila es un grupo".

**Fix:** `_bloques_columnas_diccionario()` — detecta los bloques de columnas automáticamente
(separador = una columna vacía en TODAS las filas del archivo, no alcanza con mirar una sola
fila porque la fila de encabezados de grupo típicamente solo tiene texto en la primera columna
de cada bloque). `construir_homologacion()` ahora arma el mapa bloque por bloque, sin mezclar
equivalencias entre bloques distintos.

**Verificación:** con el `Diccionario` real completo, las 9 centrales homologan correctamente
hacia sí mismas (antes, `NUEVO-QUILLAGUA-2` homologaba mal). Corrida de punta a punta con el
`SOC_2607.xlsx` real + el `Diccionario` real: 9 centrales, 26.793 filas, **cero** incidencias.
Test sintético (`test_diccionario_bloques.py`, no persistido) que reproduce exactamente la
estructura real (incluida la fila corrida) y confirma retrocompatibilidad con un `Diccionario`
de una sola tabla (sin bloques separados por columnas vacías).

### Fix 2: la Prorrata SSCC (`AG:AL`) devolvía la cuenta cruda, no la proporción

El usuario reportó: "hay algunos pocos casos en los que los controles de frecuencia salen con 2
y otras diferencias donde en la planilla 11 es 0.5 y en la que genera el python es 1 [...] el
problema está en la prorrata de CF". Con la hoja de comparación se encontró la causa exacta:
para un grupo con 1 fila `CPF` + 1 fila `CSF`, la planilla real trae `AG=0.5, AH=0.5` — pero
nuestro código, que hacía `pivot_table(..., aggfunc="count")` y devolvía la cuenta tal cual, daba
`AG=1, AH=1` (la cuenta cruda de cada tipo). Para un grupo con 2 `CPF` + 1 `CSF`, la planilla
real trae `AG=0.6666..., AH=0.3333...` — exactamente `2/3` y `1/3`. El nombre "Prorrata" lo decía
literalmente: es una **proporción** (reparte el 100% del grupo entre los tipos que aparecen), no
una cuenta. El "sale con 2" que reportó el usuario es el mismo bug: cuando había 2 filas del
mismo tipo sin ningún otro tipo compitiendo, la cuenta cruda daba 2 en vez de la proporción
correcta (1.0, ya que 2/2=1).

**Fix:** `construir_prorrata_sscc()` ahora divide cada fila del pivot por la suma de esa misma
fila (entre todos los valores de `Control` presentes), antes de devolverlo — convirtiendo la
cuenta cruda en una proporción que suma exactamente 1 por fila.

**Verificación:** los DOS casos exactos que trajo la comparación real (`1+1 → 0.5/0.5` y
`2+1 → 0.6666../0.3333..`) coinciden ahora al dígito. Test sintético (`test_prorrata_
normalizada.py`, no persistido) con esos dos casos más un caso de una sola fila (retrocompatible,
sigue dando 1.0) y una prueba general de que cada fila del pivot siempre suma 1. Regresión
completa de las 18 sesiones anteriores: pasa (ningún test viejo dependía de un grupo con más de
una fila del mismo tipo, así que ninguno se rompió con la normalización).

**Downstream, sin cambios de código:** `AS`/`AT` (`_calcular_costo_ponderado`, "Energía descarga/
carga con FD") y todo lo que depende de `AG:AL` ya eran fieles a la fórmula real
(`factor = precio1*cantidad1 + precio2*cantidad2`, una suma ponderada, no un promedio) — el bug
estaba enteramente aguas arriba, en qué valores de `cantidad1`/`cantidad2` recibían. Con `AG:AL`
corregidos, `AS`/`AT` deberían salir bien ahora sin tocar esa función.

**Archivos de referencia guardados** (mismo criterio que sesiones anteriores — si es la fuente de
una decisión, vive en el repo): `docs/Centrales_real.xlsx`, `docs/SOC_real_2607.xlsx`,
`docs/Pagos_BESS_comparacion_real.xlsx` (esta última es la primera comparación real vs Python
lado a lado que existe en el proyecto — vale la pena mirarla de nuevo si aparece otra
discrepancia).

**Pendiente:** el usuario todavía no confirmó si con estos dos fixes el resto de la hoja coincide
completamente contra la planilla 11 — falta una corrida nueva de punta a punta.

## 2026-09-11 (20) — Fix: `Calculo RE545!AR:AT` no eran `SUMIFS`, eran la constante 1

El usuario reportó, con una nueva comparación real (`Pagos_BESS.xlsx` con una hoja pegada
`RE545 P11`, análoga a la de la sesión anterior pero para `Calculo RE545`): "En el R545 tengo
diferencias igual parten en AK:AN".

**Primer paso, para no perseguir un fantasma:** como ya se documentó (plan §26.2, BITACORA de
sesiones anteriores), la salida de `Calculo RE545` **no reproduce la letra de Excel real** — las
columnas vacías del original (`W:AB`) no se escriben, así que todo lo que sigue queda corrido de
letra. Comparando por **orden/contenido** (no por letra) los tres bloques de 6 columnas
(`Subastas`, `FD`, `FMA`) contra la hoja `RE545 P11` pegada por el usuario, fila a fila por las
26.784 filas, usando como clave `Mes+Dia+Hora+Minuto+Configuracion`:

- Bloque `Subastas`: 0 diferencias.
- Bloque `FD` (donde el usuario ubicó "AK:AN" en SU comparación, que sí tiene las columnas reales
  sin correr): 0 diferencias.
- Bloque `FMA`: **80.352 diferencias** (positions `CPF(+)`, `CSF(+)`, `CTF(+)` de ese bloque, en
  TODAS las filas).

O sea: la molestia real no estaba en `AK:AN` de nuestra salida (que es el bloque `FD`, sin
diferencias), sino en el bloque `FMA` (que en nuestra salida compacta cae en columnas distintas,
pero el usuario lo estaba mirando alineado con la posición real del `.xlsm`, donde si cae en
`AK:AN`... da igual: la comparación por contenido mostró exactamente dónde estaba el problema real,
sin necesidad de discutir letras).

**Causa exacta**, confirmada contra las fórmulas guardadas del archivo real
(`docs/Calculo_RE545_reducido_para_IA.xlsx`, hoja `Mapa_Formulas`, que lista la fórmula de cada
rango de columnas del `.xlsm` real): `AO:AQ` (`CPF(-)`/`CSF(-)`/`CTF(-)` del bloque `FMA`) SÍ son
`SUMIFS(Subastas!$Q:$Q, ...)`, igual que los otros 15 valores de los tres bloques — pero `AR:AT`
(`CPF(+)`/`CSF(+)`/`CTF(+)`, el mismo bloque) **no aparecen como fórmula en el mapa**: son el
valor literal `1`, constante, en las 26.784 filas del archivo real (confirmado también
directamente en la hoja `RE545 P11` pegada por el usuario: `CPF(+)=CSF(+)=CTF(+)=1` sin ninguna
excepción). El plan (§26.3) asumía "los tres [bloques] son el mismo SUMIFS" — cierto para 15 de
las 18 columnas, falso para estas 3. Nuestro código, al no distinguir el caso, les aplicaba el
mismo `SUMIFS`, que casi siempre da 0 (rara vez hay match exacto central+hora+tipo en `Subastas`
para esas combinaciones) — de ahí el "sale con 2 [en realidad 0] donde debería ser otra cosa" que
reportaba el usuario, y el arrastre a `AU` (`SUMA Reservas*FMA*FD = SUMPRODUCT(...)/4*1000`, que
multiplica por estas columnas).

**Fix:** `calcular_reservas_re545()` — al armar el tercer bloque (`FMA`), las posiciones 3, 4 y 5
(`CPF(+)`, `CSF(+)`, `CTF(+)`) ya no consultan el diccionario de `SUMIFS`: se fijan directamente
en `1.0`, para todas las filas, sin excepción — es dato constante, no una fórmula que dependa de
`Subastas`.

**Verificación:**
- La función nueva es, por construcción, independiente de los diccionarios de `SUMIFS` para esas
  3 columnas (siempre devuelve 1.0 sin mirar `Subastas`) — coincide automáticamente con las
  26.784 filas reales, que también son constantes.
- Recalculando `AU` fila a fila con los valores reales de los otros 15 componentes (que ya
  coincidían) más `AR=AS=AT=1` en vez del `SUMIFS` viejo: **0 diferencias** contra el `AU` real en
  las 26.784 filas (antes del fix: 4.628 filas con diferencia, hasta 36.609 de magnitud).
- Test sintético actualizado/agregado (`test_re545.py`, `test_re545_fma_constante.py`, no
  persistidos) que confirman que `AR`/`AS`/`AT` dan 1.0 sin importar el contenido de `Subastas`
  (incluso con diccionarios vacíos), y que las otras 15 columnas del bloque de reservas siguen
  siendo el `SUMIFS` de siempre, sin tocarse.
- Regresión completa de las 19 sesiones anteriores: pasa (se actualizó a mano el valor esperado
  de un test viejo, `test_re545.py`, que tenía hardcodeado el resultado incorrecto de `AU` para
  un caso con `CSF(+)`; el valor nuevo es el correcto según la fórmula real).

**Archivo de referencia:** `015231ed-Pagos_BESS.xlsx` (hoja `RE545 P11`, pegada por el usuario) —
no se copió a `docs/` porque no aporta nada que `docs/Pagos_BESS_comparacion_real.xlsx` (sesión
anterior) o `docs/Calculo_RE545_reducido_para_IA.xlsx` (ya en el repo, fuente directa de la
fórmula real que confirmó el fix) no tuvieran ya.

## 2026-09-11 (21) — Fix grande: `BK/BL/BS` de `Calculo RE545!BI:CE` cruzaban mal S contra BI

El usuario, todavía mirando la misma comparación real (`RE545 P11`), reportó más diferencias
usando los nombres de columna de **nuestra** salida (no las letras reales): "AR:AT, BI, AZ" y
después, más específico: "Curva Cmg Decendente promedio horario, (blanco), Curva Cmg Decendente".

**Primer chequeo (para no repetir la confusión de letras de la entrada anterior):** comparando por
contenido, `AR:AT` (el fix de la sesión 20) y `BI`/`AZ` con letra REAL (`Orden`, `EiniT` del
resumen) daban 0 diferencias — esos ya estaban bien. El problema real estaba en dos columnas
específicas que el usuario nombró explícitamente: `Curva Cmg Decendente promedio horario` (`BK`
real) y `Curva Cmg Decendente` (`BM` real).

**Causa exacta**, encontrada comparando fila a fila y confirmada contra las fórmulas guardadas del
archivo real (`docs/Calculo_RE545_reducido_para_IA.xlsx`, hoja `Mapa_Formulas`):

```
BK4 = SUMIFS(R:R, G:G,G4, S:S,BI4, T:T,T4, E:E,BJ4)
BL4 = SUMIFS(Q:Q, S:S,BI4, E:E,BJ4, G:G,G4, T:T,T4)
BS4 = IFERROR(INDEX(I, MATCH(1, (BR=BR4)*(S=BI4)*(G=G4)*(E=BJ4), 0)), "")
```

El criterio `S:S,BI4` compara la columna `S` (`ranking cmg`) de las OTRAS filas contra el `BI`
(`Orden`) de LA FILA ACTUAL — no `BI` contra `BI`. `calcular_bk_bl_bm_bs_re545()` armaba una sola
clave usando `BI` de los dos lados (la de acumulación Y la de búsqueda), lo que da el resultado
correcto únicamente cuando `S` y `BI` coinciden fila a fila por casualidad — que es exactamente lo
que pasaba en el único caso sintético que existía hasta ahora (por eso nunca se detectó). Con
datos reales, donde `S` y `BI` difieren, el agrupamiento salía mal.

Esto no se quedaba en `BK`/`BM`: al ser el insumo de `BN` (`Edisp_Asig`), `BO` (`Total C1_545`),
`CC` (`Total C2_545`) y, al final, `CE` (`Monto a compensar` — la ÚLTIMA columna de toda la hoja),
el error se propagaba a lo largo de toda la sección `BI:CE`.

**Fix:** `calcular_bk_bl_bm_bs_re545()` ahora arma DOS listas de claves — `claves_acumulacion`
(con `S` de cada fila, para poblar los diccionarios de suma/primer-valor) y `claves_busqueda` (con
`BI` de cada fila, para leer el resultado) — en vez de una sola clave usada de los dos lados.
También se corrigió el valor por defecto de `BS` sin match: antes daba `KeyError` (nunca pasaba
porque la clave vieja siempre existía en el propio diccionario); ahora, correctamente, da `NA`
(blanco), igual que el `IFERROR` real — a diferencia de `BK`/`BL`, que sin match dan `0` (`SUMIFS`
real).

**Verificación exhaustiva** (con `Pagos_BESS.xlsx` real, hoja `RE545 P11`, 26.784 filas):
- Reconstruyendo `BK`/`BM` a partir de los datos crudos de la propia salida (`Configuracion`,
  `Ventana de valorizacion`, `ranking cmg`, `CMg Promedio`, `CMg`, `Orden`, `Periodo`) y pasándolos
  por la función corregida: **0** diferencias contra los valores reales, en las 26.784 filas.
- Corriendo la etapa completa `BI:CE` (`calcular_componentes_re545()`) con `AU` y el resumen
  `AW:BG` reconstruidos con sus valores REALES (para no arrastrar la contaminación del bug de `AU`
  de la sesión anterior, que también ensuciaba el resumen `AW:BG` generado con el código viejo):
  `Edisp_Asig`, `Total C1_545`, `Total C2_545` y **`Monto a compensar`** (la columna final de toda
  la hoja) dan **0** diferencias en las 26.784 filas.
- Test sintético nuevo (`test_re545_bk_bl_bs_cruzado.py`, no persistido) que reproduce un caso con
  `S != BI` fila a fila y confirma que el resultado cambia respecto del comportamiento viejo (el
  test viejo, `test_re545d.py`, nunca hubiese detectado esto porque su `S` sintético coincidía con
  `BI` por construcción — se le agregó una columna `S` explícita, documentando por qué, para que
  siga siendo válido sin ocultar el hueco de cobertura).
- Regresión completa de las 20 sesiones anteriores: pasa.

**Confirmado también:** la contaminación de `Edisp_Asig`/`Total C1_545`/`Total C2_545` que
aparecía al principio de esta validación (antes de sustituir el resumen `AW:BG` por valores
reales) no era un bug nuevo — era el mismo bug de `AU` de la sesión 20 propagándose a través de
`BV` → `Margen ultima hora`/`Total Reservas* FD *FMA`/`Edisp_T` del resumen. Con los dos fixes (20
y 21) aplicados juntos en el pipeline real (donde el resumen se reconstruye siempre a partir del
`AU` ya corregido), no hace falta ningún fix adicional para esa cadena.

---

## 2026-09-11 (22) — `cmg.xlsx` ahora lo genera el programa (botón "Generar" en su fila del árbol)

**Pedido del usuario:** `cmg.xlsx` (entrada de `Cmg/`) se armaba a mano corriendo un script suelto
(`Extrae_CMG_barras.py`) al lado del CSV. Quería (1) un botón "Generar" al lado del nombre en la
ventana de Balance_BESS, (2) que el CSV de origen se busque en la ruta de red
`T:\CMgReales 15MIN\AAAA\AAMM\Mensual\CMg\Cmg para balance` en vez de al lado del `.py`, y (3) que
las barras a filtrar salgan de `Centrales.xlsx` (hoja `Resumen BESS`) en vez de estar escritas en
el código.

**Lo que se hizo:**

- `nucleo.py`: sección nueva "GENERACION DE cmg.xlsx DESDE EL CSV 15-MINUTAL" con
  `ruta_csv_cmg_15min()`, `barras_desde_resumen_bess()`, `construir_cmg_desde_csv()`,
  `_validar_layout_cmg()` y `generar_cmg(carpeta_base, aamm, ruta_csv=None, ...)` — misma firma
  `registrar`/`progreso` que `generar_consolidado`/`generar_pagos_bess`, así entra sin cambios en
  el `lanzar_generacion()` que ya existe en la ventana.
- Constantes nuevas: `RAIZ_CMG_REALES` (`T:\CMgReales 15MIN`), `SUBCARPETAS_CMG_REALES`,
  `PLANTILLA_CSV_CMG_15MIN`, `SEPARADOR_CSV_CMG`, `CODIFICACION_CSV_CMG`,
  `COLUMNA_CSV_CMG_VALOR`. La letra de unidad queda en UN solo lugar por si cambia.
- Las barras salen de `construir_mapa_barra()` (la misma función que ya alimenta
  `Calculo E Costos!Barra`), no de una lectura nueva: así el filtro del CSV y la homologación
  posterior **no se pueden desincronizar**. Se comparan en mayúsculas (mismo criterio que
  `_buscar_cmg`) pero se conserva el texto tal cual viene del CSV.
- `Balance_BESS.py`: la fila `cmg.xlsx` del árbol lleva su propio botón "Generar". No abre ventana
  con casillas como las dos salidas — no hay nada que elegir. Pide confirmación si el archivo ya
  existe. `pintar_arbol()` guarda ahora las referencias de los botones dibujados dentro del árbol
  (`botones_arbol`), y `terminar()` tolera que ese widget ya no exista (el árbol se repinta entero
  en cada `revisar()`).
- La fila `cmg.xlsx` del diagrama dice además si el CSV de origen del período está disponible o
  no, que es lo que decide si el botón va a poder hacer algo.

**Detalle que importa y es fácil de romper:** `leer_cmg()` lee `cmg.xlsx` **por posición**
(D = Barra, F = valor de Q, H = Cuarto de Hora, I = CMg promedio). El layout que sale de
`construir_cmg_desde_csv()` (columnas del CSV + `Cuarto de Hora` + promedio horario) es justo ese,
pero depende de que el CSV siga trayendo 7 columnas. Por eso `_validar_layout_cmg()` avisa en el
log si `BARRA` deja de caer en D o `Cuarto de Hora` en H, y corta con `ErrorEntrada` si quedan
menos de 9 columnas.

**Lo que NO cambió respecto del script original:** la numeración del `Cuarto de Hora` global sigue
saliendo de los bloques que el CSV realmente trae (no se asumen 96 por día), así que los días de
cambio de hora con 92/100 cuartos siguen funcionando; ahora además se listan en el log los días
que no tienen 24 h.

**Verificación:** caso sintético end-to-end (CSV de 2 días —uno con 23 h—, 3 barras en el CSV y 3
en `Centrales.xlsx`, una de ellas sin datos): genera `cmg.xlsx`, avisa de la barra sin datos,
detecta el día de 23 h, y el archivo resultante se vuelve a leer con `leer_cmg()` +
`construir_dic_cmg()` dando las claves `BARRA|cuarto` esperadas. `python -m py_compile
Balance_BESS.py nucleo.py` pasa. La ventana en sí no se pudo abrir (no hay `tkinter` en el
contenedor de la sesión): el cableado del botón se revisó a mano.

**Pendiente:** correrlo una vez contra el CSV real de la unidad `T:` para confirmar que las barras
de `Centrales.xlsx` están escritas exactamente igual que en el CSV (con el relleno de guiones
bajos, ej. `TOCOPILLA_____110`). Si alguna no coincide, el log lo dice barra por barra.

---

## 2026-09-11 (23) — Reorganización: paquete `Script/`, un botón por fila, salidas desglosadas por hoja

**Pedido del usuario**, seis puntos:

1. los botones a la izquierda del detalle;
2. el cálculo adentro de una carpeta `Script/`, con `nucleo.py` ahí y una subcarpeta `Cmg/`
   con `Extrae_CMG_barras.py` (la idea declarada es ir modularizando `nucleo.py` por etapas);
3. sacar la fila "Periodo (AAMM)" de entre las carpetas del diagrama — el SoC va dentro de
   `Medidas/`;
4. desglosar `Consolidado_entradas.xlsx` como carpeta, una fila por hoja, cada una con botón
   "Actualizar", y eliminar la ventana intermedia de "Generar". Si el archivo no existe, que
   se genere;
5. lo mismo para `Pagos_BESS.xlsx`;
6. el `cmg<AAMM>_def_15minutal.csv` pasa a vivir en la carpeta `Cmg/` del caso, al lado de
   `cmg.xlsx`, con un botón "Traer cmg_15min" que lo baja de la ruta de red.

**Estructura nueva del repo:**

```
Balance_BESS.py
Script/
    __init__.py
    nucleo.py
    Cmg/
        __init__.py
        Extrae_CMG_barras.py
```

`Balance_BESS.py` hace `from Script import nucleo`; `nucleo.py` hace `from .Cmg import
Extrae_CMG_barras` (con fallback a `from Cmg import ...` por si se importa suelto con
`Script/` en el `sys.path`). **Regla nueva, importante para la modularización que viene:**
un módulo de etapa NO importa `nucleo` — recibe rutas y datos, y levanta su propia excepción
(`ErrorCmg`), que `nucleo` traduce a `ErrorEntrada`. Así no hay ciclos de import cuando se
saquen más etapas. El nombre del archivo usa guiones bajos y no espacios para que sea
importable (el usuario lo escribió como "Extrae CMG barras.py").

**Contrato nuevo de `revisar_estructura()`:** antes devolvía tuplas
`(etiqueta, estado, detalle)` y la ventana deducía la profundidad de cada fila mirando el
TEXTO de la etiqueta (`_profundidad_fila()`: ¿termina en "/"?, ¿empieza con dos espacios?).
Eso ya venía frágil y con el desglose por hojas no daba más. Ahora devuelve dicts
`{id, etiqueta, nivel, estado, detalle}`:

- el **nivel** lo pone `nucleo` porque es estructura, no dibujo (que un archivo esté adentro
  de una carpeta, o una hoja adentro de un archivo, no es una decisión de interfaz);
- el **id** es estable y es lo único que la ventana necesita para saber qué botón colgarle a
  cada fila (`_boton_de_fila()`), así `nucleo` sigue sin saber nada de botones.

**Ventana:**

- columnas `Estructura | Estado | Acción | Detalle`. La celda de acción tiene ancho FIJO en
  píxeles (`ANCHO_ACCION`), si no cada fila correría el detalle según el largo de su botón.
- desaparecieron las dos ventanas "Generar" con casillas. Botones: `Traer cmg_15min` y
  `Generar` en `Cmg/`, `Actualizar` en cada fila-hoja de las dos salidas, `Actualizar todo`
  en la fila del archivo.
- mientras corre algo, TODOS los botones del árbol quedan deshabilitados (`corriendo` +
  `habilitar_botones`), y al terminar el árbol se repinta entero (que es lo que los
  rehabilita).
- el campo del período (AAMM) sigue arriba; lo que se sacó es la FILA del diagrama. El SoC
  quedó donde corresponde, colgando de `Medidas/`.

**CMg en dos pasos:** `nucleo.traer_csv_cmg()` copia el CSV de la unidad de red a
`<CARPETA_BASE>/Cmg/` y `nucleo.generar_cmg()` arma `cmg.xlsx` con el CSV que quedó ahí. Se
copia en vez de leer directo de la red a propósito: el caso queda autocontenido (se puede
regenerar `cmg.xlsx` sin la unidad conectada) y queda registrado con qué archivo se trabajó.

**Dos bugs encontrados de paso, los dos arreglados:**

1. **El `Log` de `Consolidado_entradas.xlsx` perdía todos los avisos de preservación.**
   `escribir_salida()` armaba el `df_log` ANTES del bloque `with pd.ExcelWriter(...)`, pero
   los avisos de "no se regeneró la hoja X y no había versión anterior" los agrega
   `_preservar_o_avisar()` DENTRO de ese bloque. Resultado: el `Log` decía "Sin
   observaciones" mientras cuatro hojas quedaban vacías. Nunca se había notado porque hasta
   ahora lo normal era generar todo junto; con un botón por hoja, generar una sola es el caso
   normal y esos avisos son justamente los que hay que ver. El log ahora se arma al final
   (`_armar_log()`).
2. **El estado de una hoja no se podía deducir de que la hoja existiera.** Al generar una sola
   sección, el archivo se crea con TODAS las hojas (las no pedidas, vacías), así que el
   diagrama las mostraba todas como "generada". Se agregó `hojas_con_datos()` (openpyxl en
   `read_only`, `max_row > 1`): una hoja vacía se ve PENDIENTE. El estado de la fila del
   archivo es ahora el rollup de sus hojas ("le faltan hojas por generar").

**Verificación:** caso sintético con `Consolidado_entradas.xlsx` a medias — el árbol
renderizado (niveles, prefijos, estados, botones) sale correcto y el SoC cuelga de `Medidas/`;
`traer_csv_15min` + `generar_cmg` contra una "unidad de red" falsa, con el `cmg.xlsx`
resultante releído por `leer_cmg()`/`construir_dic_cmg()`; `generar_consolidado()` de una sola
sección sobre una carpeta sin el archivo (lo crea, deja el resto de las hojas vacías y ahora
sí las lista en el `Log`); y el error esperado cuando se pide "Generar" sin haber traído el
CSV. `py_compile` de los tres módulos pasa. La ventana en sí no se pudo abrir (no hay
`tkinter` en el contenedor de la sesión): el cableado se revisó a mano y los helpers del árbol
(`_prefijos_arbol`) se probaron aparte, sin tkinter.

**Pendiente que deja esta sesión:** abrir la ventana una vez en Windows para confirmar el
ancho de la columna "Acción" (`ANCHO_ACCION = 150 px`) contra los botones más largos
("Traer cmg_15min", "Actualizar todo") y el alto de fila (`ALTO_ACCION = 26 px`).

---

## 2026-09-11 (24) — Encabezados de grupo (celdas combinadas) en Pagos_BESS.xlsx

> **Nota de integración:** esta entrada se escribió en paralelo a las (22) y (23) (rama
> `claude/brave-wozniak-24n138`, PR #11), en las dos se numeró como "(22)". Al fusionar se
> renumeró a (24), que es el orden real en que entró a `main`. El trabajo de esta sesión es
> independiente del de las otras dos y no se pisan: acá se tocó `escribir_pagos_bess()`; allá,
> la estructura del repo y la ventana. Dos cosas que esta entrada menciona cambiaron de lugar
> con la (23): `nucleo.py` ahora es `Script/nucleo.py`, y el `py_compile` de verificación es
> `python -m py_compile Balance_BESS.py Script/nucleo.py Script/Cmg/Extrae_CMG_barras.py`.


El usuario pidió agregar "los encabezados de celdas combinadas que van arriba de algunas hojas",
como en el libro real. Pregunté alcance (Calculo RE545 / Calculo E Costos / hojas de
`Consolidado_entradas.xlsx`) — contestó "todas las anteriores", respetando la lógica que ya
funciona y siguiendo la estructura del `Libro1` que había entregado antes.

**Investigación primero, antes de tocar código** (todos los archivos reales ya en `docs/`):
- `docs/Calculo_RE545_reducido_para_IA.xlsx` (hoja "Calculo RE545 reducido"): confirma, en la
  fila 3 del archivo real, los títulos de grupo `Dia`/`Nombre`/`BESS` (sobre `Mes:Hora`,
  `Configuracion:Barra`, `Descarga:SoC%`) y `Subastas`/`FD`/`FMA` (sobre los tres bloques de
  reservas) y `Componente 1`/`Componente 2` (sobre partes de `BI:CE`) — el merge en sí se perdió
  en la reducción, pero el texto queda solo en la celda más a la izquierda de cada grupo, igual
  que en un merge real.
- `docs/Libro1_Subastas_real.xlsx` (hoja "E COSTOS", agregada en una sesión anterior para el fix
  de `NOMBRES_SUBASTAS` pero nunca revisada a fondo hasta ahora): es la ÚNICA referencia real que
  tenemos con los merges de Excel intactos (`ws.merged_cells.ranges`) — confirma exactamente los
  mismos 8 grupos que después aparecen (sin merge, solo texto) en Calculo RE545: `Dia`(A2:C2),
  `Nombre`(G2:H2), `BESS`(I2:K2), `Componente 2`(S2:U2), `Prorratas (-)`(AG2:AI2),
  `Prorratas (+)`(AJ2:AL2), `FD`(AM2:AR2), `Componente 1`(AU2:AX2).
- Las otras dos hojas del mismo archivo ("FD", "subastas") NO tienen merges ni título de grupo —
  confirmado, no se les agregó nada. No hay archivo real de referencia para "Medidores"/"CMg"/
  "Ofertas SSCC" con títulos de grupo, así que tampoco se tocaron esas hojas de
  `Consolidado_entradas.xlsx` — agregar algo ahí sería inventar sin dato real, contra la
  metodología del proyecto.

**Trampa importante:** las letras reales de estos merges (ej. `S2:U2`, `AG2:AI2`) NO son las
letras de nuestra salida — la hoja de salida nunca reprodujo la letra de Excel real, solo el
orden y el contenido (documentado desde antes, ver plan §26.2). Cada grupo se definió entonces
como `(etiqueta, [claves internas del diccionario NOMBRES_CALCULO_XXX])`, y la posición real en
la hoja se calcula recién al escribir, buscando esas claves en `list(NOMBRES_CALCULO_XXX)` — así
el grupo cae en la columna que le toca en NUESTRO orden, sea cual sea. Se verificó que las claves
de cada grupo real (ej. `S`,`T`,`U` para "Componente 2") caen CONTIGUAS en nuestro orden interno
también (aunque a veces con nombres repetidos entre grupos, como `CPF(-)`/`CSF(-)`/`CTF(-)`, que
aparecen tanto en "Prorratas (-)" como en el bloque "FD" — por eso los grupos se identifican
siempre por clave interna, nunca por nombre de columna, que no es único).

**Implementación:**
- `GRUPOS_CALCULO_E_COSTOS` y `GRUPOS_CALCULO_RE545` (nuevas constantes, mismo estilo que
  `NOMBRES_CALCULO_XXX`): 8 grupos cada una.
- `_escribir_encabezados_grupo(ws, columnas_internas, grupos, fila=1, columna_inicio=1)` (nueva
  función): calcula la columna real de cada grupo buscando sus claves en `columnas_internas` y
  hace el merge con openpyxl (`ws.merge_cells(...)`) si el grupo tiene más de 1 columna.
- `escribir_pagos_bess()`: ahora escribe ambas hojas con `startrow=1` (deja la fila 1 libre) y
  llama a `_escribir_encabezados_grupo()` después de cada `to_excel()`. **Cambio de estructura**:
  antes la fila 1 tenía los nombres de columna y los datos arrancaban en la fila 2; ahora la fila
  1 tiene los títulos de grupo (celdas combinadas, con huecos donde no hay grupo), la fila 2 tiene
  los nombres de columna, y los datos arrancan en la fila 3. La tabla resumen de RE545 (`AW:BG`,
  que se escribe al lado del bloque principal) también se corrió a `startrow=1` para que sus
  nombres de columna queden alineados con los del bloque principal — no tiene título de grupo
  propio, no hay evidencia real de uno.
- De paso, se corrigió `completar_calculo_e_costos_grupos()`: devolvía `df.rename(columns=
  NOMBRES_CALCULO_E_COSTOS)` sin seleccionar antes por `list(NOMBRES_CALCULO_E_COSTOS)` (a
  diferencia de `renombrar_calculo_re545()`, que sí lo hacía) — no causaba ningún bug visible
  porque `df` nunca traía columnas de más, pero dejaba el orden de salida atado al orden de
  asignación interno en vez del orden documentado, y sin esa garantía los encabezados de grupo
  (que dependen de la posición) se podían desalinear. Ahora usa el mismo patrón `df[list(...)]
  .rename(...)` que RE545.

**Verificación:**
- Test nuevo (`test_encabezados_grupo.py`, no persistido): escribe una hoja completa de cada una,
  confirma que los 8 merges de cada hoja existen, no se pisan entre sí, caen exactamente en la
  columna esperada (recalculada independientemente a partir de `GRUPOS_CALCULO_XXX`), que las
  columnas sin grupo (`AZ`/Monto a compensar en E Costos; `AU`/`BI`/`BJ`/`BQ`/`BR`/`CE` en RE545)
  quedan sueltas sin merge, y que los nombres de columna/datos quedaron en las filas 2/3.
- `test_secciones_pagos.py` (existente) actualizado: los `assert ws.cell(row=2, ...)` que
  apuntaban a datos pasaron a `row=3` (fila 2 ahora son nombres de columna); se agregaron
  aserciones nuevas para los encabezados de grupo. De paso se corrigió un problema en el fixture
  del propio test: `pd.DataFrame({c: [v] for c in NOMBRES_XXX.values()})` colapsa los nombres
  reales duplicados (dict comprehension sobre `.values()`), armando una hoja sintética más
  angosta que la real — no afectaba nada antes de esta sesión porque no se validaba la posición
  de nada, pero desalineaba los encabezados de grupo. Reemplazado por un helper que selecciona
  primero por clave interna (`list(NOMBRES_XXX)`) y recién ahí renombra, igual que la producción.
- Regresión completa de las 21 sesiones anteriores: pasa.

**Pendiente:** no se agregó nada a `Consolidado_entradas.xlsx` (Medidores, Ofertas SSCC, CMg, FD,
Subastas) por falta de un archivo real de referencia con títulos de grupo para esas hojas — si el
usuario comparte uno, se puede repetir el mismo mecanismo (`_escribir_encabezados_grupo()`, ya
genérico) para esas hojas también.

## 2026-09-11 (25) — Nuevo documento: `docs/Estructura_Archivos_Reales.md`

El usuario pidió que los archivos Excel reales que manda (de entrada del proceso, o de
comparación/validación) queden bien descritos en algún `.md`, específicamente dudando si
`Centrales.xlsx` estaba bien documentado. Motivo explícito: va a abrir otra sesión de chat y no
quiere tener que re-adjuntar archivos ni re-explicar cosas ya dichas — la sesión nueva tiene que
entender la estructura real "a la primera".

**Diagnóstico:** la información SÍ existía, pero repartida y parcialmente desactualizada —
`docs/Plan_Traspaso_Python_Balance_BESS.md` secciones 3.4/3.5/3.6 (las que describen
`Medidas_SAE.xlsx`/`SOC_AAMM.xlsx`/`Centrales.xlsx`/Ofertas SSCC) son de las primeras del
documento, escritas ANTES de que se validara nada contra datos reales, y solo algunas tienen una
nota de "corrección" pegada al final (la de `Diccionario` sí, gracias a una sesión anterior; la de
`SOC` no tanto). Para reconstruir el resto de una sesión nueva había que rastrear `BITACORA.md`
sesión por sesión — exactamente lo que el usuario no quiere tener que hacer, y exactamente lo que
esta sesión no quiere que la SIGUIENTE tenga que hacer.

**Qué se hizo:** documento nuevo, `docs/Estructura_Archivos_Reales.md`, escrito de cero (no
parcheado sobre el plan viejo) como referencia única y actual de CADA archivo Excel real del
proyecto, con dos secciones:

- **A. Archivos de entrada del proceso** (los 7 que arma el usuario en la carpeta del caso:
  `Medidas_SAE.xlsx`, `SOC_AAMM.xlsx`, `Centrales.xlsx` [`Resumen BESS` + `Diccionario`],
  `*OfertasSSCC*`, `cmg.xlsx`, `SSCC_Desempeño_*.xlsx`, `3_REMUNERACIÓN_SUBASTAS_E_ID_*.xlsx`).
  Para cada uno: ubicación/patrón de nombre, hoja(s), fila exacta donde arrancan encabezados y
  datos, columnas reales en orden (tabla), trampas ya confirmadas, función de `nucleo.py` que lo
  lee, y el archivo real guardado en `docs/` que lo respalda (si hay). Reconstruido leyendo
  directo con `openpyxl` los 5 archivos reales que ya teníamos guardados
  (`docs/Centrales_real.xlsx`, `docs/SOC_real_2607.xlsx`, `docs/Libro1_Subastas_real.xlsx` [hojas
  `FD`/`subastas`/`E COSTOS`]) más el código ya funcionando de `nucleo.py` para los dos archivos
  sin copia real guardada (`Medidas_SAE.xlsx`, `*OfertasSSCC*`, `cmg.xlsx` — marcados con ⚠️ donde
  la estructura sale del código y no de un archivo real visto).
- **B. Archivos de referencia real para validar la salida** (las hojas de comparación tipo
  "planilla 11" que el usuario pega/adjunta cuando reporta una diferencia): qué trae cada uno de
  los 5 archivos ya guardados en `docs/` y el método de comparación (por clave/nombre de columna,
  nunca por letra de Excel — la salida propia nunca reprodujo la letra real).
- **C.** una guía corta de qué hacer si mandan un Excel nuevo no descrito acá.

**Corrección real encontrada al escribir la sección de `Centrales.xlsx!Diccionario`:** el ejemplo
que se iba a poner tenía un typo de transcripción (`"BESS PE LA CAÑADA"` en vez del valor real
`"BESS PE LA CABAÑA"`) — se relee el archivo real (`docs/Centrales_real.xlsx`) letra por letra
antes de cerrar el documento, en vez de confiar en la memoria de la sesión, y se corrige.

**Enganchado a `README.md`** (fila nueva al principio de la tabla de documentación, con el caso de
uso explícito) **y a `METODOLOGIA.md`** (fila nueva en "el set de documentos" + una regla nueva:
"si el usuario menciona o adjunta un archivo Excel real, leer `docs/Estructura_Archivos_Reales.md`
primero" — para que una sesión nueva lo encuentre solo con la regla de expansión de contexto
normal, sin que el usuario tenga que señalarlo).

**Pendiente:** no hay archivo real guardado todavía para `Medidas_SAE.xlsx`, `*OfertasSSCC*` ni
`cmg.xlsx` — si el usuario comparte alguno, guardarlo en `docs/` y sacar el ⚠️ de esa sección
(pasa a ✅).

---

## 2026-09-11 (26) — Fusión de la rama de encabezados de grupo con la reorganización en `Script/`

El PR #11 (`claude/brave-wozniak-24n138`, encabezados de grupo en `Pagos_BESS.xlsx` +
`docs/Estructura_Archivos_Reales.md`) salió de `main` ANTES de las sesiones (22) y (23) y quedó en
conflicto. Se fusionó `main` dentro de esa rama.

**Conflictos reales: uno solo, `BITACORA.md`.** Las dos ramas agregaron su entrada al final y las
dos la numeraron "(22)" (y la segunda de esa rama, "(23)"). Se conservaron las cuatro entradas y
las de la rama se renumeraron a (24) y (25) — el orden real en que entran a `main` — con una nota
de integración explicando el renumerado. Ninguna entrada vieja se tocó.

`Script/nucleo.py` lo fusionó git solo (detectó el rename `nucleo.py` → `Script/nucleo.py` y los
dos lados tocaban partes distintas de `escribir_pagos_bess()`: allá los encabezados de grupo, acá
el mensaje de preservación). El resto de los archivos, automático.

**Lo que SÍ hubo que arreglar a mano es una interacción que ninguna de las dos ramas podía ver
sola:** `_copiar_hoja_existente()` copia valores y anchos de columna, pero **no** las celdas
combinadas. Con los encabezados de grupo de la rama (24) y el botón "Actualizar" por hoja de la
(23), cada vez que se actualiza UNA de las dos hojas de `Pagos_BESS.xlsx` la otra se preserva…
perdiendo la combinación de sus encabezados (el texto quedaba, la combinación no). Antes esto no
se notaba porque preservar era la excepción; ahora es el caso normal. Se agregó el copiado de
`merged_cells.ranges`.

**Verificación:** test dedicado a esa interacción — se generan las dos hojas (8 rangos combinados
cada una), se actualiza solo `Calculo RE545` y se comprueba que `Calculo E Costos` conserva sus 8
rangos y el texto de los grupos. Se confirmó además, monkeypatcheando la versión vieja de
`_copiar_hoja_existente()`, que sin el fix quedaban en **0** (no es una suposición: es el
comportamiento medido). Más la regresión de las sesiones (22) y (23) (árbol del diagrama,
`Traer cmg_15min` + `Generar`, `generar_consolidado()` de una sola sección) y `py_compile` de los
tres módulos.

También se actualizó `docs/Estructura_Archivos_Reales.md` (que nació en la rama (25) y cuyo
propósito es ser la referencia AL DÍA): su §5 ahora documenta los dos archivos de `Cmg/` —el CSV
15-minutal y el `cmg.xlsx` derivado—, cómo se generan, y la trampa de que `leer_cmg()` los lee por
posición; y las rutas `nucleo.py` pasaron a `Script/nucleo.py`.

## 2026-09-11 (27) — `Medidas_SAE.xlsx` ahora lo genera el programa (paquete `Script/Medidas/`)

> **Nota de integración:** las entradas (27) y (28) se escribieron en paralelo a las (24)-(26)
> (rama `claude/dazzling-gauss-jjxo0q`) y las dos ramas las habían numerado (24) y (25). Al
> fusionar se renumeraron al orden real en que entran a `main`. Los trabajos son independientes
> y no se pisan: allá, los encabezados de grupo de `Pagos_BESS.xlsx`; acá, `Medidas_SAE.xlsx`.


Misma idea que con `cmg.xlsx` (sesión 22), aplicada a la primera entrada del caso. El usuario
entregó los cuatro scripts sueltos que se corrían a mano uno detrás de otro (autor original:
Freddy.Arriagada) y pidió seis cosas:

1. el Excel de homologación pasa a `Auxiliares/`, al lado de `Centrales.xlsx`;
2. `Medidas_SAE.xlsx` tiene un botón "Actualizar" que corre todo de un viaje;
3. la lista `FILTROS_TOPOLOGY` del script 3 sale del código y pasa a `Centrales.xlsx`; **deja de
   ser un reemplazo**: esas centrales se sacan del archivo de homologación, así que el paso 3 las
   **agrega**, con la clave que indique la lista;
4. la hoja nueva de `Centrales.xlsx` la diseño yo y él la crea;
5. los códigos van en `Script/Medidas/`, hermana de `Script/Cmg/`;
6. las salidas intermedias no se ven en la ventana.

**El hallazgo que ordenó todo:** el paso 2 terminaba escribiendo exactamente las 9 columnas de
`nucleo.COLUMNAS_AI`, en el mismo orden (`Mes, Dia, Hora, Minutos, Hora Mes, Cuarto de Hora,
clave, intervalo, Gen_Unidad`). O sea que esta cadena ya era, sin saberlo, el generador de la
entrada que hasta ahora había que dejar a mano en `Medidas/`.

**Estructura:** `Script/Medidas/` con `comun.py` (`ErrorMedidas` + helpers de texto),
`Homologacion.py`, `Descarga_PRMTE.py`, `Claves_Balance.py` y `Generacion_Real.py` — uno por
script original. Se mantiene la regla de `Cmg/`: un módulo de etapa **no importa `nucleo`**;
recibe rutas y datos y levanta su propia excepción, que `nucleo` traduce a `ErrorEntrada`.
`nucleo.generar_medidas_sae()` orquesta los cuatro pasos.

**La hoja `Medidas API` de `Centrales.xlsx`** (diseñada acá, la crea el usuario): `topologyName`
(el nombre exacto de la API), `clave` (con qué nombre aparece en `Medidas_SAE.xlsx`) y `Factor`
(opcional, 1 por defecto; es el equivalente del `Flujo` del archivo de homologación, `-1` para los
retiros). Dos filas pueden apuntar a la misma clave: se suman. La hoja entera es opcional — sin
ella no se agrega ninguna central por ese camino y el resto corre igual.

**Decisión de diseño que importa:** el `Cuarto de Hora` es un índice global del mes que después
cruza contra `CMg`, así que las dos fuentes **comparten un solo calendario** (el que arma
`Claves_Balance`, numerado por `intervaloUtc` porque la hora local se repite en el cambio de hora).
Si cada fuente numerara por su cuenta, un día de cambio de hora las desalinearía en silencio. La
API de operación real no entrega UTC, así que sus filas se pegan por hora local: en un día de
cambio de hora hacia atrás hay ambigüedad real, se toma la primera ocurrencia y se avisa en el
log. Si el cruce da 0 coincidencias, se corta con un error explícito en vez de escribir un
`Medidas_SAE.xlsx` al que le faltan esas centrales.

**Credencial:** los scripts traían el `user_key` escrito adentro (vacío en el 1, `"-"` en el 3).
Primero lo saqué a un campo de la ventana guardado en `config.json`, señalando que al ser una
credencial no debería quedar versionada; **el usuario decidió dejarlo en el código** y se hizo así.
Queda en una sola constante, `USER_KEY` en `Script/Medidas/comun.py`, para las dos APIs — el
problema real que tenía era estar repetido en dos archivos y con valores distintos. Consecuencia
asumida, anotada en `METODOLOGIA.md` §5: queda versionada, así que el repositorio no puede volverse
público sin rotarla antes.

**Dos bugs de los scripts originales, arreglados** (no estaban en el pedido):

1. **Los lotes descargados no llevaban el período en el nombre.** El paso 2 hacía
   `glob("medidas_batch_*.parquet")` y el paso 1 anotaba los puntos ya procesados en un archivo
   único, así que correr dos meses en la misma carpeta mezclaba los lotes de los dos y daba por
   procesados puntos de otro mes. Ahora los dos llevan el período en el nombre. (De paso se fue el
   `punt∟os_procesados.txt`, con un carácter raro en medio del nombre.)
2. **El umbral de "punto de medida completo" era la constante `2976`** (= 31 × 96). Está mal para
   cualquier mes de 30 días o menos —descartaría todos los puntos— y para los meses con cambio de
   hora. Ahora es la cantidad de cuartos de hora que el mes descargado realmente trae
   (`intervaloUtc` distintos), que es lo mismo que `2976` pretendía ser.

**Lo que se sacó a propósito:** el `log_inconsistencias_medidas.xlsx` del script 3, que comparaba
el criterio de desempate viejo contra el de mayor `idMeasure`. Era una investigación ya cerrada (el
propio script titula esa sección "CRITERIO DEFINITIVO"). Lo que sí queda en el log es cuántos
grupos venían duplicados. Los diagnósticos del paso 2 (`reporte_medidas_consolidadas.xlsx`) también
se mudaron al log, por el punto 6 del pedido. Los lotes y la marca de reanudación viven en
`<CARPETA_BASE>/Medidas/_trabajo/`, que la ventana no muestra.

**Refactor menor:** `_leer_resumen_bess()` se generalizó en `_leer_hoja_con_encabezado(ruta, hoja,
columnas_buscadas)` para poder leer la hoja nueva con el mismo criterio (buscar la fila de
encabezados en vez de asumir la primera, porque las hojas reales traen un título arriba).
`_leer_resumen_bess()` quedó como un caso particular.

**Verificación:** test sintético de punta a punta con DataFrames con la forma de las dos respuestas
de API (monkeypatch de `Descarga_PRMTE.descargar` y `Generacion_Real.descargar_mes`; el resto del
proceso es puro pandas). Dos días, tres puntos de medida —uno incompleto a propósito—, tres
centrales en la hoja `Medidas API` —dos apuntando a la misma clave con factores opuestos— y un
grupo duplicado para probar el desempate. Resultados verificados a mano: el punto incompleto se
descarta (187 vs 192 cuartos), `SAE-UNO` da −1 por cuarto (canal 1 con `Flujo` +1 menos canal 3 con
`Flujo` −1), `SAE-ANDES-III` da 0 (inyección + retiro), `PFV-ANDES-IV` da 192 + 4 del desempate, la
central que no está en la lista no aparece, las tres claves comparten los cuartos 1..192, y el
archivo resultante lo relee `leer_medidas_sae()` con las columnas exactamente iguales a
`COLUMNAS_AI`. Más el caso sin la hoja `Medidas API` (devuelve lista vacía y el proceso sigue), la
regresión de las sesiones 22 y 23, y `py_compile` de todos los módulos.

**Lo que NO se pudo probar:** las dos funciones que hablan con la API (no hay red ni credencial en
el contenedor) y la ventana (no hay `tkinter`). Ver pendientes.

**Pendientes que deja esta sesión:**

- Correr el botón una vez contra la API real: confirmar la forma de la respuesta (que
  `mediciones` traiga `intervalo`/`intervaloUtc`/`canalVal1`/`canalVal3`/`principal`, y que el
  filtro de puntos incompletos siga descartando lo mismo que antes), y que el `intervalo` de las
  dos APIs sea el INICIO del cuarto de hora en las dos — de eso depende el cruce contra el
  calendario, y si no coincidiera el log lo va a decir fuerte ("ninguna fila cruza").
- Confirmar con el usuario si `Factor` hace falta o si todas las centrales de la hoja van con 1.
  El script original no aplicaba ningún signo; lo agregué porque, sin él, un retiro no tiene cómo
  expresarse — pero con el archivo real puede resultar que la API ya entregue el signo.

---

## 2026-09-11 (28) — La homologación de Gen real se muda al archivo de homologación

> **Nota de integración:** las entradas (27) y (28) se escribieron en paralelo a las (24)-(26)
> (rama `claude/dazzling-gauss-jjxo0q`) y las dos ramas las habían numerado (24) y (25). Al
> fusionar se renumeraron al orden real en que entran a `main`. Los trabajos son independientes
> y no se pisan: allá, los encabezados de grupo de `Pagos_BESS.xlsx`; acá, `Medidas_SAE.xlsx`.


El usuario entregó el `Homologacion ClavesTF y PRMTE.xlsx` real y cambió de opinión sobre dónde va
la lista de centrales de operación real: **no** en `Centrales.xlsx` (como se había hecho en la
sesión 24) sino en el **mismo archivo de homologación**, en una hoja `Gen real`, "con la misma info
que ahí: clave, Punto de Medida, Canal, Flujo". Tiene sentido: es homologación igual que `homol`, y
así se mantiene con el mismo archivo en vez de repartida en dos.

**El archivo real** (útil para futuras sesiones): una sola hoja `homol`, 74 filas, columnas
`clave | Punto de Medida | Canal | Flujo`. `Punto de Medida` es el `idPuntoMedida`
(`DNHUMBER_033_FB1_EGP`), `Canal` es el `slugCanal` (`kWhD` / `kWhR`) y `Flujo` es ±1. 7 claves,
37 puntos de medida. Los lectores se probaron contra él antes de tocar nada.

**Lo que se movió:** `nucleo.leer_medidas_api()` y la constante `HOJA_MEDIDAS_API` desaparecen;
ahora es `Homologacion.leer_gen_real()` + `HOJA_GEN_REAL`. `generar_medidas_sae()` ya **no necesita
`Centrales.xlsx` para nada** (era su única dependencia con ese archivo). En el diagrama, el archivo
de homologación se desglosa por hojas igual que `Centrales.xlsx`: `homol` (obligatoria) y
`Gen real` (opcional, se ve PENDIENTE si no está).

**La decisión que tuve que tomar solo, porque el usuario está fuera:** de las cuatro columnas, tres
se leen solas (`clave` es la clave del balance; `Flujo` es el ±1 que en la sesión 24 se llamaba
`Factor`; `Punto de Medida` tiene que ser el `topologyName` de la API de operación real, que es lo
único que identifica a una central en esa API). La que no tiene lectura obvia es **`Canal`**: la
API de operación real no expone canales. Se acepta la columna (para que la hoja tenga la misma
forma que `homol`, que es lo que se pidió) pero **no se usa** para nada. Queda anotado como
pendiente por si tenía que significar algo.

**Verificación:** el test sintético de punta a punta de la sesión 24, con el `Gen real` ahora en el
archivo de homologación y con las columnas en el orden real (`clave` primero): mismos resultados
exactos que antes (`SAE-UNO` −192, `SAE-ANDES-III` 0 por inyección+retiro, `PFV-ANDES-IV` 196 con
el desempate), o sea que mover la hoja no cambió ningún comportamiento. Más: `leer_homologacion()`
y `leer_gen_real()` contra el archivo real subido (74 filas / hoja ausente → lista vacía), el árbol
renderizado con ese archivo en `Auxiliares/`, y la regresión de CMg. `py_compile` de todo.

**Pendiente nuevo:** confirmar si la columna `Canal` de la hoja `Gen real` tiene que significar
algo. Hoy se ignora.

---

## 2026-09-11 (29) — Fusión a `main` de la rama de Medidas

Se fusionó `claude/dazzling-gauss-jjxo0q` (entradas 27 y 28) a `main`, que mientras tanto había
recibido el PR #11 (entradas 24-26).

**Conflicto real: otra vez solo `BITACORA.md`**, y por la misma razón de siempre — las dos ramas
agregan su entrada al final y las dos numeran desde donde estaba `main` cuando salieron. Las de
esta rama se renumeraron a (27) y (28) con su nota de integración. Ninguna entrada vieja se tocó.
Todo el código (`Script/nucleo.py` incluido, donde una rama tocó `escribir_pagos_bess` y la otra
agregó `generar_medidas_sae`) lo fusionó git solo.

**Verificado que sobrevivieron las dos mitades:** los encabezados de grupo de `Pagos_BESS.xlsx`
(rama 24), el paquete `Script/Medidas/` (ramas 27-28), y el copiado de celdas combinadas al
preservar una hoja (entrada 26). Los cuatro tests sintéticos pasan, incluido el de la interacción
entre ramas (actualizar una hoja de `Pagos_BESS.xlsx` conserva los 8 rangos combinados de la otra).

**`docs/Estructura_Archivos_Reales.md` actualizado** — es el documento que existe para estar al día,
así que se le agregó lo que cambió con las entradas 27-28: `Medidas_SAE.xlsx` ahora dice cómo se
genera, y se agregó la sección 3.b del archivo de homologación, con las dos hojas (`homol`
confirmada contra el archivo real: 74 filas, 7 claves, 37 puntos de medida, canales `kWhD`/`kWhR`;
`Gen real` todavía sin archivo real, estructura definida en la conversación) y la advertencia de
que una central no puede estar en las dos hojas: el paso de operación real **agrega**, no
reemplaza, así que se contaría dos veces.

---

## 2026-09-12 — Las subastas salen de su origen real (los Access), y "Calcular" en Pagos_BESS

Cinco pedidos del usuario en una: renombrar los botones de `Pagos_BESS.xlsx`, y sobre todo dejar
de sacar las subastas de la planilla 3 para sacarlas de donde realmente vienen. Entregó como
referencia el script `entradas_sscc.py` (el suelto que corre a mano, autor original
Gerardo.Vieyra), su `archivo_de_configuracion.yaml`, un documento de trazabilidad
(`subastas_AAMM.xlsx` → `DB!B:K`) y la salida real de ese script para marzo 2026.

**1) Botones de `Pagos_BESS.xlsx`: "Calcular" en vez de "Actualizar"** (`Balance_BESS.py`). La
fila del archivo dice ahora **"Calcular todo"** y la de cada hoja **"Calcular"**. No cambió
ninguna función: solo el texto del botón y el diagrama del docstring. Los de
`Consolidado_entradas.xlsx` siguen diciendo "Actualizar" (son entradas que se refrescan, no
cálculos).

**2) `Script/Subastas/Ofertas_Adjudicadas.py` (paquete nuevo).** Mismo patrón que `Script/Cmg/`:
sabe todo lo de los Access y no importa `nucleo`. Replica la rutina `calc_subastas` del script
viejo, incluida la parte delicada: por cada día se parte del Access del PO (hora 0, sin sufijo) y
cada PID de la hora `HH` **reemplaza las horas >= HH**, dejando las anteriores como estaban; de
qué archivo salió cada fila queda en `Hora_PID`. Después, los mismos filtros: fuera las filas con
`CANTIDAD MW = 0`, precio 0 fuera de la banda 1, `CANTIDAD PONDERADA MW` vacía completada con
`CANTIDAD MW`, y duplicados fuera.

Cambios respecto del script viejo, todos pedidos: los `.accdb` no se leen de la red sino que se
copian primero a la carpeta del caso; el período es el AAMM de la ventana y no un `.yaml`; y no
escribe ningún Excel. `pyodbc` se importa **dentro** de las funciones que leen, no arriba: es la
única dependencia del proyecto que solo existe en Windows (necesita el *Microsoft Access Database
Engine*) y no tiene por qué romper el import de todo el programa donde no esté.

**3) `Subastas/DB subastas/` + botón "Traer subastas".** La carpeta la crea el programa si no
existe (pedido explícito), en `revisar_estructura()`. El botón (`nucleo.traer_subastas`) copia de
`\\nas-cen1\Estadisticas\progdiar_adjudicaSEN\` todos los `OfertasSSCCAdj*` del período —hasta
744 por mes— salteando los que ya están copiados y no cambiaron (mismo tamaño y misma fecha). La
ruta de red es la segunda y última del programa que apunta fuera de la carpeta base del caso, y
está en una sola constante (`RAIZ_SUBASTAS_ORIGEN`), igual que la del CSV de CMg.

**4) `construir_subastas_desde_accdb()` arma la hoja `Subastas` del consolidado** con las mismas
columnas `B:Q` de siempre —el formato no cambió, solo de dónde sale—. Siguiendo el documento de
trazabilidad: `Concepto` = `SERVICIO` tal cual; `Control` = `SERVICIO[:3]`; `Sub_Baj` mirando el
penúltimo carácter (no "si contiene un +": se replicó al pie de la letra
`=IF(LEFT(RIGHT(C,2),1)="+",...)`); `Fecha` = `DATE(Año,Mes,Dia)`; `Hora_dia` = `HORA` **sin
sumar ni restar 1**; `Hora_mes` = `(Dia-1)*24 + Hora_dia`; `Configuración` directa. `Ciclo` sigue
vacía acá (se calcula en `Calculo E Costos`, como antes).

La planilla 3 **no se borró**: quedó como respaldo y se usa solo si `DB subastas/` no tiene ningún
Access del período, avisándolo en el log. Así los casos viejos siguen andando. Su fila del
diagrama dejó de ser un "FALTA" que bloquea.

**5) `Propietario` desde `Centrales.xlsx`** (punto 7 del usuario: lo dejó en la columna B de
`Resumen BESS` y corrió todo lo demás una columna a la derecha). `construir_mapa_propietario()`
lo busca **por nombre de columna**, como ya hacían todos los lectores de esa hoja
(`_leer_resumen_bess` detecta la fila de encabezados), así que el corrimiento no rompió nada — se
verificó explícitamente armando un `Centrales.xlsx` con la columna insertada y volviendo a correr
`construir_mapa_barra`, `construir_dic_resumen_factor`, `..._capacidad` y `..._eficiencia`: los
cuatro dan lo mismo que antes. Un `Centrales.xlsx` viejo sin esa columna no revienta: la columna
`Propietario` queda vacía y se avisa en el log.

**Lo que NO se resolvió, y está anotado en "Pendientes abiertos":** `FD` y `FMA` (`Subastas!P:Q`)
quedan vacías por pedido del usuario —venían pegadas en `DB!Y`/`DB!V` y no existen en el Access—;
`Energía SSCC` se tomó de `CANTIDAD PONDERADA MW` y falta que el usuario confirme que no es
`CANTIDAD MW`; y el ajuste de `Hora_mes` por cambio de hora quedó como parámetro apagado, con un
aviso en el log si algún día del mes no trae 24 horas.

**Verificación.** No hay Access ni driver de Access en el entorno de esta sesión, así que se probó
en dos mitades: (a) la transformación, contra la **salida real** de `entradas_sscc.py` para marzo
2026 que entregó el usuario (`docs/subastas_2603_salida_entradas_sscc.xlsx`, 39.181 filas), que
tiene exactamente las columnas que devuelve la consulta a los Access — de ahí salen 2.904 filas
BESS/SAE y los pares `Concepto`/`Control`/`Sub_Baj` quedan consistentes en las cuatro
combinaciones presentes (`CSF(+)`, `CSF(-)`, `CPF(+)`, `CPF(-)`); y (b) la lectura de los Access,
con un caso sintético que reemplaza `leer_accdb` y verifica el solapamiento PO/PID hora por hora
y los cuatro filtros. Además, un caso end-to-end que crea la carpeta, copia desde una "red"
simulada (dos veces, para ver que la segunda no vuelve a copiar), corre `generar_consolidado` con
solo la sección `subastas` y revisa la hoja escrita. **Falta correrlo una vez contra Access
reales en Windows**: es el próximo paso natural.

---

## 2026-09-12 (2) — FMA desde sus tres salidas, carpeta "FD y FMA" y Subastas ordenada

Continuación de la entrada anterior, con un documento de trazabilidad nuevo del usuario
(`docs/Trazabilidad_FMA_a_DB_V_planilla3.md`) y tres pedidos.

**1) La hoja `Subastas` sale ordenada por `Hora_mes`** (`_ordenar_subastas_por_hora_mes()`), con
desempate por `Configuración` y `Concepto` para que dos corridas sobre los mismos datos den
exactamente el mismo archivo. Se aplica a los dos caminos (Access y respaldo planilla 3). El orden
es puramente de presentación: todo lo que consume `Subastas` lo hace por clave (`calcular_l`,
`construir_prorrata_sscc`, `construir_dic_umbrales_subastas`), nunca por posición de fila. Detalle:
se ordena por el valor **numérico** de `Hora_mes`, porque por el camino de la planilla 3 puede
llegar como texto y ahí "10" iría antes que "9".

**2) La carpeta `SSCC_Desempeño/` pasa a llamarse `FD y FMA/`** (`CARPETA_FD_FMA`), porque ahora
guarda las dos cosas: el `SSCC_Desempeño_*` del que sale la hoja `FD` y las tres salidas de FMA.
Un caso viejo que todavía tenga la carpeta con el nombre anterior sigue funcionando: si la nueva
no existe y la vieja sí, se usa la vieja (`resolver_rutas`), y el diagrama muestra el nombre de la
que realmente se está usando, no la constante.

**3) `Subastas!FMA` (la columna `Q`, antes `DB!V`) ya se calcula.** Módulo nuevo
`Script/Subastas/Fma.py` (lee y normaliza los tres archivos) + `calcular_fma_subastas()` en
`nucleo` (el cruce fila por fila, replicando la fórmula de `DB!V`):

| Concepto | De dónde sale | Clave |
|---|---|---|
| `CPF(+)` / `CPF(-)` | `fma_cpf_*` | central (traducida) + Año + Mes + Día + Hora |
| `CSF(+)` / `CSF(-)` | `fma_csf_*` | Año + Mes + Día + Hora |
| `CTF(+)` / `CTF(-)` | `fma_cft_*` | Configuración + Año + Mes + Día + Hora |
| otro | — | 0 |

Lo que no se encuentra queda en **0**, igual que el `IFERROR` exterior de la fórmula original.

**Tres trampas del documento que se respetaron al pie de la letra**, porque ninguna es deducible
mirando los archivos:

- **CPF**: se suman las 7 columnas de horas y, si la suma **llega a 0,98, pasa a ser 1** (no se
  redondea: se reemplaza). Los `Tiempo f<49.975 [%]` / `f>50.025 [%]` **ya vienen como fracción**
  (0,33 = 33 %), así que **no** se vuelven a dividir por 100.
- **CSF**: la hora del archivo va de **0 a 23** y la de las subastas de **1 a 24** → se le suma 1.
  Y se usan las columnas **`-m`/`+m`** (`FMA CSF-m [%]`, `FMA CSF+m [%]`), **no** las que se llaman
  igual sin la "m", que también existen en el archivo y darían otro número.
- **CTF**: varias activaciones en la misma hora **se suman**; y `CTF(-)` busca por `unidad` y, si no
  encuentra, por `Configuracion`, mientras que `CTF(+)` busca **solo en la primera** — eso es un bug
  de la fórmula original (el segundo bloque `CTF(+)` es inalcanzable por cómo están anidados los
  `SI`) que el documento pide **conservar** en la primera réplica. Quedó en la constante
  `CTF_MAS_BUSCA_EN_LAS_DOS_TABLAS = False` para poder comparar después.

**Dependencia auxiliar nueva:** la equivalencia `Configuración → central como la nombra fma_cpf`
(en la planilla vivía en `FMA_CPF!AD:AE`) se busca en un bloque titulado **`FMA CPF`** de la hoja
`Diccionario` de `Centrales.xlsx`, con el mismo mecanismo de bloques lado a lado que ya tenía esa
hoja (`construir_dic_bloque_diccionario()`, genérica). Si el bloque no está, se prueba con el
nombre tal cual y se avisa — hay que agregarlo (ver "Pendientes abiertos").

**Lo que quedó fuera:** el **Vector de Participación CSF**, que multiplica al FMA de las filas CSF.
Vive en `CSF_FD` de la planilla 3, o sea depende de la trazabilidad de FD, que todavía no existe.
Se usa 1 y queda anotado como el único supuesto que hoy puede dar un número distinto del de la
planilla. `Subastas!FD` (la columna `P`) sigue vacía por la misma razón.

**Verificación.** Archivos sintéticos armados con los esquemas exactos del documento: se comprobó
a mano el CPF con la suma por debajo del umbral (0,97 → se usa 0,97) y por arriba (0,99 → se usa
1), que el CSF corre la hora en uno y toma las columnas "m", y que el CTF suma las tres
activaciones de una misma hora repartiéndolas por signo. Después, un cruce de 10 filas de Subastas
que cubre las seis combinaciones de concepto más la traducción de nomenclatura CPF, la hora sin
dato (→ 0), el respaldo por `Configuracion` del `CTF(-)`, el `CTF(+)` que **no** cae a la segunda
tabla y un concepto raro (→ 0): las 10 dieron el valor esperado. Y un caso end-to-end que escribe
`Consolidado_entradas.xlsx` con la carpeta nueva, verifica el orden por `Hora_mes` y que el
programa siga aceptando la carpeta con el nombre viejo. **Falta compararlo contra un `DB!V` real**
(el documento propone justamente esa validación, §24).

---

## 2026-09-12 (3) — Boton "Traer FD"

El usuario pidio un boton que baje el FD, apuntando a que en `entradas_sscc.py` y su
`archivo_de_configuracion.yaml` esta como se llega a el.

**Precision importante:** el script **no copia** el FD de ningun lado. Lo que hay ahi son dos
piezas sueltas: el nombre del archivo
(`SSCC_Disponibilidad_CSF_<Mes>_20<AA>_<version_fd>.zip`, con `version_fd` = V1 para el Preliminar
y V2 para el Definitivo) y la raiz `ruta_fma_dco`. La ruta completa aparece en el comentario de la
rutina de FMA CPF:

    ...\02 Cálculo indicadores\2021\12. Diciembre\Indicadores Publicar\V1\01 Respuesta\...

de donde sale que la carpeta del periodo es
`<RAIZ>\<AAAA>\<MM>. <Mes>\Indicadores Publicar\<version>`. Eso es lo que se armo.

**`Script/Fd/Indicadores_DCO.py`** (paquete nuevo, mismo patron que los otros: no importa
`nucleo`, solo biblioteca estandar). `nucleo.traer_fd()` lo llama desde el boton **"Traer FD"**,
que cuelga de la fila de la carpeta `FD y FMA/`. Copia lo que encuentra a esa carpeta y, si lo
publicado es el `.zip`, lo descomprime ahi mismo: lo que la etapa FD lee despues es el Excel
`SSCC_Desempeño_*` que viene adentro, y asi queda justo donde `buscar_archivo_sscc_desempeno()` lo
busca.

**Cuatro decisiones que hubo que tomar** (ninguna venia dada, todas estan en el codigo comentadas
y son de una linea si se quiere cambiarlas):

- **Que version bajar.** La ventana no tiene selector Pre/Def, asi que se toma la **mas alta
  publicada** (V2 le gana a V1) y se dice en el log cual uso. La funcion acepta `version` por si
  despues se quiere elegir a mano.
- **Las carpetas de año y mes se buscan comparando por nombre normalizado**, no con una ruta
  literal: las escribe una persona todos los meses, y `"03. Marzo"` y `"3. Marzo"` tienen que ser
  la misma (se probo que lo son).
- **La busqueda del archivo es recursiva** dentro de la carpeta de version, porque el DCO cambia
  de subcarpeta de un mes a otro. Se filtra por prefijo **y** por el año en el nombre, para no
  llevarse un archivo de otro periodo que haya quedado suelto ahi (se probo con uno de febrero de
  2025 puesto a proposito en la misma carpeta: no se copia).
- **Al descomprimir se sacan solo los Excel**, sueltos en la carpeta (sin recrear el arbol del
  zip), y se ignoran las entradas con ruta absoluta o con `..`: un zip no tiene por que poder
  escribir fuera de la carpeta a la que se lo descomprime.

**Verificacion:** un arbol del DCO simulado con las dos versiones publicadas, un zip con un Excel
y un `.txt` adentro, y ruido de otro periodo en la misma carpeta. Se comprobo la resolucion de la
ruta, que elija V2, que se pueda forzar V1, que copie y descomprima, que el `.txt` NO salga, que
el archivo de otro mes NO se copie, que lo que queda lo encuentre despues
`buscar_archivo_sscc_desempeno()`, y los mensajes de error de los dos casos que de verdad van a
pasar (mes todavia sin publicar y servidor sin conectar). **Falta correrlo una vez contra el DCO
real**, que es donde se va a ver si el FD esta en esa carpeta de version o mas adentro con otro
nombre.

---

## 2026-09-12 (4) — Boton "Traer FMA"

Mismo pedido que el del FD, ahora para el FMA. Y **la misma precision, mas fuerte todavia**: en
`entradas_sscc.py` el FMA no se copia de ningun lado, se **construye**, y cada una de las tres sale
de un origen distinto:

| | De donde | Que se le hace |
|---|---|---|
| CPF | reportes diarios del DCO, en el MISMO arbol del que sale el FD (`<version>/01 Respuesta/01 Indices CPF/20AA.MM_Respuesta_CPF/Reporte diario <D>-<M>-<AAAA>/tabla_resumen_<D>_<M>_<AAAA>.xlsx`) | se lee cada hoja menos "Resumen" (una por central), se le pegan Año/Mes/Dia/Hora/Central y se agrupan las columnas de horas |
| CSF | los `csf_20AAMMDD.xlsx` diarios (los del `agc_face`) | concatenacion, sin transformacion |
| CTF | el `CTF_20AAMM.csv` | se le saca la zona horaria a `t0`/`tfin` |

**`Script/Fd/Indices_FMA.py`** (modulo nuevo, al lado de `Indicadores_DCO.py` porque comparten la
raiz de red y el destino). `nucleo.traer_fma()` lo llama desde el boton **"Traer FMA"**. Escribe
las tres con los nombres exactos de `entradas_sscc.py` (`fma_cpf_AAMM.xlsx`, `fma_csf_AAMM.xlsx`,
`fma_cft_AAMM.xlsx` -- "cft" incluido) en `FD y FMA/`, que son justo los que busca
`Script/Subastas/Fma.py` para armar `Subastas!FMA`. **Las dos mitades quedaron probadas juntas**:
lo que escribe este modulo lo lee el otro sin tocar nada, y el FMA CPF de una fila dio el 0,3201
esperado a mano.

**Los botones en la ventana.** Como cada fila del arbol lleva un solo boton, quedaron asi: "Traer
FD" en la fila del `SSCC_Desempeño_*` (que es lo que trae) y "Traer FMA" en la fila de
`fma_cpf_<AAMM>.xlsx`, con las otras dos filas de FMA diciendo en su detalle que las arma ese
mismo boton. Un boton por cosa que produce, sin repetir el mismo tres veces.

**Un error que aparecio en la prueba y vale la pena dejar escrito:** al sacarle la zona horaria al
CTF, la primera version uso `pd.to_datetime(..., utc=True).dt.tz_localize(None)`, que **convierte**
a UTC: `04:00-03:00` terminaba en las 07:00 y eso corria TODAS las horas del CTF (se vio como
horas 8 y 10 donde tenian que ser 5 y 7). Lo correcto es lo que hace el script original -dejar la
hora tal como esta escrita y solo sacar el ofset-, que es lo que ahora hace
`_sacar_zona_horaria()`, contemplando ademas el caso de ofsets mezclados en la misma columna (el
dia del cambio de hora, donde pandas ya no devuelve una columna tz-aware).

De paso, `buscar_archivos_fma()` ahora prefiere el `.xlsx` sobre el `.csv` cuando estan los dos
(el CTF se escribe en los dos formatos): antes ganaba el mas reciente, que es no determinista.

**Otras decisiones:** un dia sin reporte se saltea con aviso en vez de cortar todo (el original
revienta con la excepcion de pandas); si falta el origen de una de las tres, las otras dos se
arman igual y se avisa cual falto; y no se uso `chardet` para la codificacion del CSV de CTF -se
prueban en orden las que de verdad aparecen-, para no sumar una dependencia por eso.

**Verificacion:** un arbol del DCO simulado con reportes diarios de dos dias, cada uno con la hoja
"Resumen" (que se saltea) y dos centrales, con el encabezado en la fila 5 y las 29 columnas
reales; mas los insumos de CSF y CTF en la carpeta del caso. Se comprobo que salen las tres, que
"Resumen" no queda como central, que las horas van de 1 a 24, que los `"-"` quedan en 0, que el
CTF conserva la hora local, y que `cargar_tablas_fma()` lee las tres salidas y da el numero
esperado. **Falta correrlo contra el DCO real**, que es donde se va a ver si los reportes estan
donde dice el script y si el reporte sigue teniendo 29 columnas.

---

## 2026-09-12 (5) — Las rutas reales de los FMA y un boton "Generar" por cada uno

El usuario paso las tres rutas de entrada que faltaban y pidio que **cada FMA tenga su propio
boton "Generar"** en vez del unico "Traer FMA" de la entrada anterior.

**Las tres rutas** (confirmadas por el usuario, ya no inferidas):

| | Origen |
|---|---|
| CPF | `...\Indicadores Publicar\<V1/V2>\01 Respuesta\01 Indices CPF\20AA.MM_Respuesta_CPF\Reporte diario D-M-20AA\tabla_resumen_D_M_20AA.xlsx` (confirma lo que ya estaba) |
| CSF | `\\nas-cen1\D. Transferencias\SCADA\reporte_agc_face_NM10` |
| CTF | `...\Indicadores Publicar\<V1/V2>\01 Respuesta\06 Indices CTF\CTF_<AAAA><MM>.csv` |

La del CPF confirma la que ya se habia deducido del comentario del script. Las otras dos son
nuevas: el **CTF esta en el mismo arbol del DCO pero en otra rama** (`06 Indices CTF` en vez de
`01 Indices CPF`), y el **CSF no esta en el DCO en absoluto** -vive en otro servidor, en una
carpeta donde estan TODOS los meses juntos-.

**Lo que cambio en el codigo:**

- `SUBCARPETAS_CPF` y `SUBCARPETAS_CTF`: ahora se baja por los nombres reales en vez de buscar a
  ciegas. Se mantiene la busqueda recursiva **como respaldo** por si el DCO cambia el anidamiento:
  primero la ruta que dio el usuario, y si ahi no esta, se busca desde la carpeta de version.
- `traer_agc_face()`: copia a `FD y FMA/agcface/` **solo los archivos del mes** (en la carpeta de
  red estan todos), salteando los que ya estan al dia. El nombre de la subcarpeta es `agcface`,
  sin guion bajo, tal cual lo escribio el usuario -ojo, el script original usaba `agc_face`-.
- El nombre de los reportes del AGC en esa carpeta de red no lo vimos nunca, asi que se prueban
  dos criterios en orden: primero `csf_<AAAAMMDD>` (el del script original) y, si en todo el mes
  no aparece ninguno, cualquier Excel cuyo nombre **contenga** `<AAAAMMDD>`. El log dice con cual
  de los dos los encontro. **Es lo unico de esta entrada que puede necesitar un ajuste al correrlo
  contra la carpeta real.**
- `traer_fma()` paso a ser **`generar_fma(..., tipos=...)`**, que arma solo las pedidas. Los tres
  botones "Generar" del arbol mandan un tipo cada uno. Consecuencia util: **el CSF no necesita que
  el DCO tenga el mes publicado** (su origen es otro servidor), asi que su boton anda igual aunque
  todavia no haya indicadores; se probo explicitamente apuntando la raiz del DCO a una carpeta
  inexistente.
- El CTF ademas se **copia a la carpeta del caso** antes de usarlo, para que quede registrado con
  que archivo se armo la salida (igual que se hace con los Access de subastas y con el FD).

**Verificacion:** un arbol del DCO simulado con las **dos ramas reales** (`01 Indices CPF` y
`06 Indices CTF`) y una carpeta del SCADA con archivos de junio, julio y agosto mezclados. Se
genero **de a un tipo por vez**, como lo van a hacer los botones, y se comprobo: que el CPF
encuentre los reportes por la ruta nombrada, que a `agcface/` lleguen **solo los de julio** (no
junio ni agosto), que el CTF salga de `06 Indices CTF` conservando la hora local, que las tres
salidas las lea despues `cargar_tablas_fma()` dando el 0,3201 esperado, y que el CSF funcione con
el DCO caido. **Falta correrlo contra las carpetas reales**, sobre todo por el nombre de los
archivos del AGC.

---

## 2026-09-12 (6) — La version del DCO es la mas alta QUE TENGA EL ARCHIVO

Correccion del usuario sobre lo de la entrada anterior: no alcanza con quedarse con la carpeta de
version mas alta, porque **puede existir la carpeta y no estar adentro el documento que se busca**
(recien creada, a medio subir, o esa version no incluye esa entrega). Tiene que ser la mas alta
**que tenga disponible el archivo**.

`Indicadores_DCO.buscar_en_versiones(carpeta_publicacion, buscar, ...)` recorre las versiones **de
mayor a menor** y devuelve la primera en la que la funcion `buscar` encuentre algo, junto con la
lista de las que reviso (para poder decirlo en el error). Los tres usos pasan por ahi:

- **FD**: busca los `SSCC_Disponibilidad_CSF*` / `SSCC_Desempeño*` del año.
- **CPF** (`buscar_reportes_cpf`): no le alcanza con que exista la carpeta `*Respuesta_CPF*` —
  exige que tenga adentro **al menos un `tabla_resumen` del periodo**. Ese "al menos uno" es
  justo el caso que planteo el usuario.
- **CTF** (`buscar_ctf`): busca el `CTF_<AAAA><MM>.csv` bajo `01 Respuesta/06 Indices CTF`.

Consecuencia de diseño: **cada tipo elige su version por separado**. Puede pasar perfectamente que
el CPF salga de V1 y el CTF de V2, asi que `generar_fma()` ya no devuelve una carpeta de version
sino un **dict por tipo**, y el log de cierre dice de donde salio cada uno. El log tambien avisa
cuando bajo de version ("se usa V1: en V3, V2 no estaba el archivo"), que es informacion que el
usuario necesita ver: si esperaba la definitiva y salio la preliminar, tiene que enterarse.

**Verificacion:** un arbol del DCO con tres versiones armado a proposito para esto — `V3` existe
pero esta **vacia**, `V2` tiene la carpeta `Respuesta_CPF` **sin reportes adentro** y si tiene el
CTF, y `V1` tiene los reportes de CPF y el FD. Resultado: el FD baja a V1, el CPF baja a V1
(salteando la carpeta vacia de V2, que es el caso fino), el CTF se queda en V2, y forzar una
version que no lo tiene da un error que nombra las versiones revisadas.

---

## 2026-09-12 (7) — El nombre de los reportes del AGC, confirmado

El usuario confirmo el formato: **`csf_20260301`**, o sea `csf_<AAAA><MM><DD>` — exactamente el
que espera el script original.

Con eso **se saco el segundo criterio** que se habia dejado en la entrada anterior ("cualquier
Excel cuyo nombre contenga la fecha"), que era una suposicion mientras no supieramos el nombre.
Queda solo el criterio confirmado, que ademas es mas seguro: un archivo que apenas contenga la
fecha en el medio del nombre ya no se cuela. Tambien se saco `PLANTILLA_FECHA_DIARIA`, que solo
existia para ese respaldo, y el mensaje de error ahora dice que nombre se esta buscando
(`'csf_20260301' y siguientes`), que es lo util si algun mes no aparece nada.

**Verificacion:** una carpeta del SCADA con cinco archivos a proposito — dos de marzo con el
nombre real, uno de abril, uno que **contiene** la fecha pero no empieza con `csf_`
(`reporte_20260303.xlsx`) y uno con sufijo despues de la fecha (`csf_20260304_v2.xlsx`). Se copian
los tres de marzo que empiezan con `csf_` (el del sufijo incluido), no se copia el de abril ni el
que solo contenia la fecha, y el `fma_csf` resultante tiene las 72 filas de los tres dias.

---

## 2026-09-12 (8) — La subcarpeta del FD dentro del arbol del DCO

El usuario paso la ruta que faltaba: los **factores de desempeño** (el FD) estan en

    <version>\04 Desempeño para transferencias

Hasta ahora el boton "Traer FD" los buscaba **recursivamente** desde la carpeta de version, que
funcionaba pero recorria todo el arbol publicado del mes -que tiene adentro los reportes diarios
de CPF, o sea cientos de carpetas- para encontrar un archivo que esta en una sola. Ahora se busca
primero en la subcarpeta nombrada (`SUBCARPETAS_FD`) y la busqueda recursiva queda solo como
respaldo, mismo criterio que ya se habia usado para el CPF y el CTF.

Se agrego algo util para el unico caso que puede fallar: si en ninguna version aparece un archivo
con los nombres conocidos (`SSCC_Disponibilidad_CSF*` / `SSCC_Desempeño*`), el error **lista lo que
si hay** en esa carpeta, version por version. Si algun mes le cambian el nombre al archivo, se ve
de una y se corrige el patron sin tener que ir a mirar el servidor.

`bajar_por_subcarpetas()` se mudo de `Indices_FMA` a `Indicadores_DCO` (lo usan los dos modulos, y
el segundo es el de mas abajo); en `Indices_FMA` quedo reexportado para no tocar sus llamados.

**Verificacion:** un arbol con `V2` que tiene la carpeta `04 Desempeño para transferencias` pero
**vacia** y `V1` que si tiene el zip. Baja a V1 (la version mas alta que TIENE el archivo, como se
corrigio en la entrada anterior), lo copia, lo descomprime y el Excel queda donde la etapa FD lo
busca. Y renombrando el archivo a algo que no matchea, el error muestra
`V1: Factores_desempeno_agosto.zip`, que es exactamente el dato que haria falta para corregirlo.

---

## 2026-09-12 (9) — FD y Vector de Participación CSF: se cierran los dos pendientes

Con el documento de trazabilidad de FD (`docs/Trazabilidad_FD_a_DB_Y_planilla3.md`) se cerraron
**los dos** pendientes que quedaban de la hoja `Subastas`, porque los dos salen del mismo archivo,
el `SSCC_Desempeño_*` que ya trae el boton "Traer FD":

- **`Subastas!FD`** (columna `P`, antes `DB!Y`), y
- el **Vector de Participación CSF** (antes `DB!AC`), que multiplica al FMA de las filas CSF y era
  el unico supuesto que podia dar un numero distinto del de la planilla sin avisar.

**`Script/Fd/Desempeno_Horario.py`** (modulo nuevo) normaliza las tres hojas horarias:

| Hoja | Columnas | El FD es |
|---|---|---|
| `CPF Horario` | B:J | `I` (Fd_CPF) |
| `CSF Horario` | B:H | `H` (Fd_CSF) |
| `CTF Horario` | B:I | `I` (Fd_CTF) |

Encabezados en la fila 11 y datos desde la 12, que es **el mismo criterio que ya usaba
`nucleo.construir_fd()`** para armar la hoja `FD` del consolidado desde las dos primeras (la
tercera, `CTF Horario`, no la leia nadie hasta ahora). Se dejaron las dos lecturas separadas
a proposito: arman cosas distintas y la de `construir_fd` ademas filtra por BESS/SAE.

**Tres detalles del documento que no son deducibles mirando el archivo:**

- la `Hora` de estas hojas va de **0 a 23**, asi que `Hora_Mes = (dia-1)*24 + hora + 1`. Ese "+1"
  es lo que la deja en la misma escala 1..24 por dia que usa `Subastas` — sin el, el cruce entero
  se corre una hora;
- **CPF prueba una segunda nomenclatura**: si no encuentra la unidad, intercambia el sufijo
  `TG` <-> `TV` y busca de nuevo. CSF y CTF buscan una sola vez (asi es la formula original, y asi
  quedo: se probo que CSF NO cae a la alternativa);
- el **Indicador de Participación CSF** es 0 solo si la unidad figura como `"No Participó"` **y**
  su alternativa TG/TV tampoco participo. No alcanza con mirar la propia fila.

Y una que importa aunque parezca menor (§25 del documento): **el FD no se recalcula** a partir de
las respuestas, se toma tal cual viene. Puede venir `Respuesta = "No Participó"` con `FD = 1`, y
asi tiene que quedar.

**La nomenclatura** (`Configuración` -> unidad del archivo de desempeño) NO necesito un bloque
nuevo en el `Diccionario`: es el bloque **`FD`** que esa hoja ya tenia, el mismo que usa
`Calculo E Costos!AM:AR`. Se reusa `construir_dic_bloque_diccionario()` con respaldo a
`construir_dic_mapeo_diccionario()` (columnas A:B).

**Una diferencia deliberada con la planilla:** cuando la busqueda no encuentra nada, la formula
original escribe el texto `"ERRORCPF"` / `"ERRORCSF"` / `"ERRORCTF"` en la celda. Aca la columna
queda **vacia** y el log dice cuantas filas fueron, por familia: meter texto en una columna
numerica rompe cualquier cuenta posterior. Para el Vector de Participación, si la fila no aparece
se usa **1** (`VECTOR_PARTICIPACION_CSF_SIN_DATO`) y se avisa: la formula original ahi daria #N/A,
y dejar el FMA en su valor base es mas prudente que ponerlo en 0, que seria no pagar.

**Verificación:** un `SSCC_Desempeño_*` sintetico con las tres hojas en su layout real (columna A
vacia, encabezados en la fila 11), una unidad que solo existe como `TV` para probar el salto
`TG`<->`TV`, y una central con `"No Participó"`. Se comprobo: las tres tablas toman la columna de
FD que corresponde, `Hora_Mes` da 1 para el dia 1 hora 0 y 25 para el dia 2 hora 0, el CPF
encuentra por la alternativa y el CSF **no**, la columna FD sale por familia y la hora inexistente
queda vacia, y el FMA de la central que no participo queda en 0 mientras la que si participo
conserva su valor base. Ademas, un caso end-to-end que escribe `Consolidado_entradas.xlsx` con las
seis filas completas: `FD` con dato en todas y `FMA` = 0,3201 / 0,5 / 0 segun corresponda.
**Falta compararlo contra un `DB!Y` y un `DB!AC` reales** (§30 del documento).

---

## Sesión 2026-09-14 — alertas para cruces que antes fallaban en silencio

- Se agregó un diagnóstico común que agrupa y nombra centrales sin barra,
  Pmax, capacidad o eficiencia en `Resumen BESS`, incluyendo también valores
  existentes pero vacíos. El aviso aclara que los resultados dependientes
  pueden terminar vacíos o en cero.
- `Calculo E Costos` ahora marca explícitamente como `[AVISO]` las filas sin
  CMg y reporta centrales que no aparecen en `Diccionario!A:B`, además de las
  claves CPF/CSF homologadas que no existen en FD y que se completan con cero.
- `Calculo RE545` ahora alerta por barra, CMg, CMg Promedio, capacidad y
  eficiencia faltantes. Se conserva el resultado compatible con la planilla;
  cambia la visibilidad del problema, no la fórmula.
- Se agregaron pruebas unitarias para el agrupamiento de alertas y para los
  dos caminos de FD faltante (vacío por diccionario y cero por clave FD).

---

## Sesión 2026-09-14 (segunda pasada) — los cruces que quedaron sin aviso

Revisión del commit anterior (`Alertar cruces faltantes antes de completar con
cero`). Lo que hacía está bien y los dos tests pasan; lo que sigue son los
huecos que quedaron y cómo se cerraron.

- **`tests/` no era importable**: `python -m unittest discover` desde la raíz
  devolvía `Ran 0 tests ... OK` — un falso verde, justo en el commit que iba
  sobre errores silenciosos. Se agregó `tests/__init__.py`.
- **Pmax faltante solo se avisaba en `Calculo E Costos`**. `Calculo RE545` usa
  el mismo `dic_factor` en `calcular_componentes_re545()` (BN): sin Pmax, BN
  queda vacío y `BQ = BS + BN` lo pasa a 0 con `fillna(0.0)`. Ahora avisa, y el
  aviso va al principio de la función para que salga aunque algo más adelante
  falle.
- **Pmax presente pero en 0** (o no numérico) deja `AE`/`AF` vacías en
  `calcular_ae_af()` y no lo cubría `_avisar_claves_sin_mapeo()`, que solo mira
  ausente/vacío. Aviso propio con las centrales afectadas.
- **Prorrata SSCC** (`calcular_prorratas()`, AG/AH): sin match, 0. Se distingue
  el caso legítimo (una hora suelta sin SSCC → aviso por filas) del que no lo es
  (una central sin **ninguna** hora en la prorrata → la homologación contra
  `Subastas!Configuración` no está cruzando y se cae todo el prorrateo de SSCC
  de esa central).
- **Reservas de RE545** (`calcular_reservas_re545()`): un SUMIFS sin
  coincidencias da 0 y eso es correcto hora a hora, pero una central que no
  aparece en ninguna clave de `Subastas` se queda con las 18 reservas y `AU` en
  0 — se paga como si no hubiera tenido reservas. Aviso por central.
- **Rendimiento**: `_avisar_claves_sin_mapeo()` llamaba a `normalizar()`
  (unicodedata + regex) una vez por fila, o sea decenas de miles de veces por
  aviso para unas pocas centrales distintas. Ahora deduplica antes de
  normalizar.
- **README**: decía que se avisa por "central sin propietario"; ese aviso no
  existe. Se corrigió la lista y se documentó cómo correr las pruebas.

Queda a la vista, sin tocar (son agregaciones internas, no homologaciones):
los `dic_bc` / `dic_bf` / `dic_bg` de `calcular_componentes_re545()` usan
`.get(clave, 0.0)`; si una central+ventana faltara ahí sería una inconsistencia
interna del propio cálculo, no un cruce contra un archivo de entrada.

---

## Sesión 2026-09-14 (tercera pasada) — `nucleo.py` pasa a ser un paquete

`Script/__init__.py` venía diciendo hace meses que la idea era "ir sacando de
`nucleo.py` un módulo por etapa, como ya se hizo con `Cmg/`". Se hizo: el
archivo de 8.360 líneas y 193 símbolos es ahora `Script/nucleo/`, 27 módulos.

**Qué se movió.** Nada de lógica. El corte fue por bloques de líneas, así que
cada función viajó con sus comentarios de sección y de encabezado intactos. Se
verificó comparando el AST de los 193 símbolos antes y después: los únicos seis
que cambian son los que se tocaron a propósito (ver más abajo).

**La fachada.** `nucleo/__init__.py` re-exporta todo con un `__all__` de 205
nombres — los de guion bajo incluidos, porque las pruebas los usan. Ni
`Balance_BESS.py` (que usa 20 nombres de `nucleo`) ni las pruebas cambiaron una
línea.

**Sin ciclos.** Los imports van en una sola dirección: `parametros`/`utiles`
hacia las etapas, y las etapas hacia `proceso`. Para lograrlo hubo que sacar
dos cosas de donde estaban:

- `_texto_seguro`, `_tiene_valor`, `_es_numero`, `_valor_clave`,
  `_entero_a_texto`, `_normaliza_valor_vba` y `_columna_clave_vba` vivían
  dentro de la sección "Ofertas SSCC" y de la etapa 2 de E Costos, pero las
  usan cinco o seis etapas: ahora están en `utiles.py`.
- `calcular_l`, `calcular_m` y `calcular_n_o` son las MISMAS para las dos hojas
  de cálculo. Estaban en la etapa 2 de E Costos, así que RE545 tenía que
  importar de E Costos para calcular sus propias columnas — una dependencia
  falsa. Ahora son `columnas_compartidas.py`.
- `SECCIONES_CONSOLIDADO`/`SECCIONES_PAGOS` se movieron a `estructura.py`, que
  es quien las recorre (`proceso.py` las usaba solo de paso).
- `externos.py` concentra el `try/except` que importa los paquetes hermanos
  (`Cmg/`, `Fd/`, `Medidas/`, `Subastas/`): estaba escrito una vez y ahora lo
  usan siete módulos sin repetirlo.

**Redundancia eliminada.** `construir_mapa_barra()` y
`construir_dic_resumen_factor()` repetían, cada una, el buscador de columna y
el bucle `iterrows()` que `_mapa_resumen_bess_por_nombre()` ya hacía para
capacidad y eficiencia. Los cuatro diccionarios de "Resumen BESS" pasan ahora
por el mismo helper, con `_columna_resumen_bess()` / `_exigir_columna()` para
la búsqueda y un conversor por parámetro (la barra es texto, el resto números).
El error de columna faltante ahora **nombra cuál falta** en vez de listar las
tres juntas. `iterrows()` se reemplazó por `zip()` de las dos columnas.
Equivalencia comprobada contra la implementación vieja (misma salida en el caso
normal y con las columnas en orden invertido) y fijada en
`tests/test_diccionarios_resumen.py`.

**Limpieza menor** que salió de pasar pyflakes por todo: dos f-strings sin
placeholders en `estructura.py`, la variable muerta `descartadas` en
`extraer_soc()`, y un `import re` sin uso en `Script/Fd/Indices_FMA.py`.
pyflakes queda limpio sobre todo el repo.

**Cómo verificar que esto no rompió nada** (por si hay que repetirlo):

    python -m py_compile Balance_BESS.py Script/nucleo/*.py Script/*/*.py
    python -m unittest discover     # 14 pruebas

y, dentro de Python, que `dir(nucleo)` siga teniendo los mismos nombres que
antes y que los 20 `nucleo.<algo>` de `Balance_BESS.py` resuelvan.

**Pendiente, a la vista:** `proceso.py` importa de 17 módulos. Es lo esperable
en un orquestador, pero si crece más conviene partirlo en
`proceso_consolidado.py` / `proceso_pagos.py`, que son dos caminos
independientes que hoy solo comparten el archivo.

---

## Sesión 2026-09-14 (cuarta pasada) — las claves de las APIs salen del código

Pedido del usuario: las dos `user_key` (PRMTE y Generación real) van a
`config.json`, en una sección que **no depende del usuario** — el valor es el
mismo para todo el equipo.

**Lo que estaba mal antes** (y que este cambio corrige de paso): había **una
sola** constante `USER_KEY` en `Script/Medidas/comun.py`. La sesión que la
unificó lo hizo creyendo que las dos APIs pedían la misma clave ("antes estaba
repetida en dos archivos, y con valores distintos" — los valores distintos eran
lo correcto, no el error). El usuario confirmó ahora que **son distintas**:
`prmte` es la de `medidas.coordinador.cl` y `generacion_real` la de
`operacion.coordinador.cl`.

**Cómo quedó.** `Script/config.py` es ahora el único módulo que toca
`config.json`. El archivo tiene dos clases de sección:

- `"<hostname>_<usuario>"` — carpeta base y AAMM recordados, por PC/usuario.
- `"claves_api"` — compartida, con las dos claves. Nombre reservado.

`Balance_BESS.py` dejó de leer y escribir el JSON a mano: sus `leer_config()` /
`guardar_config()` ahora delegan en `config.seccion()` /
`config.actualizar_seccion()`, que mezclan sin pisar el resto del archivo. O
sea: guardar la carpeta base **no puede** borrar las claves (hay una prueba que
lo fija).

`comun.leer_clave_api(cual)` traduce `ErrorConfig` a `ErrorMedidas`, que es lo
que el resto de Medidas ya sabe convertir en `ErrorEntrada` para la ventana.
La clave se lee **en el momento de usarla**, no al importar: así se puede
completar el `config.json` con el programa ya abierto.

**El error cuando falta** trae la ruta absoluta del archivo y el JSON exacto
para pegar, y aclara cuál clave es de cuál API. Está fijado en las pruebas
(`tests/test_config_claves.py`, 7 casos: cada clave por separado, archivo
ausente, JSON roto, el `PEGAR_AQUI_LA_CLAVE` del ejemplo sin reemplazar, una
puesta y la otra no, y el que las secciones no se pisan entre sí).

**`config.ejemplo.json`** se versiona (no tiene ninguna clave adentro) y es lo
que hay que copiar como `config.json`. `config.json` sigue en `.gitignore`.

**Consecuencia buscada:** el repositorio ya no tiene ni un lugar donde poner
credenciales. La nota de METODOLOGIA §5 que decía "queda versionada, así que el
repositorio no puede volverse público sin rotarla antes" quedó sin efecto. Se
revisó el historial completo (`git log --all -p` buscando cualquier
`user_key`/`api_key` con un literal de 6+ caracteres): **nunca se commiteó una
clave real** — `USER_KEY` siempre estuvo en `""`. Así que no hay nada que rotar
por este motivo. Si en algún momento se hubiera commiteado una, sacarla del
código no la saca del historial y habría que rotarla igual.

---

## Sesión 2026-09-14 (quinta pasada) — el registro de alertas, la conciliación y el manifiesto

El usuario trajo un catálogo de controles para el traspaso (ahora en
`docs/Alertas_y_Controles_Traspaso_Python_BESS.md`). Se revisó contra el código
y se implementó la parte de la Fase 1 con mejor relación costo/beneficio.

**Lo que el catálogo detectó y estaba mal:**

- **`Pagos_BESS.xlsx` no tenía hoja de log.** Todos los `[AVISO]` de la etapa de
  cálculo —FD, prorrata, reservas, Pmax, CMg: los que tocan plata— salían solo a
  la caja de texto de la ventana, que no se guarda en ningún lado.
- **`_avisar_claves_sin_mapeo()` mostraba 15 faltantes y perdía el resto**, que
  es literalmente lo que `FD-005` del catálogo prohíbe.
- **No existía la conservación de energía** (`TRA-001…009`), el control más
  barato y más potente que hay: el reparto por `Ventana_No_Completa` es
  justamente donde una fila se puede perder.
- **No había estado de corrida**: todo era "aviso" plano, sin id ni severidad.

**Lo que se hizo:**

- `nucleo/alertas.py` — `Alerta` (id, severidad, etapa, central, clave, acción,
  origen del control) y `Registro`, que junta las alertas de una corrida y
  calcula su estado con las reglas de la sección 21 del catálogo (`APROBADA`
  solo sin CRÍTICAS ni ALTAS). El truco para no tocar 10 firmas: **un `Registro`
  ES un `registrar`** — se llama como `print` y además sabe guardar. `anotar()`
  funciona con los dos, así que las pruebas y los scripts sueltos siguen andando
  sin armar una corrida.
- `nucleo/conciliacion.py` — `Medidores = E Costos + RE545`, con tolerancia
  **escrita** (relativa `1e-9`, piso absoluto `1e-6`) y las dos guardadas en el
  libro junto con la diferencia observada. Detecta fila perdida (`TRA-007`),
  energía en la hoja equivocada (`TRA-005`/`TRA-006`) y filas de más o de menos
  (`TRA-001`). Ojo: **concilia contra `df_re545_base`**, no contra `df_re545`,
  que ya pasó por `renombrar_calculo_re545()` y no tiene `Energia_Positiva`.
- `nucleo/manifiesto.py` — nombre, ruta, tamaño, fecha y `sha256` de cada archivo
  que alimentó la corrida. Sin esto, `APROBADA` no es reproducible.
- Hojas **`Alertas`** y **`Ejecucion`** en `Pagos_BESS.xlsx`. La regla de volumen
  quedó definida: detalle completo al archivo **sin tope**, resumen con 15
  ejemplos a la pantalla.
- Ids en todos los cruces que ya avisaban: `MAE-001..005` (barra, Pmax, Pmax=0,
  capacidad, eficiencia), `CMG-004`, `FD-005`, `DIC-001`, `PRO-001/002`,
  `SUB-011`, y `PAG-001` (hoja del libro de pagos que quedó vacía).
- `PER-001` en la etapa de pagos: si la hoja `Medidores` del consolidado trae más
  de un mes, el libro quedó mezclado entre corridas y es CRÍTICA.

**Una consecuencia que conviene entender antes de que asuste:** una corrida
parcial (solo `Calculo RE545`) sobre un libro que no existía deja la otra hoja
vacía → `PAG-001` → `NO APROBADA`. Es correcto: el libro tiene una hoja de pagos
vacía. Por eso `Ejecucion` escribe siempre `hojas_recalculadas`.

**Pendiente, anotado en las secciones 25-27 del catálogo:** unificar el `Log` del
consolidado con estas hojas; escribir un libro de diagnóstico cuando la corrida
muere por `ErrorEntrada` (hoy `FALLIDA` no llega a ningún archivo); el sello de
período por hoja (`CTX-001`); la regla de la hora repetida del cambio de horario
(`DST-001`); los controles de unidades (`UNI-001`) y de supuestos hardcodeados
(`SUP-001`).

33 pruebas, incluida una que corre `generar_pagos_bess()` de punta a punta sobre
un caso sintético y verifica que el libro salga con las dos hojas nuevas.


## Sesión 2026-09-14 (sexta pasada) — PRORRATA_RETIROS y Resumen

Se cerró la asignación económica por empresa. El caso ahora reconoce la carpeta
`Prorrata retiros/` y el archivo `Prorrata_Retiros_AAMM_pre/def.xlsx`; usa
exclusivamente `Prorrata 15min`, sin reconstruir la prorrata desde las hojas
mensual u horaria.

Se agregó `nucleo/prorrata_retiros.py`, que valida tipos, vacíos, negativos,
duplicados y suma uno por cuarto (`RET-001`, `RET-004..007`), cruza la
compensación de `Calculo E Costos` y `Calculo RE545`, calcula `Pago =
Compensación × Prorrata` y verifica conservación por cuarto y mensual
(`RET-002/003/008/009`). `Pagos_BESS.xlsx` recibe las hojas
`PRORRATA_RETIROS` (detalle más consolidado mensual) y `Resumen` (`RECIBE`,
`PAGA`, `NETO`). Las cuatro hojas de negocio conservan el mismo comportamiento
de actualización independiente que ya tenían las dos hojas de cálculo.

La ventana muestra **Traer prorrata** en la fila de `PRORRATA_RETIROS` y
**Asignar pagos** en `Resumen`. Se añadieron cuatro pruebas unitarias; la suite
completa quedó en 37 pruebas.

---

## 2026-09-14 — Ocho pedidos del usuario sobre la corrida real

El usuario corrió el programa de punta a punta por primera vez sobre un
caso real y mandó la lista de lo que anduvo y lo que no. Esta sesión
resuelve los siete puntos que pedían un cambio de código.

### 1. La planilla 3 se fue del programa

`3_REMUNERACIÓN_SUBASTAS_E_ID_*` era el respaldo de la hoja `Subastas`
para los casos que todavía no tenían los Access copiados. El usuario
confirmó que **ya no se usa**. Se sacó la fila del diagrama de la ventana,
el camino de respaldo de `generar_consolidado()` (ahora, sin Access del
período, se corta con un error que dice qué botón apretar) y todo lo que
colgaba de eso: `construir_subastas()`, `buscar_archivo_subastas()`,
`PATRON_NOMBRE_SUBASTAS`, `HOJA_SUBASTAS_ORIGEN`.

### 2. FMA CPF: 6 minutos que eran el mismo árbol recorrido 60 veces

Reporte del usuario: *"se demora más de 6 minutos en correr, lo tuve que
parar. Debe tener algún loop"*. No era un loop infinito.
`buscar_tabla_resumen()` hacía un `rglob("*")` COMPLETO de la carpeta de
respuesta del DCO **cada vez que se la llamaba**, y se la llama una vez
por día del mes: primero desde `buscar_reportes_cpf()` (hasta 31 veces
por cada versión del DCO que se revise) y después otras 31 desde
`construir_fma_cpf()`. Sobre una unidad de red, 60-120 recorridas
completas del árbol.

Ahora el árbol se recorre **una vez por carpeta** (`_indice_excels()`,
`os.walk` en vez de `rglob` para no pagar un `stat` por entrada) y las 62
búsquedas se resuelven contra esa lista en memoria. El caché se vacía al
empezar cada `generar_fma()`, para que dos apretadas seguidas del botón no
se pisen si el usuario copió archivos en el medio.
`buscar_carpeta_respuesta_cpf()` también pasó a `os.walk`, con corte
apenas encuentra la carpeta.

### 3. El `Diccionario` nuevo (y las diferencias de FMA CPF y CSF)

El usuario adjuntó `Centrales_Propuesta_de_mejora.xlsx` con un formato
nuevo para la hoja `Diccionario`: **una sola tabla** con encabezados
`Balance_BESS | FD | Subastas | Ofertas | FMA_CPF`, en vez de las varias
tablas sueltas puestas lado a lado y separadas por columnas vacías. Su
diagnóstico era correcto: el formato viejo no tenía **dónde** poner la
nomenclatura de FMA CPF (el bloque "FMA CPF" nunca existió), así que
`construir_dic_bloque_diccionario()` devolvía `{}`, el FMA de las filas
CPF se buscaba con la `Configuración` tal cual y no encontraba nada.

Se aceptan **los dos formatos** (`encabezado_diccionario()` detecta cuál
es; `filas_diccionario()` y `mapa_diccionario()` dan acceso por nombre de
columna). Con el nuevo, además, las homologaciones de FD y de FMA CPF se
buscan por el nombre de **Subastas** y no por el canónico — que es la
segunda mitad del problema: la `Configuración` que trae el Access para
Tocopilla es `BAT_TOCOPILLA`, y contra un diccionario con clave
`SAE-TOCOPILLA` nunca iba a cruzar. Medido contra el
`Consolidado_entradas.xlsx` real que mandó el usuario: la nomenclatura de
FMA CPF pasó de homologar **0 de 9** `Configuración` a **7 de 9** (las dos
que faltan son las que la propuesta dejó con la celda `FMA_CPF` vacía).

Lo que este cambio **NO** explica es el reporte de CSF: en ese mismo
archivo las 9 `Configuración` ya cruzaban 9/9 contra el bloque `FD` del
`Diccionario` viejo, así que la homologación de la unidad no era el
problema ahí. Queda como pendiente, con el Vector de Participación CSF
como sospechoso.

Las cuatro lecturas de la hoja (`construir_homologacion`,
`construir_dic_mapeo_diccionario`, `_mapas_homologacion_fge`,
`construir_dic_bloque_diccionario`) se verificaron una por una contra
`docs/Centrales_real.xlsx` (formato viejo) y contra la propuesta (formato
nuevo): dan lo mismo para las centrales que ya funcionaban.

**Dos cosas de la propuesta quedaron como pendiente**, no se tocaron
porque son datos del usuario: el `FMA_CPF` de Andes 3 y Andes 4 está
cruzado, y Nuevo Quillagua 2 y Víctor Jara no tienen `FMA_CPF`.

### 4. `Subastas` vuelve al orden del origen

*"Al final que no esté ordenado por hora_mes, mala mía yo lo pedí pero
no."* `_ordenar_subastas_por_hora_mes()` quedó como el único lugar donde
se decide el orden de esa hoja (hoy, un `reset_index` y nada más), en vez
de borrar los llamados: ninguna columna de más abajo depende del orden de
las filas (todas cruzan por clave), así que si mañana hay que volver a
ordenar se cambia una función y no cinco llamados.

### 5. `Medidores` ya no depende de Ofertas SSCC

*"Para construir Medidas se leen las ofertas, me gustaría sacar lo de
ofertas de esa hoja y dejarlas en la hoja de ofertas."* Hecho, y la
dependencia quedó **dada vuelta**: ahora `Ofertas SSCC` necesita
`Medidores` (de ahí salen las centrales y las ventanas) y no al revés.

- La hoja `Medidores` pasó de `A:U` a `A:Q + U`: R
  (`Oferta_Completa_Dia`), S (`Indicador_Ventana_Oferta`) y T
  (`Ventana_No_Completa`) ya no se escriben.
- Las dos tablas de las que se derivan esas tres columnas ya estaban en la
  hoja `Ofertas SSCC`; no hizo falta agregar ninguna.
  `completar_ofertas_en_medidores()` las reconstruye en memoria cuando
  hacen falta, que es en `generar_pagos_bess()`: **T es la que reparte
  cada fila entre `Calculo E Costos` y `Calculo RE545`**, así que sin ella
  no hay cálculo posible.
- `leer_ofertas_sscc_consolidado()` lee las dos tablas ubicándolas por su
  título, no por una posición fija de columna.
- Consecuencia práctica: el botón **Actualizar** de `Medidores` ya no abre
  el archivo `*OfertasSSCC*` ni lo exige.

### 6. `Calculo E Costos` toma el FD del consolidado

*"Ecostos: se leen los sscc_desempeño, ¿no debería apuntar al consolidado
de entradas?"* Sí. `AM:AR` se armaba releyendo el `SSCC_Desempeño_*`, lo
que significaba que la hoja `FD` del consolidado y el FD que usaba E
Costos podían no ser el mismo dato (bastaba con dejar un archivo más
nuevo en la carpeta después de generar el consolidado). Ahora sale de la
hoja ya generada (`leer_fd_consolidado()`), igual que `Medidores` y
`Subastas`: el consolidado es la única foto de las entradas.

### 7. El `KeyError: 'Energia_Positiva'` y el archivo que no se guardaba

Dos bugs en el mismo reporte.

**El KeyError.** `completar_calculo_e_costos_grupos()` devolvía la hoja
**ya renombrada** a los nombres del Excel, así que cuando
`conciliar_energia()` pedía `Energia_Positiva` esa columna se llamaba
"Descarga kWh". Y como `NOMBRES_CALCULO_E_COSTOS` tiene nombres repetidos
a propósito (AG:AL y AM:AR comparten los seis CPF/CSF/CTF, y "Total" es U
y AX), pandas ni siquiera avisaba "columna renombrada": tiraba un
`KeyError` pelado desde `_get_loc_duplicates`. RE545 ya tenía resuelto
esto (`renombrar_calculo_re545()` aparte, y `df_re545_base` con los
nombres internos); E Costos no. Se le dio la misma forma
(`renombrar_calculo_e_costos()` + `df_ecostos_base`), y `_energia_total()`
ahora explica el problema en vez de tirar el `KeyError` a secas.

**El archivo que no se guardaba.** *"Además no se guarda nada de lo
progresado como para revisar."* La conciliación corre justo antes de
escribir, así que cualquier excepción ahí se llevaba puesto todo lo
calculado. Ahora va dentro de un `try`: si revienta, queda una alerta
`TRA-010` (crítica) y el libro **se escribe igual**. Lo único que se
pierde es la conciliación.

### Verificación

- `python -m py_compile` sobre todo y `python -m unittest discover`: **49
  pruebas** (eran 37), todas verdes, sin avisos de `pyflakes`.
- La prueba de punta a punta `PagosBessEndToEndTest` **dejó de saltearse**:
  antes el caso sintético no alcanzaba y el test se auto-skipeaba (un
  verde que no probaba nada). Ahora corre de verdad y se le agregó un
  segundo caso con las **dos** hojas juntas, que es exactamente el camino
  donde se rompía la conciliación — sin el arreglo, ese test falla.
- `tests/test_hojas_del_consolidado.py` (nuevo): los dos formatos del
  `Diccionario` lado a lado, el ida y vuelta de las hojas `FD` y
  `Ofertas SSCC`, R/S/T reconstruidas, y el mensaje de error de la
  conciliación.
- **Contra el `Consolidado_entradas.xlsx` real que mandó el usuario**:
  `leer_fd_consolidado()` devuelve los dos bloques (6.696 filas cada uno) y
  `leer_ofertas_sscc_consolidado()` las dos tablas (310 y 288 filas). Y lo
  más importante: `completar_ofertas_en_medidores()` reproduce
  **exactamente** las columnas R, S y T que ese archivo ya tiene escritas
  en `Medidores` — cero filas distintas en las tres.
- El índice de carpetas del FMA CPF se probó con un árbol sintético:
  31 búsquedas, **1** `os.walk`.

Lo que NO se probó: nada de esto se corrió todavía contra el caso real del
usuario. En particular, el número real de la mejora del FMA CPF y si con
el `Diccionario` nuevo desaparecen del todo las diferencias de CSF.

---

## 2026-09-14 (2) — La unidad de la generación real y las 2.232 alertas FD-005

Segunda tanda de la misma corrida real. Dos correcciones, y las dos
resultaron ser otra cosa de lo que parecían a primera vista.

### La generación que viene de la API de operación real estaba ×1000

Reporte del usuario: *"la generación que viene de Generación real está por
1000, o sea en MWh, y creo que hay que corregir el factor. En canal ahora
le puse MWh pero no he probado si afecta"*.

Tenía razón, y **el `Canal` no afectaba nada**: esa columna se leía sólo
para descartarla (la API de operación real no expone canales). Confirmado
con su `Consolidado_entradas.xlsx` real, comparando `Medidores!Gen_Unidad`
por central:

| central | Pmax | máximo por cuarto de hora | camino |
|---|---|---|---|
| `SAE-TOCOPILLA` | 116 MW | 29.493 | punto de medida (kWh) |
| `SAE-CRCA-PFV-ANDES3` | 170,78 MW | **44,07** | Gen real (MWh) |
| `SAE-CRCA-PFV-ANDES4` | 129,72 MW | **17,81** | Gen real (MWh) |

170,78 MW × 0,25 h = 42,7 — o sea los 44 son MWh y los 29.493 son kWh:
exactamente el factor 1000.

La corrección **le da sentido a la columna `Canal`**, que era un pendiente
abierto de hace varias sesiones: en la hoja `Gen real` esa columna es
ahora la **unidad** de la medida. `MWh` (o vacía, o cualquier otro texto)
= ×1000; `kWh` = ×1. El default tiene que ser MWh porque es lo que
devuelve la API: un default de "no convertir" dejaría el bug en pie para
quien no escriba nada. Qué unidad se le aplicó a cada central queda dicho
en el log — un factor de 1000 no se ve a simple vista en la hoja
`Medidores`, sólo al comparar una central contra otra.

### Las 2.232 alertas FD-005 del CPF

Reporte del usuario: *"el cálculo RE545, el CPF me dice 2.232 alertas
FD-005. Está buscando en el archivo original, pero debería buscar en el
consolidado. No sé si es eso pero por ahí va, porque en la hoja FD no
tengo diferencias en el consolidado"*.

**No era eso** (y la hoja `FD` del consolidado está bien, como él decía).
Dos cosas:

1. `FD-005` sale de **`Calculo E Costos`** (`AM:AR`), no de RE545 — RE545
   no usa el FD.
2. Lo de "busca en el archivo original" ya se había arreglado en la tanda
   anterior; y de todos modos no cambiaba nada, porque las claves que se
   buscan son idénticas vengan del archivo o del consolidado.

Reproducido contra su archivo real, las 2.232 se explican enteras:

| central | homologada a | falta en |
|---|---|---|
| `SAE-CRCA-PFV-NUEVO-QUILLAGUA-2` | `SAE-CRCA-PFV-NUEVO-QUILLAGUA-2` (a sí misma) | CPF (744) y CSF (744) |
| `SAE-CRCA-PFV-ANDES4` | `BESS PFV ANDES SOLAR IV` | CSF (744) |

744 × 3 = 2.232. Y la hoja `FD` trae una unidad que **ninguna** central
reclama: `BESS PFV MARIA ELENA`. Casi seguro es Nuevo Quillagua 2 (su
barra es `PEQ___________220`) — queda como pendiente para que el usuario
lo confirme.

**El cambio de código** no es arreglarle el diccionario, es dejar de
gritar 2.232 veces la misma frase. `calcular_fd_prorrateado()` ahora
separa dos cosas que informaba igual:

- a una unidad que **sí está** en el bloque le falta una hora suelta →
  `FD-005`, una alerta por clave, como pide el catálogo;
- una unidad **no aparece nunca** en el bloque → `FD-007` (control nuevo),
  **una alerta por central**, con la unidad a la que está homologada,
  cuántas filas quedan en 0 y **la lista de unidades que sí trae ese
  bloque** — que es justo donde aparece `BESS PFV MARIA ELENA`.

Lo segundo no es un dato faltante: es una homologación que no cruza, o una
central que no presta ese servicio. Sobre el archivo real, las 2.232
alertas pasan a ser **3**, y las 3 dicen qué hacer. Cuando se emite
`FD-007` para una unidad, no se emiten además las 744 `FD-005` de sus
horas: es una excepción deliberada a "guardar todos los faltantes" del
catálogo, anotada ahí mismo.

### Y el `FMA_CPF` vacío

El usuario confirmó que las celdas vacías de `FMA_CPF` (Nuevo Quillagua 2
y Víctor Jara) **están bien**, y pidió tolerar pero advertir. El aviso
pasó de una línea por central diciendo "se probó con el nombre tal cual" a
una sola que dice cuántas son, cuáles, y que su FMA CPF **queda en 0**.

### Verificación

- 57 pruebas (eran 49), todas verdes, sin avisos de `pyflakes`.
- `tests/test_unidades_y_fd.py` (nuevo): la conversión MWh→kWh (incluido
  el default y la combinación con el signo de `Flujo`), la lectura de la
  unidad desde `Canal`, y los tres casos de FD (unidad ausente → una
  `FD-007` por central; hora suelta → `FD-005`; sin la lista de unidades →
  se comporta como antes).
- Contra el `Consolidado_entradas.xlsx` real: las 2.232 alertas quedan en
  3, y el mensaje lista `BESS PFV MARIA ELENA` entre las unidades
  disponibles.

Lo que NO se probó: la corrida completa con el `Canal` corregido (no hay
acceso a las APIs del Coordinador desde acá).

---

## 2026-09-14 — PRORRATA_RETIROS rearmada: repartir, no validar

El usuario corrió la hoja con el archivo real del período y no obtuvo hoja
sino un error de entrada: `RET-004/005/006: Prorrata negativa: 14`. Pidió dos
cosas: dejar de lado por ahora las alertas de ese tipo y rearmar la prorrata.

**Qué estaba mal, además de las alertas.** El monto de cada cuarto de hora se
armaba sumando `Bloque Mes Descarga` (columna `Y` de `Calculo E Costos`) con
`Bloque horario` de `Calculo RE545`. `Bloque Mes Descarga` no es el cuarto de
hora del mes: es la posición del bloque dentro de la curva monótona de CMg. O
sea, se sumaba plata de dos numeraciones distintas y después se repartía como
si fuera un mismo cuarto. Ahora las dos hojas se agrupan por `Bloque horario`,
que sí es el "Cuarto de Hora" cronológico del mes renombrado.

**Cómo reparte ahora.** Para cada cuarto de hora: monto del cuarto × peso
relativo de la empresa dentro de ese cuarto (`prorrata / suma de prorratas del
cuarto`). Dividir por la suma es lo que hace que sea un reparto "según peso":
si la columna `C` ya viene normalizada la división no cambia nada, y si no
viene normalizada el cuarto igual se reparte entero en vez de caerse.

**Las alertas duras se fueron.** `leer_prorrata_retiros()` ya no levanta
`RET-001/004..007` ni `construir_prorrata_retiros()` `RET-002/003/008/009`.
Prorratas negativas, cuartos que no suman 1 y filas repetidas (estas últimas
se suman, que es lo que significan dos retiros de la misma empresa en el mismo
cuarto) se informan por el log y la corrida sigue. Quedan como error de
entrada sólo tres cosas: que el archivo no se pueda abrir, que le falte la
hoja `Prorrata 15min` y que no quede ninguna fila utilizable.

**La fuente se lee por posición.** `A` cuarto de hora, `B` suministrador, `C`
prorrata, primera fila encabezado, sin mirar el texto del encabezado (venía
exigiendo los nombres exactos `Cuarto de Hora`/`Suministrador`/`Prorrata`).

**La hoja quedó con tres cuadros**, en el orden en que se leen: `B:C` el monto
a compensar de cada cuarto de hora; `F:I` la prorrata leída (cuarto,
suministrador, peso) con el monto que le toca pagar a esa empresa en ese
cuarto; `L:M` el total del mes de cada empresa. Antes el detalle iba primero y
el monto por cuarto al medio.

**Verificación:** `python -m unittest discover` en 60 pruebas (antes 56; las
de prorrata se reescribieron y son seis) y una corrida sintética de punta a
punta que escribió la hoja y se revisó celda por celda con `openpyxl`, con
pesos negativos y cuartos que no suman 1 adentro para ver que no corta. Falta
correrla contra el archivo real del período.

---

## 2026-09-14 — Planillas ordenadas y auxiliares afuera

Pedido del usuario: *"Está funcionando y los resultados dan igual, sin
modificar el cálculo, necesito que ordenes las planillas, quita los
auxiliares innecesarios"*. Nada de esto toca un número: se verificó
corriendo el caso sintético completo antes y después y comparando las dos
hojas de cálculo celda por celda — **idénticas salvo las columnas que se
sacaron**.

### Qué salió de las hojas

| Hoja | Fuera | Por qué | Cómo se recupera |
|---|---|---|---|
| `Medidores` | M, P, Q, U | las cuatro columnas deliberadamente vacías del original | en Python no hay letra de Excel que alinear |
| `Medidores` | N `Clave_Dia_HoraMes` | clave auxiliar `Dia & Hora Mes` | no la lee nadie; `calcular_clave_auxiliar()` sigue estando |
| `Medidores` | K `Copia_Ventana` | copia fila a fila de `Ventana` | `reponer_auxiliares_medidores()` la repone al leer la hoja |
| `Calculo E Costos` | X `Ciclo` | repetía `Copia_Ventana` ("Ciclo de Carga del mes"), dos columnas con el mismo número una al lado de la otra | se sigue calculando; la hoja ya lo muestra una vez |
| `Calculo RE545` | BL | columna **sin nombre**: la suma de CMg que sólo existe para que BM saque su k-ésimo mayor | se sigue calculando; BM no cambia |
| `Calculo RE545` | BR `Ventana de Valorizacion` | repetía `T` ("Ventana de valorizacion") | se sigue calculando |

La hoja `Medidores` pasa de 18 a 12 columnas. Las que quedan mantienen el
orden de siempre (el de `LETRA_A_CAMPO`).

**Lo que NO se sacó**, aunque sea intermedio: las curvas monótonas (`W`,
`Y`, `AB`, `AC`, `AD`), el `ranking cmg`, las energías con FD y todo el paso
a paso de los Componentes 1 y 2 de RE545. El usuario lo puso así: *"si es
importante para que los coordinados entiendan y vean parte del cálculo no
hay que sacarlo, si es trivial sí"*. Tampoco las CTF en 0 (`AI`, `AL`, `AO`,
`AR`): que el CTF sea 0 es un dato, no relleno.

La lista de lo que se escribe vive en un solo lugar por hoja
(`COLUMNAS_MEDIDORES_SALIDA`, `COLUMNAS_SALIDA_E_COSTOS`,
`COLUMNAS_SALIDA_RE545`), y los encabezados de grupo se ubican contra esa
misma lista — antes se ubicaban contra `NOMBRES_CALCULO_*`, que ahora tiene
columnas que no se escriben. El grupo "Componente 1" de RE545 dejó de
nombrar a BL; sigue quedando en celdas contiguas (`BK`, `BM`, `BN`, `BO`).

### El orden de las hojas

`Pagos_BESS.xlsx` abría por `Calculo E Costos` (47 columnas de detalle) y
dejaba el `Resumen` cuarto. Ahora abre por el `Resumen` — quién paga y quién
recibe — y sigue con `Calculo E Costos`, `Calculo RE545`,
`PRORRATA_RETIROS` y las dos hojas de control (`Alertas`, `Ejecucion`) al
final. `Consolidado_entradas.xlsx` ya salía en su orden de lectura.

### El formato (`Script/nucleo/formato.py`, nuevo)

Las hojas salían tal cual las deja pandas: encabezados sin negrita, todo de
ancho 8 (los montos como `####`), sin panel fijo. Ahora, en los dos libros y
también en las hojas que se preservan de una corrida anterior:

- filas de título y de nombres de columna en negrita y centradas;
- panel inmovilizado justo debajo del encabezado (fila 1, 2 o 3 según la
  hoja: las de cálculo llevan arriba los encabezados de grupo, y
  `PRORRATA_RETIROS` el título de cada cuadro);
- ancho de columna según lo que hay adentro, con tope para que un nombre de
  central largo no empuje el resto fuera de la pantalla;
- separador de miles, dos decimales cuando la columna los tiene, y fecha
  legible en las columnas de fecha.

El formato de número se escribe celda por celda (openpyxl no tiene formato
por columna que Excel respete): medido, ~1 segundo cada millón y medio de
celdas, sobre una escritura que ya tarda varias veces eso. Hay un tope
(`MAXIMO_CELDAS_FORMATO`) para un caso disparatado; si se cruza, el libro
sale igual, sólo que sin separador de miles.

### Verificación

- 72 pruebas (eran 60), todas verdes, sin avisos de `pyflakes`.
- `tests/test_planillas_ordenadas.py` (nuevo, 12 pruebas): las columnas que
  salen de `Medidores` y el orden de las que quedan, la reposición de
  `Copia_Ventana` (incluido un libro viejo que todavía la trae, que se
  respeta), los auxiliares fuera de las dos hojas de cálculo y los
  intermedios que se quedan, los grupos cayendo en celdas contiguas, el
  `Resumen` primero, y que formatear no cambia ni un valor.
- Corrida sintética de punta a punta (`generar_pagos_bess` con las dos
  hojas) antes y después: las dos hojas de cálculo dan exactamente los
  mismos valores, columna por columna, salvo las quitadas; los encabezados
  de grupo caen sobre los mismos nombres de columna.

Lo que NO se probó: la corrida contra el archivo real del período.

---

## 2026-09-14 — Que no se abra de mas, y un boton que corre todo

Tres pedidos del usuario en la misma tanda: *"1. Que no se abran planillas
innecesarias. 2. Que si alguna información está en el consolidado que se
saque de ahí, quitar redundancias, etc. Que quede optimizadito. 3. Quiero un
botón que ejecute todo teniendo la información inicial necesaria [...] que el
botón se bloquee si falta algo [...] que se calcule solo lo restante pero que
yo pueda seleccionar si quiero repetir algún cálculo [...] Hay cosas que
pueden ir en paralelo y cosas que dependen de otras. Con cuidado"*.

Ningún número cambia: la corrida sintética de punta a punta da las mismas
hojas celda por celda antes y después.

### 1 y 2 — lo que se dejó de abrir

| Antes | Ahora |
|---|---|
| `Pagos_BESS.xlsx` abría el consolidado **cinco veces** (una por hoja: cada `pd.read_excel(ruta, sheet_name=...)` parsea el libro entero) | una sola apertura (`pd.ExcelFile`) para las cinco (`_leer_entradas_del_consolidado`) |
| ...y además reabría `cmg.xlsx` | el CMg sale de la hoja `CMg` del consolidado (`leer_cmg_consolidado`) |
| Recalcular solo `PRORRATA_RETIROS` o el `Resumen` abría igual Medidores, Ofertas SSCC, Subastas, CMg y Centrales | esas dos hojas salen de las hojas de cálculo ya escritas; no se abre el consolidado (ni se exige que exista) |
| `Centrales.xlsx` se abría dos veces por corrida del consolidado (una por "medidores", otra por "subastas") | una sola vez |
| La ventana reabría Centrales, el Excel de homologación y las dos salidas **en cada repintado** (al cambiar el AAMM, al terminar cualquier botón) | se acuerda de lo leído mientras el archivo no cambie (clave: ruta + mtime + tamaño) |

Lo del CMg es además una **redundancia peligrosa**, no solo lentitud: bastaba
con dejar un `cmg.xlsx` más nuevo en la carpeta después de generar el
consolidado para que los pagos usaran un CMg distinto del de la foto de las
entradas. Es el mismo criterio que ya se había aplicado a `FD` en la tanda
anterior. `cmg.xlsx` salió también del manifiesto de entradas de
`Pagos_BESS.xlsx`: ya no es una entrada de esa etapa.

Lo que NO se tocó: la hoja `Subastas` sigue leyendo el `SSCC_Desempeño_*`
aunque la hoja `FD` del consolidado salga del mismo archivo. No es el mismo
dato: de ahí saca el CTF y el Vector de Participación CSF, que la hoja `FD`
no guarda.

### 3 — "Ejecutar todo"

Módulo nuevo, `Script/nucleo/orquestador.py`: 17 tareas agrupadas en 8 pasos,
con sus dependencias. No calcula nada nuevo — llama a las mismas funciones
que los botones sueltos.

```
Medidas_SAE ─┐
cmg_csv → cmg.xlsx ─┤
FD → FMA ────┼→ Consolidado (5 hojas, UNA escritura) → Pagos (4 hojas, UNA escritura)
subastas ────┘
```

- **Se hace solo lo que falta.** Lo que está al día no se rehace salvo que se
  tilde; al tildarlo se tilda solo todo lo que sale de ahí
  (`propagar_seleccion`), porque rehacer `Medidores` sin rehacer los pagos
  deja el libro mezclado entre dos corridas. Si se destilda algo de lo que
  otra cosa depende, esa otra queda `bloqueada` y lo dice, en vez de calcular
  con datos viejos.
- **El botón se bloquea** si falta una entrada inicial (`Centrales.xlsx` y sus
  dos hojas, el Excel de homologación con su hoja `homol`, el `*OfertasSSCC*`,
  el SoC del período, o el `AAMM`), con el nombre de lo que falta. El Excel de
  prorrata de retiros —que llega después— **no** bloquea: sus dos hojas quedan
  fuera del plan, a la vista y con el motivo.
- **Paralelismo, con tres seguros**: (a) dos tareas que escriben el mismo
  archivo comparten `recurso` y nunca corren a la vez (por eso "Traer FD" y
  "Generar FMA", que escriben en la misma carpeta, se turnan); (b) las hojas
  de cada salida se mandan JUNTAS en una sola llamada, que además escribe el
  libro una sola vez en vez de cinco; (c) si un paso falla, no se corre nada
  que dependa de él (y el resumen final dice qué corrió y qué no).

La ventana nueva (`ventana_ejecutar_todo` en `Balance_BESS.py`) solo dibuja el
plan que devuelve `planificar()`: una fila por tarea con su checkbox, su
estado y el motivo si no puede correr, más "Solo lo que falta" / "Rehacer
todo" y el botón `Ejecutar`, que está `disabled` mientras haya bloqueos.

### Verificación

- 93 pruebas (eran 72), todas verdes, sin avisos de `pyflakes`.
- `tests/test_orquestador.py` (nuevo, 18 pruebas): el plan de cero, el de
  todo al día, el de "solo falta el Resumen", los bloqueos por entrada
  inicial y por dependencia, la propagación al rehacer algo del medio, el
  aviso de lo que queda viejo, el orden topológico, y la corrida real (con
  grupos falsos) mirando paralelismo, recurso compartido, agrupación de hojas
  en una escritura y corte de la rama que falla.
- `tests/test_control_corrida.py`: una corrida de `prorrata_retiros` +
  `resumen` **después de borrar** `Consolidado_entradas.xlsx` y `cmg.xlsx`,
  que pasa — es la prueba de que no se abren.
- La ventana se probó de verdad (Xvfb): abre, dibuja los 17 pasos con sus
  estados, el botón queda habilitado cuando se puede y bloqueado al destildar
  una dependencia.
- Corrida sintética completa antes/después: las hojas de `Pagos_BESS.xlsx`
  dan exactamente los mismos valores.

Lo que NO se probó: la corrida real contra las APIs y la unidad de red (no
hay acceso desde acá), así que el paralelismo de las cuatro bajadas está
probado con funciones falsas, no contra el Coordinador.

---

## 2026-09-14 — COMPENSACION_CENTRAL: el otro lado de la plata

La prorrata dice quién paga. Faltaba la hoja que dice quién recibe y por qué,
y verificar que el `Resumen` cruce bien los dos lados.

**Hoja nueva `COMPENSACION_CENTRAL`**, con tres cuadros: `B:E` la compensación
de `Calculo E Costos` por central y `Ciclo de Carga del mes`; `H:K` la de
`Calculo RE545` por central y `Ventana de valorizacion`; `N:O` el total que
recibe cada empresa.

**Por qué cada hoja va con una agrupación distinta.** No es una decisión de
presentación: es cómo se calcula el monto. `AZ` de E Costos es
`(suma AX - suma U) / filas` por grupo **(central, `Copia_Ventana`)**, repetido
en todas las filas del grupo; `CE` de RE545 reparte
`max(suma BO - suma CC, 0)` por grupo **(central, `T` = ventana de
valorización)** en proporción a `AU`. En los dos casos sumar las filas del
grupo devuelve el total del grupo, así que agrupar y sumar es correcto para
las dos — pero cada una por SU columna.

**La empresa sale del `Propietario` de `Resumen BESS`** (`Centrales.xlsx`).
Una central sin propietario mapeado queda a su propio nombre y se avisa por el
log, en vez de perder la plata en una fila con la empresa vacía.

**Bug del `Resumen`, que es lo que se pidió verificar.** Los dos lados vienen
de fuentes distintas: `RECIBE` del `Propietario` de `Resumen BESS`, `PAGA` del
`Suministrador` de la prorrata del CEN. El merge era por el texto crudo, así
que `"COLBUN S.A."` y `"colbun  s.a."` no cruzaban: la misma empresa aparecía
en dos filas, cada una con la mitad de la historia y un `NETO` que no era su
neto. Ahora el cruce es por nombre normalizado (minúscula, sin tildes, sin
espacios de más) y el nombre visible es el del `Propietario` cuando la empresa
recibe. Además el log dice cuántas empresas reciben y pagan a la vez, y avisa
si el total recibido no es igual al total pagado (son la misma plata; la
diferencia es lo que no se repartió).

**Reorganización menor.** `construir_compensacion_total` y `construir_resumen`
se fueron de `prorrata_retiros.py` a un `compensacion.py` nuevo: un módulo por
lado de la transacción (`prorrata_retiros.py` = quien paga, `compensacion.py` =
quien recibe). La llamada a `escribir_pagos_bess()` pasó a ser toda por nombre:
son once tablas y el orden posicional ya se prestaba a confusión.

**Verificación:** `python -m unittest discover` en 66 pruebas (eran 60; seis
nuevas en `tests/test_compensacion_central.py`), más una corrida sintética que
escribió la hoja y se revisó celda por celda con `openpyxl`, y un caso con el
nombre de la empresa escrito distinto en cada lado para ver que el `Resumen`
la deja en una sola fila. Falta correrla contra el archivo real del período.

---

## 2026-09-14 — La rueda del mouse: el registro se desplaza solo

Dos cosas que el usuario reportó de la ventana: que desplazarse "se ve como
pegado", y que el registro no tiene su propio desplazamiento (poniendo el
mouse encima se movía toda la ventana).

**Las dos salían de la misma línea**, el binding que había:

```python
canvas.bind_all("<MouseWheel>",
                lambda e: canvas.yview_scroll(int(-e.delta / 120), "units"))
```

1. `bind_all` es global: la rueda movía la ventana estuviera donde estuviera
   el puntero, incluido encima del registro.
2. `int(-delta / 120)` trunca hacia cero. Un touchpad de precisión de Windows
   manda deltas **menores a 120** (40, 60...), que son fracciones de muesca, y
   todos daban 0: no se movía nada hasta que el gesto era grande y ahí saltaba
   de golpe. Eso es lo "pegado".
3. Y como yapa, una "unidad" de `Canvas` sin `yscrollincrement` es un décimo
   del alto visible: cada muesca era un salto enorme.

**Cómo quedó.** La rueda la atiende un `rueda()` que mira qué widget está
**debajo del puntero** (`winfo_containing`) y sube por los padres hasta
encontrar uno anotado en `desplazables`. El registro está anotado: con el
puntero encima se desplaza el registro y nada más, ni siquiera cuando ya está
en un extremo (encadenar ahí es justo lo que hacía que recorrer el registro
terminara moviendo la ventana). Fuera del registro, se desplaza la ventana.

El resto fraccionario de cada evento se acumula (`_acumular`), así que los
deltas chicos del touchpad suman en vez de perderse, y el resto se descarta al
cambiar de sentido. El `Canvas` va con `yscrollincrement=1`: desplaza de a
píxeles (45 por muesca), no de a décimos de pantalla. El registro va de a
líneas (3 por muesca), que es lo natural en un widget de texto.

**Un detalle que no es obvio:** la rueda además se ata al widget mismo
(`registrar_desplazable`), no solo al diccionario. Las ataduras de widget
corren ANTES que las de clase, así que el `"break"` evita que la atadura de
clase de `Text` —que también desplaza— lo mueva una segunda vez. Sin eso, una
muesca sobre el registro movía el doble.

**Verificación:** `tests/test_ventana_rueda.py` (nuevo) abre la ventana de
verdad y le manda eventos `<MouseWheel>` reales: que sobre el registro se mueva
solo el registro, que fuera se mueva solo la ventana, que el registro al final
no arrastre la ventana, que tres deltas de 40 sumen exactamente una muesca, y
que una muesca sobre el registro no mueva el doble. La suite quedó en 71
pruebas. Las cinco de la ventana se saltean si no hay tkinter o no hay display
(acá corrieron con `xvfb-run`); en Windows corren solas.

---

## 2026-09-15 — Un mes que todavia no existe: crear la carpeta con todo adentro

Pedido del usuario: *"cuando yo elija un mes que no exista, me permita elegir
y crear la carpeta con las carpetas dentro"*.

**Cuando se ofrece.** Al escribir un período nuevo arriba, la ventana mira la
carpeta que está abierta y pregunta solo cuando está segura: o no existe, o su
nombre trae un AAMM (el del mes anterior) distinto del que se acaba de
escribir. Si el nombre no tiene ningún AAMM no se puede saber y no se
pregunta nada — mejor callarse que molestar en cada `FocusOut`. También está
el botón fijo **Crear carpeta del caso**, para cuando se quiere hacer a mano.

**Qué propone.** No se impone ninguna convención de nombre: se toma el nombre
de la carpeta que se estaba usando y se le cambia SOLO el AAMM
(`Balance BESS 2607` → `Balance BESS 2608`, `nombre_caso_sugerido()`); si no
trae ninguno, se propone el AAMM pelado. La ventana muestra ese nombre y la
carpeta donde crearla (por defecto, al lado de la del mes anterior) en dos
campos editables, más la lista de lo que va a crear. Recién ahí se crea, y el
caso queda abierto en ese período.

**Qué crea** (`crear_estructura_caso()`, `SUBCARPETAS_CASO`): la carpeta base
y `Medidas/`, `Auxiliares/`, `Ofertas/`, `Cmg/`, `FD y FMA/`, `Subastas/`,
`Subastas/DB subastas/` y `Prorrata retiros/`. NO crea `Medidas/_trabajo`
(son los andamios de la descarga, los arma `generar_medidas_sae`).

Tres cuidados:

- **Es idempotente y no toca nada**: crea solo las carpetas que faltan, nunca
  borra ni mueve. Por eso sirve igual para completar un caso al que le falta
  una subcarpeta — lo que ahora también se ofrece al elegir carpeta con
  "Examinar" y al cambiar de período sin cambiar de carpeta.
- **Un caso viejo con `SSCC_Desempeño/`** (el nombre anterior de `FD y FMA/`)
  no recibe una `FD y FMA/` vacía al lado: se respeta la que ya está usando,
  que es la misma regla de `resolver_rutas()`.
- Si el nombre elegido ya existe como ARCHIVO, se dice y no se hace nada.

### Verificación

- 103 pruebas (eran 93), todas verdes, sin avisos de `pyflakes`.
- `tests/test_caso_nuevo.py` (nuevo, 10 pruebas): la carpeta que nace con
  todo adentro, la segunda corrida que no hace nada, completar solo lo que
  falta sin tocar un archivo que ya estaba, el caso viejo que conserva su
  carpeta de FD, el nombre ocupado por un archivo, y que
  `revisar_estructura()` sobre el caso recién creado no marca NINGUNA carpeta
  en falta (las que se crean son exactamente las que el programa busca).
  Más el nombre sugerido: cambia el AAMM y conserva el resto, no se come un
  número de 6 dígitos, y avisa si el período no es válido.
- La ventana se probó de verdad (Xvfb): se escribe `2608` con un caso
  `Balance BESS 2607` abierto, aparece la pregunta, la ventana propone
  `Balance BESS 2608` en la carpeta de al lado, se crea con sus 7
  subcarpetas + `DB subastas`, y la ventana queda abierta en ese caso con el
  período guardado.

---

## 2026-09-15 — El FD sale de `03 Desempeño para publicar` (unidad `F:`), y la ventana es toda links

**Lo que pidió el usuario**, textual: que los FD se saquen de
`F:\11 SSCC\05 Verificación SSCC\02 Cálculo indicadores\2026\08. Agosto\Indicadores Publicar\V1\03 Desempeño para publicar`
("genérico, si po, te mando agosto v1 de ejemplo"), que si encuentra V2 se
quede con V2 y no vaya a buscar V1; que en el detalle de cada archivo con
botón **"Traer"** diga `Origen: DCO` con `DCO` como link a la ruta exacta, y
que los FMA —que no son "traer" pero tienen inputs que sí se traen— digan
`Origen inputs: ...`, también con link; y que cada archivo y cada carpeta de
la estructura sea link a su ruta (rutas, no abrir archivos).

**La ruta del FD cambió en dos tramos.** `RAIZ_DCO_INDICADORES` pasó del UNC
`\\nas-cen1\DCO\11 SSCC\05 Verificación SSCC\02 Cálculo indicadores` a
`F:\11 SSCC\05 Verificación SSCC\02 Cálculo indicadores` (es el mismo árbol
montado de otra forma), y `SUBCARPETAS_FD` de
`04 Desempeño para transferencias` a `03 Desempeño para publicar`. La carpeta
anterior NO se borró: quedó como `SUBCARPETAS_FD_ANTIGUA` y se prueba después
de la nueva (`CADENAS_FD`), porque los meses ya cerrados siguen teniendo el FD
ahí y el botón tiene que servir igual para volver a un período viejo. Recién
si en ninguna de las dos hay nada se cae a la búsqueda recursiva de siempre.

**Lo de V2 sobre V1 ya estaba** y no se tocó: `buscar_en_versiones()` recorre
las versiones de mayor a menor y se queda con la primera que TENGA el archivo.
Con V1 y V2 publicadas usa V2 sin mirar V1; sólo baja a V1 si la carpeta V2
existe pero está vacía (recién publicada, a medio subir). Eso es exactamente
"la mayor que tenga disponible el archivo" y hay dos pruebas nuevas que lo
fijan.

**`Script/nucleo/origenes.py` (módulo nuevo).** Un `id -> (título, etiqueta,
resolver)` para las siete filas que vienen de afuera del caso: `fd` (Origen:
DCO), `cmg_csv` (Origen: CMg Reales), `subastas` (Origen:
progdiar_adjudicaSEN) y, como **"Origen inputs"**, `cmg_xlsx`, `fma_cpf`,
`fma_csf` y `fma_ctf`. La distinción es a propósito: el FMA no se copia hecho,
se construye — lo que viene de afuera son sus insumos.

**La ruta se resuelve al hacer click, no al pintar.** Resolver el origen
implica mirar el servidor (qué versión está publicada, si existe la carpeta
del mes) y el diagrama se repinta en CADA revisada: con la unidad de red
desconectada eso congelaría la ventana varios segundos por repintado. Así que
en la fila viaja sólo la etiqueta, y `abrir_origen()` resuelve y abre en un
hilo aparte. Las funciones que resuelven (`ruta_origen_fd`,
`ruta_origen_cpf/csf/ctf`, `carpeta_origen_csv`) **no levantan**: si no se
puede leer el servidor arman igual la ruta que le correspondería al período
(con `VERSION_POR_OMISION = V1`), para poder abrir el árbol y mirar; sin AAMM
válido devuelven la raíz.

**Cada fila del diagrama trae ahora su ruta.** `revisar_estructura()` agrega
tres campos a cada fila: `ruta`, `es_carpeta` y `origen`. La ventana convierte
el nombre en link (subrayado, azul, `hand2`) y el click abre **la carpeta**:
la propia si la fila es una carpeta, la que contiene al archivo si es un
archivo — nunca se abre el archivo, para no arrancar Excel sin que se lo
pidan, que es lo que pidió el usuario ("no abrir archivos sino rutas"). Si la
ruta todavía no existe (una carpeta que falta, una salida sin generar, el mes
que el DCO no publicó), `carpeta_a_abrir()` sube hasta el primer ancestro que
sí exista: el click siempre lleva a algún lado. Las filas de hoja (`hoja
'Medidores'`, etc.) no llevan link: no son archivos.

Detalle de dibujo: el prefijo del árbol (`├── `) se separó en su propia
etiqueta para que el subrayado del link tape sólo el nombre. Las dos son
Consolas 9 y juntas miden lo mismo que la columna de antes
(`ANCHO_ESTRUCTURA`), así que las columnas no se movieron.

**Verificación:** `python -m py_compile` de todo y `python -m unittest
discover` en 77 pruebas (antes 60; las 17 nuevas están en
`tests/test_origenes_fd.py`: la raíz `F:`, la carpeta nueva, que el FD viejo
se siga encontrando en la carpeta anterior, V2 sobre V1, la caída a V1 cuando
V2 no tiene el archivo, las rutas de origen con y sin servidor, y que cada
fila del diagrama traiga su ruta/origen). Además se pintó la ventana entera
una vez bajo Xvfb y se volcó fila por fila lo que queda dibujado, para
confirmar que cada archivo y carpeta salió como link y que los "Origen:" /
"Origen inputs:" quedaron donde tenían que quedar. **Falta** abrirla en
Windows con la unidad `F:` conectada: que la ruta del ejemplo (agosto 2026)
resuelva de verdad y que el click abra el explorador es lo único que no se
puede probar desde acá.

---

## 2026-09-15 (2) — La unidad de "Gen real" se lee del PRINCIPIO del canal (`MWhD`/`MWhR`)

**El reporte del usuario:** "cambié lo de los MWhD MWhR de la homologación y no
cambia; esa unidad debe ser la que indica si multiplicar por mil o no, y tiene
R y D al final, pero lee lo primero".

**Qué pasaba.** `leer_gen_real()` comparaba el texto ENTERO de la columna
`Canal` contra el diccionario (`texto in UNIDADES_GEN_REAL`, o sea exactamente
`"mwh"` o `"kwh"`). En el archivo real el canal no viene pelado: viene `MWhD` /
`MWhR` — la unidad con el tipo de medida pegado atrás. Ninguno de los dos
coincidía, así que todas las filas caían en `UNIDAD_GEN_REAL_POR_DEFECTO`
(`mwh`) y la columna quedaba de adorno: cambiarla no cambiaba nada. Con `MWhD`
el resultado igual era el correcto por casualidad (el default ES MWh), pero un
`kWhD` / `kWhR` se multiplicaba por mil igual — el error silencioso que la
columna existe para evitar.

**Lo que se hizo.** `unidad_desde_canal(valor)` (nuevo, en `Homologacion.py`):
normaliza el texto y se queda con la unidad con la que **empieza**. La `D` y la
`R` del final no entran en la cuenta: lo único que decide esta columna es si se
multiplica por mil o no. Vacío, o un texto que no empieza con ninguna de las
dos, sigue valiendo MWh, que es lo que devuelve la API de operación real.

`MWhD`/`MWhR` → `mwh` (x1000) · `kWhD`/`kWhR` → `kwh` (x1) · `MWh`/`kwh`
sueltos siguen funcionando igual · `""`/`None`/`Potencia` → `mwh`.

**Verificación:** `python -m unittest discover` en 79 pruebas (antes 77). Las
dos nuevas están en `tests/test_unidades_y_fd.py`: la tabla de canales con
sufijo, y que la MISMA respuesta de la API entra x1000 con `MWhR` y tal cual
con `kWhR` — que es la prueba de que ahora el cambio en el archivo se nota.
---
## 2026-09-15 (3) — Traer a main la rama que habia quedado afuera (COMPENSACION_CENTRAL + la rueda)

Al revisar qué ramas se podían borrar aparecieron **tres commits del 14 que
nunca volvieron a `main`**: la rama `claude/gallant-einstein-jmunn6` siguió
recibiendo trabajo *después* de que su PR (#22) se mergeara. Lo que traía:
la hoja `COMPENSACION_CENTRAL` (módulo `Script/nucleo/compensacion.py`) con el
cruce por nombre normalizado en `Resumen`, el arreglo de la rueda del mouse
(el registro se desplaza solo y sin saltos, con `tests/test_ventana_rueda.py`)
y la nota de `REGLAS.md` sobre cómo correr las pruebas de la ventana sin
pantalla.

**Los dos conflictos:**

- `escritura.py`, `_HOJAS_PAGOS`: las dos ramas tocaron el orden de las hojas
  del libro de pagos. Una puso el `Resumen` **primero** ("lo primero que se
  mira") y la otra agregó `COMPENSACION_CENTRAL` manteniendo el `Resumen` al
  final. Quedó el orden nuevo con la hoja nueva adentro: `Resumen`,
  `Calculo E Costos`, `Calculo RE545`, `COMPENSACION_CENTRAL`,
  `PRORRATA_RETIROS` (y atrás las dos de control).
- `BITACORA.md`: entradas agregadas al final por las dos. Quedaron todas, en
  orden de fecha: las cuatro del 14 y después las tres del 15.

**El problema de verdad no era un conflicto, era un hueco.** Las dos ramas no
se vieron nunca: `COMPENSACION_CENTRAL` se agregó en una y el grafo de
"Ejecutar todo" (`orquestador.py`) en la otra. Git mezcló las dos sin
quejarse y el resultado tenía la hoja con su botón de fila **pero afuera del
grafo**: "Ejecutar todo" no la generaba nunca y quedaba PENDIENTE para
siempre. Se agregó su `Tarea` (`pagos:compensacion_central`, depende de las
dos hojas de cálculo, requiere `Centrales.xlsx`/`Resumen BESS` por el
Propietario) y ahora sale en la misma escritura que las otras cuatro.

Para que no se repita: `TestGrafoCompleto` (nuevo, en
`tests/test_orquestador.py`) recorre `SECCIONES_CONSOLIDADO` y
`SECCIONES_PAGOS` y exige que **cada hoja tenga su tarea** en el grafo. La
próxima hoja que se agregue en otra rama va a hacer fallar esa prueba en vez
de desaparecer en silencio.

**Verificación:** `py_compile` de todo y **136 pruebas** (antes 122 + las 11
que traía la rama + 3 nuevas del grafo). Las 5 de la ventana corrieron de
verdad con `xvfb-run` (con el intérprete sin tkinter se saltean). Además se
volcó la ventana entera: quedó con **25 botones** — el que faltaba era
**"Resumir compensación"**, el de la fila `hoja 'COMPENSACION_CENTRAL'` —, y
la ventana del plan ahora lista las 18 tareas, con la compensación adentro.

---

## 2026-09-15 (4) — "Traer FD" traía media publicación: se pega a una sola carpeta

**El reporte:** *"SSCC_Desempeño se están descargando de otra parte. La ruta es
`…\Indicadores Publicar\V2\03 Desempeño para publicar`. Si no está en V2 se
reemplaza por V1, esa es la lógica, no descargar nada más. Cuando corro Traer
FD me trae muchos que vienen de otras rutas. Y los links de la ventana tienen
que apuntar a la ruta final a la que se ocupó, V1 o V2 como sea."*

**Por qué traía de más.** `buscar_archivos_fd()` tenía dos escalones de
tolerancia que en el árbol real se comen media publicación:

1. dentro de la carpeta del FD buscaba con `rglob` (o sea, también en sus
   subcarpetas);
2. si ahí no encontraba, se caía a un `rglob` **desde la carpeta de la
   versión entera** — cualquier `SSCC_Desempeño_*` o
   `SSCC_Disponibilidad_CSF*` con el año en el nombre, colgara de donde
   colgara (`01 Respuesta`, `02 …`, lo que sea), entraba a la lista y se
   copiaba al caso.

Esa tolerancia la habíamos puesto nosotros "por si el DCO cambia el
anidamiento". La regla real es más simple y es la que pidió el usuario: esa
carpeta, sin recursión, en V2 y si no en V1. Nada más.

**Lo que quedó.** `buscar_archivos_fd()` mira sólo los archivos **sueltos** de
`<version>/03 Desempeño para publicar` (`carpeta_fd_de_la_version()`, nueva) y
devuelve lista vacía si esa versión no tiene la carpeta. La elección de
versión no cambió: `buscar_en_versiones()` ya recorría de mayor a menor y se
queda con la primera que TENGA el archivo.

También se sacó la carpeta anterior (`04 Desempeño para transferencias`) de la
búsqueda: ya no se busca ahí. `SUBCARPETAS_FD_ANTIGUA` queda sólo para el
mensaje de error — si una versión no tiene la carpeta nueva pero sí la vieja,
el error lo dice con todas las letras, así un mes viejo falla explicando por
qué en vez de traer cualquier cosa.

**Los links, a la ruta que se ocupó.** `ruta_origen_fd()` resolvía la versión
más alta que EXISTIERA (`elegir_version`), así que con una V2 publicada pero
vacía el botón copiaba de V1 y el link apuntaba a V2. Ahora
`carpeta_version_usada()` recibe la **misma función de búsqueda que usa el
botón**, así que link y carpeta de origen son siempre la misma. Lo mismo para
los "Origen inputs" de CPF y CTF (`ruta_origen_cpf/ctf` pasan
`buscar_reportes_cpf`/`buscar_ctf`): cada entrada elige su versión por
separado, y cada link sigue a la suya.

**Verificación:** 140 pruebas (antes 136; 4 nuevas en
`tests/test_origenes_fd.py`: que no busca en la carpeta vieja, que no trae
nada de otras ramas del árbol, que ignora las subcarpetas de la carpeta del
FD, y que el link apunta a V1 cuando V2 está vacía). Además dos corridas de
`traer_fd()` de punta a punta contra un árbol armado como el real —V2 con el
zip bueno y cuatro ramas con archivos de nombre parecido— : copió **un solo
archivo**, descomprimió el Excel de adentro y el link quedó en
`…/V2/03 Desempeño para publicar`; con la V2 vacía, copió de V1 y el link
quedó en V1.

**Ojo, no se tocó:** `buscar_reportes_cpf()` y `buscar_ctf()` (los insumos del
FMA) siguen teniendo su caída a búsqueda recursiva desde la carpeta de la
versión. No traen archivos al caso —se leen en el lugar—, pero si aparece el
mismo síntoma (un CPF/CTF que sale de donde no corresponde), es el mismo
patrón y se arregla igual.

---

## 2026-09-15 (5) — El `SyntaxWarning: invalid escape sequence '\S'` del arranque

Aviso que salía en la terminal al correr el programa (con Python 3.12; en 3.11
no se ve porque ahí todavía es `DeprecationWarning`). No rompía nada —el
programa corría igual—, pero conviene no acostumbrarse a ver avisos.

**Qué era.** El docstring de `_extraer_nombre_desde_ruta_scada()`
(`Script/nucleo/soc.py`) muestra como ejemplo la ruta SCADA real
`\\SERVIDOR\SEN\Generación\SEN\<region>\<central>|<sufijo>`. En una cadena
normal de Python, `\S` no es un escape válido: hoy Python lo interpreta como
las dos letras y avisa, y en alguna versión futura va a ser error. El aviso
apuntaba a la línea del docstring, no a código que se ejecute.

**Arreglo:** el docstring pasó a ser crudo (`r"""`), que es lo que corresponde
cuando el texto tiene barras invertidas de verdad. Nada de lógica cambió.

**Verificación:** importar `Balance_BESS` entero con Python 3.12 ya no imprime
ningún aviso, y se revisó que no quedara otro `SyntaxWarning` en ningún `.py`
del repo. 140 pruebas, igual que antes.

---

## 2026-09-15 (6) — Una sola planilla (de fin a inicio), el control aparte y la ventana derecha

Cuatro pedidos del usuario en la misma tanda: *"puedes hacer que las hojas de
Alertas, Ejecución y agregar log queden en una planilla diferente, y combinar
el consolidado entradas con pagos bess pero ordenados de fin a inicio, el fin
es el resumen y el inicio las entradas"*, *"en la ventana se ve desordenado la
columna de estados y los botones quiero que se vean rectos"*, *"los botones de
resumir y traer que digan solo eso"*, *"quita los detalles, deja solo los que
dan el origen de entradas y esos links"*.

**1. Una sola planilla, ordenada de fin a inicio.**
`Consolidado_entradas.xlsx` + `Pagos_BESS.xlsx` = **`Balance_BESS.xlsx`**
(`ARCHIVO_SALIDA`). Las hojas quedan en `ORDEN_HOJAS_SALIDA`
(`parametros.py`): `Resumen`, `PRORRATA_RETIROS`, `COMPENSACION_CENTRAL`,
`Calculo RE545`, `Calculo E Costos`, `Subastas`, `FD`, `CMg`, `Ofertas SSCC`,
`Medidores`.

Las dos mitades se siguen escribiendo por separado (`escribir_salida()` /
`escribir_pagos_bess()`: cada hoja tiene su botón y su cálculo detrás), así
que lo nuevo es que **cada una preserva las hojas de la otra**
(`_preservar_ajenas()`) y reordena el libro al terminar (`_ordenar_hojas()`).

**Trampa que hay que tener presente:** el `ExcelWriter` **trunca** el archivo.
Por eso ahora el libro anterior se lee SIEMPRE antes de abrirlo
(`_abrir_existente()`), no solo cuando la corrida es parcial: si eso se
saltea, escribir una hoja se lleva puestas las otras nueve. También por eso
los grupos `consolidado` y `pagos` del orquestador comparten el recurso
`"salida"`: escriben el mismo archivo y no pueden correr a la vez.

**2. El control, en otra planilla.** `Alertas`, `Ejecucion` y `Log` salieron
de la planilla de trabajo y viven en **`Control_corrida.xlsx`**
(`ARCHIVO_CONTROL`), que escribe la nueva `escribir_control()` — pública, con
el mismo criterio de preservar lo que esa pasada no reescribe (la corrida de
entradas deja el `Log` sin tocar `Alertas`/`Ejecucion`, y al revés).

**3. La ventana derecha.** Las columnas del diagrama eran `width=N` en
*caracteres*: una fila en negrita (las de nivel 0) mide más que la misma fila
en redonda aunque las dos digan 52 caracteres, y por eso el estado y los
botones de esas filas quedaban corridos. Ahora cada celda es un `Frame` de
ancho fijo **en píxeles** con `pack_propagate(False)`, y el botón ocupa su
celda entera: medido con la ventana real (`xvfb`), las 33 filas tienen la
columna `Estado` en x=420, la de `Acción` en x=506 y **todos** los botones
miden 96x22.

**4. Botones de una palabra y sin detalle.** `Traer cmg_15min`/`Traer FD`/
`Traer subastas`/`Traer prorrata` → **Traer**; `Resumir compensación` →
**Resumir**; `Asignar pagos` → **Asignar**; `Actualizar todo`/`Calcular todo`
→ **Actualizar** (una sola fila de archivo: hace las dos mitades seguidas,
`actualizar_planilla_entera()`). La columna `Detalle` se sacó entera: a la
derecha del botón queda solo el `Origen: ...` con su link, que es lo que el
usuario pidió conservar. El motivo de un error al revisar la carpeta ahora va
al registro, que es donde se puede leer entero.

**Verificación:** 144 pruebas (4 nuevas: orden de las hojas, que cada mitad
preserve la otra, que el control salga aparte y que conserve sus hojas).
Además se abrió la ventana de verdad con `xvfb-run` sobre un caso sintético
para medir las columnas y mirar la captura.

**Ojo:** lo único que NO se probó contra un caso real es la migración de un
caso viejo: una carpeta que ya tenga `Consolidado_entradas.xlsx` y
`Pagos_BESS.xlsx` no se lee ni se convierte — el programa empieza una
`Balance_BESS.xlsx` nueva y los dos archivos viejos quedan donde están, sin
que nadie los toque. Si hace falta arrastrar el contenido viejo, hay que
volver a generar las hojas (o copiarlas a mano).

## 2026-09-15 (7) — Un periodo, una carpeta: cambiar de mes cambia de carpeta

Reporte del usuario: *"cuando cambio el mes en la casilla se mantiene en la
carpeta; dos meses no pueden tener la misma carpeta; si cambio el año y el mes
no tiene carpeta debe pedirme examinar y seleccionar o crear"*.

**Qué estaba mal.** La ventana sí preguntaba... pero solo cuando el nombre de
la carpeta abierta traía un AAMM distinto del nuevo. El viejo
`_es_de_otro_periodo()` decía, textualmente, *"si el nombre no tiene ningún
AAMM no se puede saber, y entonces no se pregunta nada: mejor callarse que
molestar"*. O sea: con una carpeta llamada "Balance BESS" (o cualquier nombre
sin 4 dígitos sueltos), cambiar de mes seguía trabajando sobre la carpeta del
mes anterior, en silencio. Eso es exactamente lo que el usuario vio.

**El arreglo, en dos partes.**

1. `nucleo.carpeta_corresponde_al_periodo()` (en `rutas.py`, con pruebas)
   devuelve **tres** respuestas, no dos: `CARPETA_DEL_PERIODO`,
   `CARPETA_DE_OTRO_PERIODO` y `CARPETA_SIN_PERIODO`. "No se sabe" deja de
   ser un sí: obliga a preguntar.
2. La ventana recuerda, por PC/usuario en `config.json`, **qué carpeta usó
   cada período** (`carpetas_por_periodo: {aamm: carpeta}`). Eso es lo que
   hace cumplible la regla aunque el nombre no diga nada: una vez que una
   carpeta quedó asociada a un mes, otro mes ya no la puede reusar. El par
   (carpeta, período) con el que arranca la ventana se anota solo.

**Qué pasa ahora al cambiar el AAMM:**

- El período ya tuvo carpeta → se abre esa, sola, sin preguntar.
- La carpeta abierta es la de otro mes → ventana con **Examinar…** /
  **Crear la carpeta** / **Cancelar**.
- No se puede saber de qué mes es (nombre sin AAMM y nunca anotada) → lo
  mismo, más un tercer botón **"Esta es la de AAMM"**, que la deja anotada.
- **Cancelar vuelve al período anterior**: la ventana nunca queda con un mes
  escrito arriba apuntando a la carpeta de otro.

**Dos detalles de tkinter que costaron.** `ventana_nuevo_caso()` ahora es
modal y **devuelve** si el caso se creó: antes se abría y se seguía de largo,
así que cancelarla dejaba el período nuevo sobre la carpeta vieja. Y mientras
se pregunta, el campo del AAMM pierde el foco, con lo que `<FocusOut>` vuelve
a llamar a `aamm_cambiado()` y se abría una segunda ventana encima de la
primera: la marca `preguntando` lo corta.

**Verificación:** 151 pruebas (7 nuevas sobre
`carpeta_corresponde_al_periodo`). Además se abrió la ventana real con
`xvfb-run` y se recorrieron los cuatro caminos apretando los botones de
verdad: cambiar de mes con la carpeta del mes anterior abierta (pregunta),
cancelar (vuelve el período anterior), Examinar (cambia y anota), volver al
mes anterior (vuelve solo), carpeta sin AAMM en el nombre (aparece el tercer
botón) y Crear la carpeta (queda creada con sus 7 subcarpetas y el caso
abierto ahí).
