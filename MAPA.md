# MAPA.md — qué hace cada script

Primera lectura obligatoria de cualquier sesión (ver `METODOLOGIA.md` §2).
Un bloque corto por script: qué hace · consume · produce · expone · depende
de.

---

## Estructura del repositorio

```
Balance_BESS.py            <- la ventana (lo unico que se ejecuta)
config.ejemplo.json        <- el formato de "claves_api" para copiar
config.json                <- local, NO se versiona (claves + por usuario)
Script/
    __init__.py
    config.py              <- lo unico que lee y escribe config.json
    nucleo/                <- el calculo del caso, una etapa por modulo
        __init__.py            <- la fachada: nucleo.<lo que sea>
        externos.py            <- los paquetes hermanos, en un solo lugar
        parametros.py          <- nombres de archivo, carpeta y hoja
        utiles.py              <- normalizacion y ErrorEntrada
        alertas.py             <- el registro de alertas y el estado
        avisos.py              <- los avisos de cruces que darian cero
        conciliacion.py        <- Medidores = E Costos + RE545 (TRA)
        manifiesto.py          <- que archivo exacto alimento la corrida
        rutas.py               <- rutas del caso y busqueda de entradas
        estructura.py          <- el arbol que dibuja la ventana
        origenes.py            <- de donde viene lo que se trae (links)
        lectura.py             <- Medidas_SAE.xlsx y Centrales.xlsx
        soc.py                 <- el SoC por bloques del SCADA
        ofertas_sscc.py        <- la hoja Ofertas SSCC (y R, S, T, V)
        hojas_entrada.py       <- hojas CMg, FD y Subastas
        subastas_accdb.py      <- Subastas desde los Access
        fma.py                 <- Subastas!Q (FMA) y Subastas!P (FD)
        diccionarios.py        <- los diccionarios de Resumen BESS y CMg
        columnas_compartidas.py<- L, M y N/O: iguales en las dos hojas
        ecostos.py             <- Calculo E Costos: base y orquestacion
        ecostos_columnas.py    <- E Costos etapa 2
        ecostos_prorratas.py   <- E Costos etapa 3
        ecostos_ciclo.py       <- E Costos etapa 4
        re545.py               <- Calculo RE545: base y orquestacion
        re545_reservas.py      <- RE545 AC:AT y AU
        re545_resumen.py       <- RE545 AW:BG y BV
        re545_componentes.py   <- RE545 BI:CE
        prorrata_retiros.py    <- prorrata 15-min: quien paga, cuarto a cuarto
        compensacion.py        <- quien recibe: por central/ciclo y por empresa
        medidores.py           <- la hoja Medidores
        escritura.py           <- los dos libros de salida
        formato.py             <- ancho/negrita/panel fijo de las hojas
        proceso.py             <- los dos procesos completos
        orquestador.py         <- el grafo de "Ejecutar todo" (plan y corrida)
        traer.py               <- los botones Traer/Generar
        medidas_sae.py         <- Medidas_SAE.xlsx
    Cmg/
        __init__.py
        Extrae_CMG_barras.py   <- arma cmg.xlsx desde el CSV 15-minutal
    Fd/
        __init__.py
        Indicadores_DCO.py     <- trae el FD del arbol del DCO
        Indices_FMA.py         <- arma las tres salidas de FMA
        Desempeno_Horario.py   <- el FD por unidad y hora
    Subastas/
        __init__.py
        Ofertas_Adjudicadas.py <- trae y lee los Access de subastas
        Fma.py                 <- normaliza las tres salidas de FMA
    Medidas/
        __init__.py
        comun.py               <- ErrorMedidas + helpers de texto
        Homologacion.py        <- el Excel de Auxiliares/ (punto+canal -> clave)
        Descarga_PRMTE.py      <- API de medidas, por punto de medida
        Claves_Balance.py      <- calendario de cuartos + agrupacion por clave
        Generacion_Real.py     <- API de operacion real (hoja "Gen real")
```

`Script/` es un paquete: la ventana hace `from Script import nucleo` y
`nucleo/externos.py` hace `from ..Cmg import Extrae_CMG_barras` (el único
lugar del núcleo donde se importan los paquetes hermanos).

`nucleo` era un solo archivo de 8.300 líneas; hoy es un paquete con un
módulo por etapa. `nucleo/__init__.py` es **solo una fachada**: re-exporta
todo, así que `nucleo.lo_que_sea` sigue funcionando igual y ni la ventana
ni las pruebas tuvieron que cambiar. Los módulos se importan entre sí en
una sola dirección (de `parametros`/`utiles` hacia las etapas, y de las
etapas hacia `proceso`): no hay ciclos.

El nombre del módulo de CMg usa guiones bajos, no espacios, para que sea
importable como cualquier módulo.

---

## `Balance_BESS.py`

- **Qué hace:** ventana tkinter única. Deja elegir la carpeta base del caso
  e ingresar el **período (AAMM, 4 dígitos, ej. `2607`)** en un campo de
  texto, y debajo dibuja un **diagrama de la estructura del caso** (árbol de
  texto tipo consola, prefijos `├──`/`└──`/`│`, patrón tomado del
  `Revisor_Reliquidacion.py` que el usuario dio como referencia) con el
  estado de cada entrada (OK/FALTA/PENDIENTE). El AAMM no se infiere del
  nombre de ningún archivo: lo escribe el usuario, y **no** es una fila del
  diagrama (no es parte de la estructura de carpetas) — las dos filas que
  dependen de él (el SoC dentro de `Medidas/`, el CSV dentro de `Cmg/`) lo
  dicen en su propio detalle cuando falta.

  **Cada acción es un botón en la fila que le corresponde**, sin ventanas
  intermedias. Abajo de todo hay además tres botones que no son de una fila
  sola: **Crear carpeta del caso** (arma el caso del período con todas sus
  subcarpetas — ver `rutas.py`), **Ejecutar todo** (abre el plan de la
  corrida y la ejecuta respetando el grafo de dependencias — ver
  `orquestador.py`) y **Abrir carpeta del caso**. Las columnas del diagrama son
  `Estructura | Estado | Acción | Origen` — ya **no hay columna de detalle**
  (el usuario la pidió sacar): a la derecha del botón queda solo el
  `Origen: ...` con su link, en las filas que traen algo de afuera del caso.
  Las cuatro columnas son celdas (`Frame`) de ancho fijo **en píxeles**
  (`ANCHO_ESTRUCTURA`, `ANCHO_ESTADO`, `ANCHO_ACCION`) con
  `pack_propagate(False)`: antes estaban en "caracteres" y una fila en
  negrita (las de nivel 0) corría el estado y el botón a la derecha respecto
  del resto. Cada botón ocupa su celda entera, así que todos miden lo mismo
  y quedan en una columna recta.

  | Fila | Botón | Qué hace |
  |---|---|---|
  | `Medidas/Medidas_SAE.xlsx` | **Actualizar** | `nucleo.generar_medidas_sae` — corre los cuatro pasos de Medidas de un viaje |
  | `Cmg/cmg<AAMM>_def_15minutal.csv` | **Traer** | `nucleo.traer_csv_cmg` — copia el CSV del período desde la unidad de red a `Cmg/` |
  | `Cmg/cmg.xlsx` | **Generar** | `nucleo.generar_cmg` — arma `cmg.xlsx` con el CSV que quedó al lado |
  | `FD y FMA/SSCC_Desempeño_*` | **Traer** | `nucleo.traer_fd` — baja el FD del período del árbol de indicadores del DCO y descomprime el zip |
  | `FD y FMA/fma_cpf_<AAMM>.xlsx` | **Generar** | `nucleo.generar_fma` con `{"cpf"}` — desde los reportes diarios del DCO |
  | `FD y FMA/fma_csf_<AAMM>.xlsx` | **Generar** | `nucleo.generar_fma` con `{"csf"}` — trae los reportes del AGC a `agcface/` y los concatena |
  | `FD y FMA/fma_cft_<AAMM>.xlsx` | **Generar** | `nucleo.generar_fma` con `{"ctf"}` — desde el `CTF_<AAAA><MM>.csv` del DCO |
  | `Subastas/DB subastas/` | **Traer** | `nucleo.traer_subastas` — copia los `OfertasSSCCAdj*.accdb` del período desde la unidad de red |
  | `Balance_BESS.xlsx` | **Actualizar** | las dos mitades de un viaje: `generar_consolidado` con todas las secciones y después `generar_pagos_bess` con todas |
  | cada `hoja '...'` de entrada | **Actualizar** | `generar_consolidado` con esa sola sección |
  | cada `hoja '...'` de cálculo | **Calcular** | `generar_pagos_bess` con esa sola |
  | `Balance_BESS.xlsx/COMPENSACION_CENTRAL` | **Resumir** | compensación por central y ciclo/ventana, y total por empresa |
  | `Balance_BESS.xlsx/PRORRATA_RETIROS` | **Traer** | reparte el monto de cada cuarto de hora segun `Prorrata 15min` |
  | `Balance_BESS.xlsx/Resumen` | **Asignar** | consolida `RECIBE`, `PAGA` y `NETO` por empresa |
  | `Control_corrida.xlsx` | (sin botón) | lo deja cada corrida: `Ejecucion`, `Alertas` y `Log` |

  Los botones dicen **solo el verbo** (pedido del usuario: *"los botones de
  resumir y traer que digan solo eso"*): de qué hoja se trata ya lo dice la
  fila en la que está el botón.

  La salida es **una sola planilla** (`ARCHIVO_SALIDA`), con las hojas
  ordenadas de fin a inicio (`ORDEN_HOJAS_SALIDA`: `Resumen` arriba,
  `Medidores` abajo). Se desglosa por hoja igual que `Centrales.xlsx`: lo que
  no se actualiza se **conserva** tal cual estaba en el archivo (no se
  recalcula ni se borra — ver `escribir_salida`/`hojas_regenerar`), y si el
  archivo todavía no existe se crea con el resto de las hojas vacías (queda
  registrado en la hoja `Log` del archivo de control y el diagrama las
  muestra como PENDIENTE). Cada mitad preserva además las hojas de la otra
  (`_preservar_ajenas`): son el mismo archivo.

  **Todo es link.** El nombre de cada archivo y de cada carpeta del
  diagrama abre en el explorador **la carpeta** de esa ruta (la que
  contiene al archivo, si la fila es un archivo: nunca se abre el
  archivo, para no arrancar Excel sin que se lo pidan). Si la ruta
  todavía no existe (una carpeta que falta, una salida sin generar) se
  sube hasta el primer ancestro que sí exista — `carpeta_a_abrir()`. La
  ruta viaja en la fila (`ruta` + `es_carpeta`, los pone
  `revisar_estructura()`); el prefijo del árbol va en su propia etiqueta
  para que el subrayado del link tape sólo el nombre.

  Y las filas de lo que viene de **afuera del caso** lo dicen en su
  detalle, con la etiqueta también como link a la carpeta de origen:

  | Fila | Detalle | Abre |
  |---|---|---|
  | `SSCC_Desempeño_*` | **Origen: DCO** | `…\Indicadores Publicar\<Vn>\03 Desempeño para publicar` — el `<Vn>` que **de verdad se usó** (V2 y, si ahí no está, V1) |
  | `cmg<AAMM>_def_15minutal.csv` | **Origen: CMg Reales** | `T:\CMgReales 15MIN\<AAAA>\<AAMM>\Mensual\CMg\Cmg para balance` |
  | `DB subastas/` | **Origen: progdiar_adjudicaSEN** | `RAIZ_SUBASTAS_ORIGEN` |
  | `cmg.xlsx` | **Origen inputs:** el CSV 15-minutal | la misma carpeta de `T:` |
  | `fma_cpf_<AAMM>.xlsx` | **Origen inputs:** los reportes diarios de CPF | `<Vn>\01 Respuesta\01 Indices CPF` |
  | `fma_csf_<AAMM>.xlsx` | **Origen inputs:** los reportes del AGC | `RAIZ_AGC_FACE` |
  | `fma_cft_<AAMM>.xlsx` | **Origen inputs:** el `CTF_AAMM.csv` | `<Vn>\01 Respuesta\06 Indices CTF` |

  Las tres de FMA dicen "Origen **inputs**" y no "Origen" a propósito: el
  FMA no se trae hecho, se **construye** — lo que viene de afuera son sus
  insumos. La ruta se resuelve **recién al hacer click**, en un hilo
  aparte (`abrir_origen()`): resolverla implica mirar el servidor (qué
  versión está publicada) y el diagrama se repinta en cada revisada, así
  que ahí sólo viaja la etiqueta.

  Todos los botones corren su función de `nucleo` en un hilo aparte
  (`lanzar()`, helper compartido) reportando al log/barra/timer de la
  ventana, y mientras algo corre **todos** los botones del árbol quedan
  deshabilitados (`corriendo`/`habilitar_botones`). Como el árbol se
  repinta entero en cada `revisar()`, las referencias a los botones se
  renuevan ahí (`botones_arbol`).
- **Consume:** `Script.nucleo` (`revisar_estructura`, `generar_medidas_sae`,
  `traer_csv_cmg`,
  `generar_cmg`, `generar_consolidado`, `generar_pagos_bess`,
  `SECCIONES_CONSOLIDADO`, `SECCIONES_PAGOS`, `validar_aamm`,
  `ErrorEntrada`, `extrae_cmg`, `carpeta_corresponde_al_periodo`);
  `config.json` (última carpeta base, último AAMM y
  `carpetas_por_periodo` —qué carpeta usó cada mes—, por PC/usuario).
- **Produce:** `config.json` actualizado con la carpeta base, el AAMM y la
  carpeta de ese período (`recordar_carpeta()`); dispara en `nucleo` la escritura del CSV de CMg, `cmg.xlsx`,
  `Balance_BESS.xlsx` (y `Control_corrida.xlsx`) dentro de la carpeta
  base del caso (cada hoja por su botón).
- **Expone:** `main()` — punto de entrada (`python Balance_BESS.py`);
  `carpeta_a_abrir(ruta, es_archivo=False)` y
  `abrir_en_explorador(ruta, es_archivo=False)` (la carpeta que abre cada
  link); helpers de presentación del árbol (`_es_ultimo_en_su_nivel`,
  `_prefijos_arbol`) que traducen la lista plana de `revisar_estructura()`
  a prefijos tipo consola — deliberadamente NO viven en `nucleo/`, que no
  conoce conceptos de interfaz. El **nivel** de cada fila sí lo pone
  `nucleo` (es estructura, no dibujo), y el **id** de cada fila es lo que
  la ventana usa para decidir qué botón le cuelga (`_boton_de_fila`): así
  `nucleo/` no sabe nada de botones.

- **Un período, una carpeta.** Regla del usuario: *"dos meses no pueden
  tener la misma carpeta"*. Al cambiar el AAMM (`aamm_cambiado()`):
  1. Si ese período ya tuvo carpeta (`carpeta_recordada()`), se vuelve a
     ella sola.
  2. Si no, se pregunta a `nucleo.carpeta_corresponde_al_periodo()`, que
     responde **tres** cosas, no dos: sí, no, o **no se sabe** (el nombre
     no trae ningún AAMM y nunca se anotó). Antes ese "no se sabe" se
     trataba como un sí y la ventana se quedaba trabajando sobre la
     carpeta del mes anterior sin decir nada — el bug que reportó el
     usuario.
  3. Si no es la de este período, `ventana_carpeta_del_periodo()`:
     **Examinar…** / **Crear la carpeta** / (solo en el "no se sabe")
     **Esta es la de AAMM** / **Cancelar**, que vuelve al período
     anterior. Elegir una carpeta la deja anotada
     (`recordar_carpeta()`), así que el mes siguiente ya no la puede
     reusar.

  Detalles de tkinter que importan: la ventana del período es modal
  (`grab_set` + `wait_window`) y `ventana_nuevo_caso()` también, porque
  hay que **esperar** su resultado (si se cancela, el período sigue sin
  carpeta). Y mientras se pregunta, el campo del AAMM pierde el foco y
  `<FocusOut>` vuelve a llamar a `aamm_cambiado()`: la marca
  `preguntando` es lo que evita que se abran dos ventanas encima.
- **Depende de:** el paquete `Script/` (mismo directorio).

---

## `Script/config.py`

- **Qué hace:** es el único lugar del proyecto que lee y escribe
  `config.json` (que vive en la raíz, junto a `Balance_BESS.py`, y no se
  versiona).

- **Dos clases de sección, y la diferencia importa:**
  - `"<hostname>_<usuario>"` — lo de cada PC/usuario: la última carpeta base,
    el AAMM y `carpetas_por_periodo` (`{aamm: carpeta}`, la carpeta que usó
    cada mes). Las escribe la ventana sola.
  - `"claves_api"` — **compartida**: las dos claves de las APIs del
    Coordinador. Nombre reservado (una sección de usuario nunca se llama
    así). El valor es el mismo para todo el equipo, pero como el archivo es
    local, cada uno lo pega una vez en su copia.

- **Expone:** `leer_todo()`, `escribir_todo()`, `seccion(nombre)`,
  `actualizar_seccion(nombre, datos)` — que mezcla sin pisar el resto del
  archivo — y `clave_api(cual)`, con las constantes `CLAVE_PRMTE` y
  `CLAVE_GENERACION_REAL`.

- **`clave_api()`** levanta `ErrorConfig` con la ruta del archivo y el JSON
  exacto para pegar cuando la clave falta, cuando el archivo no existe, o
  cuando quedó el `PEGAR_AQUI_LA_CLAVE` del ejemplo sin reemplazar. Se llama
  **en el momento de usar la clave**, no al importar: así se puede completar
  el `config.json` con el programa ya abierto.

- **Las dos claves son distintas** (`prmte` → `medidas.coordinador.cl`,
  `generacion_real` → `operacion.coordinador.cl`). Antes había una sola
  constante `USER_KEY` en el código, lo que además de dejarla versionada
  forzaba a que las dos APIs compartieran clave.

- **Quién lo usa:** `Balance_BESS.py` (su propia sección de usuario) y
  `Script/Medidas/comun.py` (`leer_clave_api()`, que traduce `ErrorConfig` a
  `ErrorMedidas` para que la ventana lo muestre igual que el resto).

---

## `Script/Medidas/` — cómo se arma `Medidas_SAE.xlsx`

- **Qué hace:** los cuatro pasos que antes eran cuatro scripts sueltos que se
  corrían a mano uno detrás de otro (autor original: Freddy.Arriagada).
  Ahora corren de un viaje desde el botón **Actualizar** de la fila
  `Medidas_SAE.xlsx` (`nucleo.generar_medidas_sae`).

  | Módulo | Script original | Qué hace |
  |---|---|---|
  | `Homologacion.py` | `0_diccionario_prmte_a_claves_balance.py` | lee las dos hojas del Excel de homologación: `homol` (`Punto de Medida` + `Canal` → `clave` + `Flujo`) y `Gen real` (ahí `Canal` es la unidad — se lee su principio: `MWhD`/`MWhR`/`kWhD`/`kWhR`) |
  | `Descarga_PRMTE.py` | `1_generacion_prmte.py` | baja las medidas de cada punto, por lotes, reanudable |
  | `Claves_Balance.py` | `2_generacion_claves_Balance.py` | calendario de cuartos de hora + agrupación por clave |
  | `Generacion_Real.py` | `3_Generacion_Real.py` | agrega las centrales de la hoja `Gen real` desde la API de operación real |

- **Qué cambió respecto de los scripts sueltos** (todo a pedido del usuario,
  salvo donde se diga):
  - el Excel de homologación vive en `Auxiliares/`, al lado de
    `Centrales.xlsx`, y se busca por patrón (`*homologacion*`) en vez de
    estar al lado del `.py`. El `homol.parquet` intermedio desapareció: se
    lee una vez y queda en memoria;
  - la lista `FILTROS_TOPOLOGY` que vivía dentro del paso 3 salió del código
    y ahora es la hoja **`Gen real`** del mismo Excel de homologación (primero
    se hizo como hoja de `Centrales.xlsx`; el usuario la movió acá: es
    homologación, igual que `homol`, y se mantiene con el mismo archivo).
    **Ya no es un reemplazo**: esas centrales se sacaron de `homol`, así que
    no llegan por el otro camino — el paso 3 las **agrega**, con la `clave`
    que diga esa hoja. Mismas cuatro columnas que `homol`, con el
    `topologyName` de la API en la columna `Punto de Medida` (esa API no
    tiene puntos de medida) y `Canal` = la **unidad** de la medida
    (`MWh`/`kWh`). Esa columna estaba sin uso —la API no expone canales— y
    ahora decide el factor de conversión: **la API devuelve MWh y todo el
    balance trabaja en kWh** (`UNIDADES_GEN_REAL` en `Homologacion.py`, se
    aplica en `expandir_a_cuartos()`). Se mira **sólo el principio** del
    texto (`unidad_desde_canal()`): en el archivo real el canal viene
    `MWhD`/`MWhR` —la unidad con el tipo de medida pegado atrás— y lo único
    que decide es la unidad. Vacía, o con un texto que no empieza con
    ninguna de las dos, = MWh. Sin
    esa conversión estas centrales entraban mil veces más chicas que las
    que vienen por punto de medida: se detectó comparando, en un caso
    real, Andes Solar III (Pmax 170,78 MW, máximo 44 por cuarto de hora)
    contra Tocopilla (116 MW, máximo 29.493);
  - el `user_key` de las dos APIs sale de una sola constante
    (`comun.USER_KEY`). Sigue en el código, a pedido explícito del usuario,
    pero deja de estar repetido en dos archivos y con valores distintos;
  - los lotes descargados y la marca de reanudación van a
    `<CARPETA_BASE>/Medidas/_trabajo/`, que la ventana no muestra;
  - **(no pedido, es un bug)** los lotes y la marca de reanudación llevan el
    período en el nombre. Antes, correr dos meses en la misma carpeta
    mezclaba los lotes (`medidas_batch_*.parquet` los levantaba todos) y
    daba por procesados puntos de otro mes;
  - **(no pedido, es un bug)** el umbral de "punto de medida completo" ya no
    es la constante `2976` (= 31 × 96) sino la cantidad de cuartos de hora
    que el mes descargado realmente trae. `2976` estaba mal para cualquier
    mes de 30 días o menos, y para los meses con cambio de hora;
  - se dejó de generar `log_inconsistencias_medidas.xlsx`, que comparaba el
    criterio de desempate viejo contra el de mayor `idMeasure`: esa
    comparación era una investigación ya cerrada (el propio script la titula
    "CRITERIO DEFINITIVO"). Lo que sí se informa en el log es cuántos grupos
    venían duplicados.
- **Detalle que importa:** el `Cuarto de Hora` es un índice global del mes
  que después cruza contra `CMg`, así que las **dos** fuentes comparten un
  solo calendario, el que arma `Claves_Balance` (numerado por `intervaloUtc`,
  porque la hora local se repite en el cambio de hora y desordenaría la
  numeración). Las filas de la API de operación real se le pegan por hora
  local — esa API no entrega UTC —, así que en un día de cambio de hora hacia
  atrás hay ambigüedad: se toma la primera ocurrencia y se avisa en el log.
- **Depende de:** `pandas`, `requests` y `pyarrow` (los lotes son parquet).
  **No importa `nucleo`** (misma regla que `Cmg/`): recibe rutas y datos y
  levanta `ErrorMedidas`, que `nucleo` traduce a `ErrorEntrada`.

---

## `Script/Subastas/Ofertas_Adjudicadas.py`

- **Qué hace:** todo lo que sabe de los Access de subastas, que son el
  **origen real** de la hoja `Subastas`. Antes las subastas se leían de la
  hoja `DB` de la planilla 3 (`3_REMUNERACIÓN_SUBASTAS_E_ID_*`, que desde
  esta sesión ya no se usa ni aparece en la ventana), pero esa planilla
  nunca fue el origen:
  ella misma se arma pegando la salida de `entradas_sscc.py` (script suelto,
  autor original Gerardo.Vieyra), que lee los `OfertasSSCCAdj*.accdb`. Este
  módulo replica esa rutina (`calc_subastas`) con tres cambios pedidos por el
  usuario: los `.accdb` se copian primero a la carpeta del caso y de ahí se
  leen (mismo patrón que el CSV de CMg), el período sale del AAMM de la
  ventana y no de un `.yaml`, y no escribe ningún Excel (devuelve DataFrames).
- **Consume:**
  `\\nas-cen1\Estadisticas\progdiar_adjudicaSEN\OfertasSSCCAdj<AAAAMMDD>[_HH].accdb`
  (`RAIZ_SUBASTAS_ORIGEN` + `PREFIJO_ACCDB` — sale de `path_subastas_origen`
  y `archivo_subasta_prefix` del `archivo_de_configuracion.yaml` del script
  original; es la segunda y última ruta del programa que apunta fuera de la
  carpeta base del caso), y después esos mismos archivos ya copiados en
  `<CARPETA_BASE>/Subastas/DB subastas/`.
- **Produce:** las copias locales de los `.accdb`; el DataFrame crudo
  equivalente a `subastas_AAMM.xlsx` (la salida del script viejo), que
  `nucleo.construir_subastas_desde_accdb()` transforma en la hoja `Subastas`.
- **Expone:** `ErrorSubastas`; `CARPETA_DB_SUBASTAS`; `nombre_accdb(aamm, dia,
  hora)`, `nombres_del_periodo(aamm)`, `ruta_origen(raiz=None)`,
  `carpeta_db_subastas(...)`, `asegurar_carpeta_db(...)`,
  `accdb_presentes(carpeta_subastas, aamm)`,
  `traer_accdb(carpeta_subastas, aamm, raiz=None, registrar=print)` →
  `(copiados, salteados, faltantes)`,
  `leer_accdb(ruta)`, `construir_crudo(carpeta_subastas, aamm, registrar)` →
  `(df_crudo, resumen)`.
- **Depende de:** `pandas` y, solo al leer, `pyodbc` + el *Microsoft Access
  Database Engine*. **No importa `nucleo`** (misma regla que `Cmg/` y
  `Medidas/`). `pyodbc` se importa **dentro** de las funciones que leen, no
  arriba: es la única dependencia que solo existe en Windows y no tiene por
  qué romper el import de todo el programa donde no esté.
- **Detalle que importa:** por cada día se parte del Access del PO (hora 0,
  sin sufijo) y cada PID de la hora `HH` **reemplaza las horas >= HH** — las
  anteriores quedan como estaban. De qué archivo salió cada fila queda en
  `Hora_PID`. Después se descartan las filas con `CANTIDAD MW = 0`, el precio
  se pone en 0 fuera de la banda 1, `CANTIDAD PONDERADA MW` vacía se completa
  con `CANTIDAD MW` y se eliminan duplicados: exactamente lo que hacía
  `entradas_sscc.py`, verificado con un caso sintético.

---

## `Script/Fd/Indicadores_DCO.py`

- **Qué hace:** sabe llegar a la carpeta donde el DCO publica los indicadores del
  mes y traer de ahí el FD (botón **"Traer FD"**). Las dos piezas de la ruta
  salen del `archivo_de_configuracion.yaml` de `entradas_sscc.py`
  (`ruta_fma_dco` y el nombre `SSCC_Disponibilidad_CSF_<Mes>_<AAAA>_<V>.zip`), y
  la forma completa la muestra el comentario de la rutina de FMA CPF de ese
  mismo script.
- **Consume:**
  `F:\11 SSCC\05 Verificación SSCC\02 Cálculo indicadores\<AAAA>\<MM>. <Mes>\Indicadores Publicar\<V1|V2>\03 Desempeño para publicar`
  (`RAIZ_DCO_INDICADORES` + `SUBCARPETAS_FD` — ruta confirmada por el usuario).
  **Esa carpeta y ninguna otra**: no se busca en la que se usaba antes
  (`04 Desempeño para transferencias`, hoy sólo `SUBCARPETAS_FD_ANTIGUA`, para
  nombrarla en el error) ni en otras ramas del árbol de la versión, y tampoco
  dentro de sus subcarpetas. La búsqueda recursiva que había como último
  recurso es lo que hacía que "Traer FD" copiara archivos de media publicación.
- **Produce:** la copia del FD dentro de `<CARPETA_BASE>/FD y FMA/`, y los Excel
  que venían dentro del `.zip`, sueltos en esa misma carpeta (que es donde
  `buscar_archivo_sscc_desempeno()` los busca después).
- **Expone:** `ErrorFd`; `carpeta_del_periodo(aamm, raiz=None)`,
  `versiones_publicadas(...)`, `elegir_version(..., version=None)`,
  `buscar_en_versiones(carpeta_publicacion, buscar, version=None, registrar)` →
  `(carpeta_version, resultado, revisadas)` — recorre las versiones de mayor a
  menor hasta que `buscar` encuentre algo,
  `carpeta_fd_de_la_version(carpeta_version)`,
  `buscar_archivos_fd(carpeta_version, anio)` (sólo los sueltos de esa carpeta),
  `traer_fd(carpeta_destino, aamm, version=None, raiz=None, registrar=print)` →
  `(copiados, extraidos, carpeta_version)`; y, para el link "Origen: DCO" de
  la ventana, `carpeta_publicacion_o_literal(aamm, raiz=None)`,
  `carpeta_version_usada(aamm, buscar=None, version=None, raiz=None)`,
  `ruta_origen(aamm, subcarpetas=None, version=None, raiz=None, buscar=None)` y
  `ruta_origen_fd(aamm, version=None, raiz=None)` — estas cuatro **no
  levantan**: con la unidad desconectada devuelven igual la ruta que le
  correspondería al período. Con `buscar` (la misma función que usa el botón)
  apuntan a la versión que **de verdad tiene el archivo**, no a la más alta que
  exista: el link de la ventana y la carpeta de la que se copió son la misma.
- **Depende de:** solo la biblioteca estándar. **No importa `nucleo`.**
- **Decisiones que se tomaron acá** (no venían dadas):
  - **qué versión usar**: la ventana no tiene selector Pre/Def, así que por
    omisión se toma la **más alta que de verdad tenga el archivo**
    (`buscar_en_versiones()`), no la más alta a secas: puede existir la carpeta
    `V2` y no tener adentro lo que se busca (recién creada, a medio subir), y ahí
    hay que caer a `V1`. El log dice cuál usó y en cuáles no estaba. Se puede
    forzar con `version`.
  - las carpetas del año/mes se buscan **comparando por nombre normalizado**
    (sin tildes, sin importar mayúsculas), no con una ruta literal: las escribe
    una persona todos los meses y `"03. Marzo"` y `"3. Marzo"` son la misma.
  - la búsqueda del archivo es **recursiva** dentro de la carpeta de versión,
    porque el DCO cambia de subcarpeta de un mes a otro; se filtra por prefijo
    **y** por el año en el nombre, para que no se cuele un archivo de otro
    período que haya quedado suelto ahí.
  - al descomprimir se sacan **solo los Excel** y se dejan sueltos en la
    carpeta; las entradas del zip con ruta absoluta o con `..` se ignoran.

---

## `Script/Fd/Indices_FMA.py`

- **Qué hace:** arma las tres salidas de FMA del período, **una por botón
  "Generar"**, replicando `calc_fmacpf`, `calc_fmacsf` y `calc_fmactf` de
  `entradas_sscc.py`. **Ojo con la palabra "traer"**: el FMA no se copia ya hecho
  de ningún lado, se **construye**, y cada una sale de un origen distinto.
- **Consume** (las tres rutas las confirmó el usuario):
  - **CPF**: `<versión>/01 Respuesta/01 Indices CPF/20AA.MM_Respuesta_CPF/Reporte diario <D>-<M>-<AAAA>/tabla_resumen_<D>_<M>_<AAAA>.xlsx`
    — dentro del mismo árbol del DCO del que sale el FD. Se lee **cada hoja menos
    "Resumen"** (una por central).
  - **CSF**: `\\nas-cen1\D. Transferencias\SCADA\reporte_agc_face_NM10`
    (`RAIZ_AGC_FACE`), donde están **todos los meses juntos**: se eligen los del
    período por su nombre (`csf_<AAAAMMDD>`), se copian a
    `<CARPETA_BASE>/FD y FMA/agcface/` y se concatenan.
  - **CTF**: `<versión>/01 Respuesta/06 Indices CTF/CTF_<AAAA><MM>.csv` — otra
    rama del mismo árbol del DCO. Se copia a la carpeta del caso antes de usarlo,
    para que quede registrado con qué archivo se armó la salida.
- **Produce:** `fma_cpf_<AAMM>.xlsx`, `fma_csf_<AAMM>.xlsx` y
  `fma_cft_<AAMM>.xlsx` (+ su `.csv`) en `<CARPETA_BASE>/FD y FMA/` — con los
  nombres exactos de `entradas_sscc.py` ("cft" incluido), que son los que después
  busca `Script/Subastas/Fma.py`. Las dos mitades están probadas juntas: lo que
  escribe este módulo lo lee aquel sin tocar nada.
- **Expone:** `ErrorIndicesFma`; `nombre_salida(tipo, aamm)`,
  `buscar_carpeta_respuesta_cpf(...)`, `buscar_tabla_resumen(...)`,
  `construir_fma_cpf/csf/ctf(...)`, `traer_agc_face(...)`, `TIPOS_FMA`,
  `buscar_reportes_cpf(...)`, `buscar_ctf(...)`,
  `generar_fma(carpeta_destino, aamm, tipos=None, version=None, raiz=None, raiz_agc=None, registrar=print)`
  → `(escritos, faltantes, versiones)`; y, para el "Origen inputs: ..." de la
  ventana, `ruta_origen_cpf(aamm, ...)`, `ruta_origen_csf(aamm=None, ...)`,
  `ruta_origen_ctf(aamm, ...)` y el dict `RUTAS_ORIGEN_FMA` — que devuelven la
  carpeta de los insumos sin levantar aunque la unidad esté desconectada. En
  `versiones` cada tipo elige la suya por separado:
  **cada uno elige su versión por separado**, porque una puede tener el CPF y
  otra el CTF.
- **Depende de:** `pandas` y `Script/Fd/Indicadores_DCO.py` (comparte con él la
  resolución del árbol del DCO y la elección de versión). **No importa `nucleo`.**
- **Detalles que importan:**
  - el reporte diario de CPF trae **29 columnas** y el encabezado en la fila 5
    (`header=4`); el script les pone 34 nombres (`NOMBRES_CPF`) contando las 5
    que agrega. Si el DCO le cambia el formato, se levanta un error que lo dice
    en vez de correr las columnas en silencio.
  - la **hora** del CPF sale del índice de fila del reporte + 1 (queda 1..24, ya
    alineada con las subastas).
  - los `"-"` de los dos `Tiempo f…` se reemplazan por 0.
  - al CTF se le saca la zona horaria **conservando la hora tal como está
    escrita** (`_sacar_zona_horaria`): `04:00-03:00` queda en las 04:00, **no**
    en las 07:00. Convertir a UTC corría todas las horas del CTF — se detectó
    justamente en la prueba.
  - un día sin archivo se saltea con aviso (el original revienta), y si falta el
    origen de una, las otras se arman igual.
  - **el CSF no necesita el DCO publicado** (su origen es otro servidor), así que
    su botón funciona aunque el mes todavía no tenga indicadores publicados.
  - los reportes del AGC se llaman **`csf_<AAAAMMDD>`** (confirmado por el
    usuario, ej. `csf_20260301.xlsx` — el mismo nombre que espera el script
    original). Se acepta un sufijo después de la fecha, pero **no** un archivo
    que solo la contenga en el medio del nombre.

---

## `Script/Fd/Desempeno_Horario.py`

- **Qué hace:** saca del `SSCC_Desempeño_*` las **dos** cosas que dependían de
  él: la columna **`Subastas!FD`** (antes `DB!Y`) y el **Vector de Participación
  CSF** (antes `DB!AC`), que multiplica al FMA de las filas CSF y era lo único
  que faltaba para cerrar `Subastas!FMA`.
- **Consume:** las tres hojas horarias de `<CARPETA_BASE>/FD y FMA/SSCC_Desempeño_*`
  — `CPF Horario` (B:J, el FD es `I`), `CSF Horario` (B:H, el FD es `H`) y
  `CTF Horario` (B:I, el FD es `I`); encabezados en la fila 11, datos desde la 12
  (mismo criterio que `nucleo.construir_fd`, que lee estas mismas hojas para la
  hoja `FD` del consolidado).
- **Produce:** un dict con tres diccionarios `(unidad, hora_mes) → FD` más
  `participacion_csf` con `(unidad, hora_mes) → 0|1`.
- **Expone:** `ErrorDesempeno`; `unidad_alternativa(unidad)`,
  `calcular_hora_mes(...)`, `leer_hoja(ruta, control)`,
  `construir_tablas_fd(ruta_sscc, dia_cambio_hora=None, registrar)`,
  `buscar_fd(tablas, control, unidad, hora_mes)`,
  `buscar_participacion_csf(tablas, unidad, hora_mes)`.
- **Depende de:** solo `pandas`. **No importa `nucleo`.**
- **Tres detalles que no son obvios** (salen del documento de trazabilidad de FD):
  1. la `Hora` de estas hojas va de **0 a 23**, así que
     `Hora_Mes = (día - 1) × 24 + hora + 1` — la misma escala 1..24 por día que
     usa `Subastas`, que es lo que permite cruzarlas;
  2. **CPF prueba una segunda nomenclatura**: si no encuentra la unidad,
     intercambia el sufijo `TG` ↔ `TV` y busca de nuevo. CSF y CTF buscan una
     sola vez (así es la fórmula original);
  3. el **Indicador de Participación CSF** es 0 solo si la unidad figura como
     `"No Participó"` **y** su alternativa TG/TV tampoco participó.
- **El FD no se recalcula** a partir de las respuestas: se toma tal cual viene en
  el archivo (§25 del documento — puede venir `Respuesta = "No Participó"` con
  `FD = 1`, y así queda).

---

## `Script/Subastas/Fma.py`

- **Qué hace:** normaliza las tres salidas de FMA de `entradas_sscc.py`
  (`fma_cpf_<AAMM>.xlsx`, `fma_csf_<AAMM>.xlsx`, `fma_cft_<AAMM>.xlsx` o `.csv` —
  "cft" está así en el script original, se aceptan las dos grafías) para que
  `nucleo` pueda calcular `Subastas!FMA`. En la planilla 3 esas tres salidas se
  pegaban en las hojas `FMA_CPF`/`FMA_CSF`/`FMA_CTF` y la columna `DB!V` las
  buscaba con `BUSCARV`; acá se hace el mismo cálculo sin la planilla.
- **Consume:** los tres archivos, desde `<CARPETA_BASE>/FD y FMA/` (el usuario le
  cambió el nombre a esa carpeta, antes `SSCC_Desempeño/`, justamente para que
  guarde las dos cosas). Se buscan por patrón (prefijo + el AAMM en el nombre),
  no por nombre literal.
- **Produce:** tres tablas normalizadas: CPF
  (`Central, Año, Mes, Dia, Hora, FMA_CPF_mas, FMA_CPF_menos`), CSF
  (`Año, Mes, Dia, Hora, FMA_CSF_mas_base, FMA_CSF_menos_base`) y CTF (dos
  resúmenes, uno por `unidad` y otro por `Configuracion`).
- **Expone:** `ErrorFma`; `buscar_archivos_fma(carpeta, aamm)`,
  `tabla_cpf(ruta)`, `tabla_csf(ruta)`, `tablas_ctf(ruta)`.
- **Depende de:** solo `pandas`. **No importa `nucleo`.**
- **Tres detalles que no son obvios** (salen del documento de trazabilidad del
  usuario, no de una suposición):
  1. **CPF:** se suman las 7 columnas de horas y, **si la suma llega a 0,98, pasa
     a ser 1**. `FMA CPF(+) = suma × Tiempo f<49.975` y
     `FMA CPF(-) = suma × Tiempo f>50.025`. Esos dos tiempos **ya vienen como
     fracción** (0,33 = 33 %): no se vuelven a dividir por 100.
  2. **CSF:** la hora del archivo va de **0 a 23** y la de las subastas de **1 a
     24**, así que se le suma 1. Y se usan las columnas **`-m`/`+m`**
     (`FMA CSF-m [%]`, `FMA CSF+m [%]`) divididas por 100, **no** las que se
     llaman igual sin la "m".
  3. **CTF:** la duración de cada activación es `(tfin - t0)` en horas, se reparte
     según el signo de `variacion`, y las activaciones **que caen en la misma hora
     se suman**. Se arman dos resúmenes porque la fórmula original busca primero
     por una nomenclatura y después por la otra.

---

## `Script/Cmg/Extrae_CMG_barras.py`

- **Qué hace:** todo lo que sabe del CSV 15-minutal de CMg. Viene del
  script suelto que se corría a mano al lado del CSV (autor original:
  Freddy.Arriagada), con tres cambios: el CSV se baja de la unidad de red a
  la carpeta `Cmg/` del caso en vez de buscarse al lado del `.py`, las
  barras a filtrar se reciben por parámetro en vez de estar escritas en el
  código, y no escribe el Excel (devuelve DataFrames).
- **Consume:**
  `T:\CMgReales 15MIN\<AAAA>\<AAMM>\Mensual\CMg\Cmg para balance\cmg<AAMM>_def_15minutal.csv`
  (`RAIZ_CMG_REALES` + `SUBCARPETAS_CMG_REALES` + `PLANTILLA_CSV_CMG_15MIN`
  — la única ruta del programa que apunta fuera de la carpeta base del
  caso; si `T:` cambia de letra se cambia ahí y nada más), y después el
  mismo CSV ya copiado en `<CARPETA_BASE>/Cmg/`.
- **Produce:** la copia local del CSV; el DataFrame de `cmg.xlsx` (lo
  escribe `nucleo.generar_cmg`).
- **Expone:** `ErrorCmg`; `nombre_csv_15min(aamm)`,
  `ruta_csv_en_red(aamm, raiz=None)`, `carpeta_origen_csv(aamm, raiz=None)`
  (la carpeta de red, para el link "Origen: CMg Reales" de la ventana),
  `ruta_csv_local(carpeta_cmg, aamm)`,
  `traer_csv_15min(carpeta_cmg, aamm, raiz=None, registrar=print)`,
  `construir_cmg_desde_csv(ruta_csv, barras, registrar=print)` →
  `(df_salida, resumen_dias)`, `validar_layout(df, registrar=print)`,
  `resumen_dias_anomalos(resumen_dias)`.
- **Depende de:** solo `pandas` (a propósito: **no importa `nucleo`**, así
  no hay ciclos de import cuando se saquen más etapas a módulos propios).
  Los errores previsibles salen como `ErrorCmg` y `nucleo` los traduce a
  `ErrorEntrada`.
- **Detalle que importa:** `nucleo.leer_cmg()` vuelve a leer `cmg.xlsx`
  **por posición** (D = Barra, F = valor de Q, H = Cuarto de Hora,
  I = CMg promedio), así que un cambio de columnas en el CSV rompería la
  etapa siguiente en silencio: `validar_layout()` avisa en el log si
  `BARRA` deja de caer en D o `Cuarto de Hora` en H. El `Cuarto de Hora`
  global se numera con los bloques que el CSV **realmente** trae (no se
  asumen 96 por día: los días de cambio de hora traen 92/100).

---

## `nucleo/` (paquete)

- **Qué hace:** todo el cálculo de la etapa Medidores (Medidores, Ofertas
  SSCC), la carga de CMg, FD y Subastas, y una primera etapa (base) de
  Calculo E Costos, sin interfaz. Resuelve las rutas de un caso a partir de
  la carpeta base, valida que existan las entradas requeridas (incluido el
  período AAMM que ingresa el usuario), y escribe `Balance_BESS.xlsx` (una
  sola planilla: `Resumen`, `PRORRATA_RETIROS`, `COMPENSACION_CENTRAL`,
  `Calculo RE545`, `Calculo E Costos`, `Subastas`, `FD`, `CMg`,
  `Ofertas SSCC`, `Medidores`) y `Control_corrida.xlsx` (`Ejecucion`,
  `Alertas`, `Log`).

  **Medidores**: calculadas J (SoC), L (Ventana), O (Indicador_SoC).

  **Los auxiliares ya no se escriben** (pedido del usuario: "quita los
  auxiliares innecesarios", ver `COLUMNAS_AUXILIARES_MEDIDORES`). La hoja
  sale con 12 columnas — `Mes, Dia, Hora, Minutos, Hora Mes, Cuarto de
  Hora, clave, intervalo, Gen_Unidad, SoC, Ventana, Indicador_SoC`— en el
  orden de siempre (el de `LETRA_A_CAMPO`) sin:

  | Fuera de la hoja | Qué era | Cómo se recupera |
  |---|---|---|
  | M, P, Q, U | las cuatro columnas deliberadamente vacías del original | no hacen falta: en Python no hay letra de Excel que alinear |
  | N (`Clave_Dia_HoraMes`) | clave auxiliar `Dia & Hora Mes` | no la lee nadie; `calcular_clave_auxiliar()` sigue estando |
  | K (`Copia_Ventana`) | copia fila a fila de L (`Ventana`) | `reponer_auxiliares_medidores()` la repone al leer la hoja, que es lo que consumen las dos hojas de cálculo como "Ciclo de Carga del mes" |

  **R, S y T ya NO están en `Medidores`** (pedido del usuario:
  "independizar Medidas de ofertas"). Las tres salen de Ofertas SSCC y
  viven, derivadas, donde corresponde:

  | Columna original | De dónde sale ahora |
  |---|---|
  | R `Oferta_Completa_Dia` | tabla "Ofertas SSCC por dia" (central × día) |
  | S `Indicador_Ventana_Oferta` | R + la Ventana de Medidores |
  | T `Ventana_No_Completa` | tabla "Resumen ventana oferta" (central × ventana) |

  Las dos tablas son las que ya se escribían en la hoja `Ofertas SSCC`
  (`V, W, X, Y, AB, AC, AD, AE` del plan **no son columnas de
  `Medidores`**: son tablas auxiliares de otro largo que se escriben juntas
  en esa hoja — ver `_escribir_tabla_con_titulo()`, plan §20.1 y §22).
  `completar_ofertas_en_medidores()` reconstruye las tres columnas en
  memoria cuando hacen falta (`generar_pagos_bess()`: T es la que reparte
  cada fila entre `Calculo E Costos` y `Calculo RE545`), leyendo la hoja
  `Ofertas SSCC` con `leer_ofertas_sscc_consolidado()`. Consecuencia
  práctica: el botón **Actualizar** de `Medidores` ya no abre el archivo
  `*OfertasSSCC*` ni lo exige. El resumen intermedio equivalente a la hoja
  "Resumen Ofertas SSCC" del `.xlsm` original sigue siendo puramente
  auxiliar: no se persiste.

  **Hoja `Diccionario` de `Centrales.xlsx`, dos formatos** (ver
  `Script/nucleo/lectura.py`): el viejo son varias tablas lado a lado
  separadas por columnas vacías (`FD` en A:B, `Subastas`/`ofertas` en
  E:F:G) y no tiene dónde poner la nomenclatura de FMA CPF; el nuevo
  (propuesta del usuario) es UNA tabla con encabezados
  `Balance_BESS | FD | Subastas | Ofertas | FMA_CPF`, con el nombre
  canónico en la primera columna. Se aceptan los dos:
  `encabezado_diccionario()` detecta cuál es y `filas_diccionario()` /
  `mapa_diccionario()` dan acceso por nombre de columna. Con el formato
  nuevo, las homologaciones de FD y de FMA CPF se buscan por el nombre de
  **Subastas** (`BAT_TOCOPILLA`, no `SAE-TOCOPILLA`), que es lo que arregla
  las diferencias de FMA en CPF y las de CSF que quedaban en cero.

  **Subastas desde su origen real** (esta sesión): `Subastas` ya no sale de
  la planilla 3 sino de los Access. `construir_subastas_desde_accdb()` arma
  las MISMAS columnas `B:Q` a partir del DataFrame crudo que devuelve
  `Script/Subastas/`: `Concepto` = `SERVICIO` tal cual; `Control` =
  `SERVICIO[:3]` y `Sub_Baj` = el signo del penúltimo carácter (las dos
  fórmulas reales de la planilla, `=LEFT(C,3)` e
  `=IF(LEFT(RIGHT(C,2),1)="+","SUBIDA","BAJADA")`); `Fecha` =
  `DATE(Año,Mes,Dia)`; `Hora_dia` = `HORA` sin corrección; `Hora_mes` =
  `(Dia-1)*24 + Hora_dia` (`calcular_hora_mes_subastas()`, con el ajuste de
  cambio de hora de la planilla como parámetro opcional, hoy apagado —
  pendiente); `Propietario` sale de la columna nueva de
  `Centrales.xlsx`/`Resumen BESS` (`construir_mapa_propietario()`);
  `Energía SSCC` = `CANTIDAD PONDERADA MW`
  (`COLUMNA_ENERGIA_SSCC_ACCDB`, a confirmar). `FD` (`calcular_fd_subastas()`) y
  `FMA` (`calcular_fma_subastas()`) **no existen en el Access** —venían pegadas
  en `DB!Y`/`DB!V`— y se calculan desde la carpeta `FD y FMA/`: el FD del
  `SSCC_Desempeño_*` y el FMA de las tres salidas `fma_*`, con el Vector de
  Participación CSF que también sale del `SSCC_Desempeño_*`. La única que
  queda vacía es `Ciclo`, que se calcula después en `Calculo E Costos`.
  La hoja **no se ordena**: sale en el orden del origen. (Se ordenó por
  `Hora_mes` durante dos sesiones, a pedido del usuario, y el usuario pidió
  deshacerlo. `_ordenar_subastas_por_hora_mes()` quedó como el único lugar
  donde se decide el orden, hoy un `reset_index` y nada más: ninguna
  columna de más abajo depende del orden de las filas, todas cruzan por
  clave.)

  **CMg**, **FD**, **Subastas** (plan §23): replican únicamente las macros
  de *carga* (`Cargar_CMg_Desde_Archivo`, `Cargar_SSCC_Desempeno_En_FD`;
  `Cargar_Remuneracion_Subastas_Rapido` se eliminó junto con la planilla 3),
  no las que las consumen después
  (`Asignar_CMg_a_Calculos_Turbo`, `Actualizar_Calculos_Columnas`), que
  pertenecen a una etapa posterior sin implementar. `FD` tiene el mismo
  patrón de "dos tablas de distinto largo compartiendo hoja" que Ofertas
  SSCC, pero por **columnas** en vez de por filas: el bloque CSF (A:M) y el
  CPF (Q:AE) van lado a lado, cada uno con su propio número de filas.
  `Subastas!N` ("Energía SSCC") queda vacía y documentada como pendiente
  (depende de `'Calculo E Costos'`, cuyo resto de columnas todavía no se
  implementa). Los nombres de columna de `FD` y `Subastas` (`NOMBRES_FD_CSF`,
  `NOMBRES_FD_CPF`, `NOMBRES_SUBASTAS`) fueron confirmados por el usuario
  contra un caso real (plan §24), no inventados — antes de eso se usaba la
  letra de Excel tal cual por no tener esa información.

  **Calculo E Costos** (plan §25, etapa base — a pedido del usuario, "por
  etapas: primero H + CMg + traspaso de Medidores"): replica parcialmente
  `Traspasar_Medidores_A_Calculos_Rapido` (módulo `B_medidores_a_calculos`)
  y `Asignar_CMg_a_Calculos_Turbo` (módulo `A_Carga_Cmg_a_Destino`). Cubre
  A:G (con D↔E invertidas, igual que la macro), H/`Barra` (antes fórmula
  `=VLOOKUP(G,Resumen!B:G,6,FALSE)`; acá homologada por **nombre** de
  columna contra `Resumen BESS!Nombre activo`/`Barra inyección`, no por
  posición, porque `Centrales.xlsx` no reproduce el layout `Resumen!B:G`
  del libro original), I/J (`Energia_Positiva`/`Energia_Negativa`, la
  energía de `Medidores!Gen_Unidad` separada por signo, solo si
  `Ventana_No_Completa = 1`; si no, la fila es de `Calculo RE545`, fuera de
  alcance), K/`SoC` (copia de `Medidores!SoC`), P/`Copia_Ventana` (copia de
  `Medidores!Copia_Ventana`) y Q/`CMg` (homologado por `Barra` + `Cuarto de
  Hora` normalizado, vía `NormalizaCuarto`). Vivió un tiempo en un archivo
  separado (`Pagos_BESS.xlsx`); hoy es una hoja más de la única planilla
  `Balance_BESS.xlsx` (el usuario pidió volver a combinarlas).

  **Calculo E Costos, etapas 2 y 3** (plan §25.6-25.10): agrega `L, M, N,
  O, R, S, T, U, W, X, Y, AB, AC, AD, AE, AF, AG, AH, AI, AJ, AK, AL, AM,
  AN, AO, AP, AQ, AR, AS, AT, AU, AV`. `L` (¿participó en una subasta?)
  homologa contra `Subastas!Sub_Baj` (confirmado por el usuario) +
  `Configuración`+`Mes`+`Dia`+`Hora_dia`; `M` (¿SoC sobre el mínimo?) y
  `AE`/`AF` (energía asignada por bloques) usan la hoja `Resumen BESS` de
  `Centrales.xlsx` — resultó ser la MISMA tabla que la hoja `Resumen` del
  libro original (no hacía falta una hoja nueva, ver plan §25.8). `AG:AL`
  (Prorratas) salen de una tabla dinámica **derivada de `Subastas`, no de
  un archivo externo** (`construir_prorrata_sscc()`); `AM:AR` (FD) salen
  de homologar la central contra `Diccionario!A→B` y buscar en `FD` un
  bloque de 4 "Cuarto de Hora" (`construir_dic_mapeo_diccionario()`,
  `calcular_fd_prorrateado()`) — en ambos grupos, `CTF` (`AI/AL/AO/AR`) es
  **0 hardcodeado** (confirmado por el usuario: no existe, y así lo hace
  también el VBA original). `AS/AT` combinan lo anterior con `AE`/`AF`
  (`_calcular_costo_ponderado()`); `AU/AV` promedian `AB`/`AD` por grupo,
  activados solo si la suma GLOBAL de energía por ventana (todas las
  centrales, no por grupo) supera/baja de ±10. `N/O/R/Y/AB/AC/AD/AE/AF`
  se calculan por grupo (central=`clave` + ventana=`Copia_Ventana`);
  `S/T/U` no agrupan; `W/X` son **globales** (no por grupo). **Nombres de
  columna reales, confirmados contra un archivo real**
  (`NOMBRES_CALCULO_E_COSTOS`, plan §25.9) — ya no son placeholders; `AG:AL`
  y `AM:AR` comparten a propósito los mismos 6 nombres cortos (así es en
  el archivo real, se distinguen por un encabezado de grupo que no se
  replica en este esquema de una sola fila; `Total` también se repite
  entre `U` y `AX`, así que a esas columnas hay que llegar por posición,
  no por nombre).

  **Calculo E Costos, etapa 4** (plan §25.11): cierra la hoja con `AW`
  (`Descuento FD`), `AX` (`Total` = `AU+AV-AW`) y `AZ` (`Monto a
  compensar`, por grupo, nunca negativo). El umbral de subida/bajada que
  las bloqueaba **no era un archivo externo**: se deriva de `Subastas` +
  `Subastas!N` contando filas por central+ciclo+tipo
  (`construir_dic_umbrales_subastas()`), igual que la Prorrata SSCC. Y
  `Subastas!N` (llamada `"Ciclo"` en el archivo real, plan §26.8 — se
  creyó "Energía SSCC" hasta corregirse) **no es una energía**: es el
  `Ciclo de Carga del mes` de `Calculo E Costos` homologado por
  `Hora_mes` + `Configuración` (`calcular_subastas_ciclo()`), que viene
  de `Medidores` — por eso la "dependencia circular" que se había
  anotado no existía. Fuera de alcance: la columna `AY` (que la macro
  original tampoco escribe).

  **Las hojas de cálculo tienen su propio botón** (`SECCIONES_PAGOS`,
  mismo patrón que `SECCIONES_CONSOLIDADO`). La hoja que no se pide se
  preserva tal cual estaba en el archivo existente (no se recalcula ni se
  borra), mismo criterio de `escribir_salida()` para las hojas de
  entrada. Solo `Calculo E Costos` exige el archivo
  `SSCC_Desempeño_*`; `Calculo RE545` no lo necesita, asi que tildar
  solo esa seccion no lo pide.

  **Calculo RE545** (plan §26): la hoja hermana, **completa** (`A:CE`).
  La alimenta la MISMA macro de traspaso: `A:G`, `K` y `P` van iguales a
  las dos hojas y lo que se reparte es la energía, según
  `Medidores!Ventana_No_Completa` (`= 1` → E Costos; cualquier otra cosa,
  incluido vacío → RE545). Solo RE545 recibe `T` (`Ventana de
  valorizacion`) y `R` (`CMg Promedio`, `CMg!I`) — por eso
  `construir_dic_cmg()` guarda el par `(CMg!F, CMg!I)`. Es casi toda
  fórmulas en la hoja, no valores escritos por macro (al revés que
  E Costos). Tiene **dos tablas de distinto largo**: el bloque principal
  (una fila por cuarto de hora) y el resumen `AW:BG` (una fila por
  central+ventana), que se escriben lado a lado. `AY` (`Oferta
  Completa`) no es fórmula ni macro: es la columna `Completa` del
  resumen central+ventana que ya alimenta `Medidores!T`. **Trampa:**
  `R`, `S`, `T`, `U` y `V` existen en las dos hojas y significan cosas
  distintas en cada una (`U` es `Total` en E Costos y `EiniT` en RE545);
  lo mismo con el `VLOOKUP` sobre `Resumen BESS`, donde el índice 2 es
  `Pmax (MW)` y el 4 es `Capacidad (MWh)` (ver plan §26.7).
- **Consume:**
  - `<CARPETA_BASE>/Medidas/Medidas_SAE.xlsx` (hoja `Medidas`)
  - Un archivo `.xlsx` dentro de `<CARPETA_BASE>/Medidas/` cuyo nombre
    contenga "SOC" y el AAMM ingresado por el usuario (no hay un nombre de
    archivo fijo; debe existir exactamente uno)
  - `<CARPETA_BASE>/Auxiliares/<algo>Homologacion<algo>.xlsx` (hoja `homol`:
    `Punto de Medida` + `Canal` → `clave` + `Flujo`; hoja `Gen real`
    (ahí `Canal` es la unidad de la medida: `MWh…`/`kWh…`, se lee el
    principio),
    opcional: las centrales que se miden por la API de operación real), solo
    para generar `Medidas_SAE.xlsx`
  - `<CARPETA_BASE>/Auxiliares/Centrales.xlsx` (hojas `Resumen BESS` y
    `Diccionario`; en el formato viejo de `Diccionario` las columnas E/F/G
    — índices 4/5/6 — se usan específicamente para homologar Ofertas SSCC,
    y en el nuevo eso sale de las columnas `Subastas`/`Ofertas` hacia
    `Balance_BESS`; `Resumen BESS` columnas
    `Nombre activo`/`Barra inyección` se usan para `Calculo E Costos!Barra`)
  - Un archivo `.xlsx`/`.xlsm`/`.xlsb`/`.xls` dentro de
    `<CARPETA_BASE>/Ofertas/` cuyo nombre contenga "OfertasSSCC" (más
    reciente si hay varios)
  - `<CARPETA_BASE>/Cmg/cmg.xlsx` (nombre literal fijo, sin AAMM). No se
    descarga: lo genera el propio programa con `generar_cmg()` a partir de
    `<CARPETA_BASE>/Cmg/cmg<AAMM>_def_15minutal.csv`, que a su vez se baja
    de la unidad de red con `traer_csv_cmg()` (ver
    `Script/Cmg/Extrae_CMG_barras.py`)
  - Un archivo Excel dentro de `<CARPETA_BASE>/SSCC_Desempeño/` cuyo nombre
    empiece con "SSCC_Desempeño_" (más reciente si hay varios), hojas `CPF
    Horario` y `CSF Horario`
  - Los Access `OfertasSSCCAdj<AAAAMMDD>[_HH].accdb` del período que estén
    en `<CARPETA_BASE>/Subastas/DB subastas/` — **el origen real de la hoja
    `Subastas` desde esta sesión** (ver `Script/Subastas/`). La carpeta la
    crea el programa si no existe y se llena con el botón "Traer".
  - La planilla `3_REMUNERACIÓN_SUBASTAS_E_ID_*` **ya no se usa** (el
    usuario lo confirmó): se sacó de la ventana y del código. Nunca fue el
    origen
- **Produce:**
  - `<CARPETA_BASE>/Balance_BESS.xlsx` — **una sola planilla**, con las
    hojas de fin a inicio (`ORDEN_HOJAS_SALIDA`): `Resumen`,
    `PRORRATA_RETIROS`, `COMPENSACION_CENTRAL`, `Calculo RE545`,
    `Calculo E Costos`, `Subastas` (con sus nombres reales), `FD` (los
    bloques CSF y CPF lado a lado, columnas A:M y Q:AE, con sus nombres
    reales), `CMg`, `Ofertas SSCC` (las tablas W:Y y AB:AE equivalentes,
    una al lado de la otra — ver `_escribir_tabla_con_titulo()`) y
    `Medidores`. `'Calculo E Costos'!AM:AR` sale de la hoja `FD` de la
    misma planilla (`leer_fd_consolidado()`), no de releer el
    `SSCC_Desempeño_*`: las hojas de entrada son la única foto con la que
    se calcula, igual que ya pasaba con `Medidores` y `Subastas`.
  - `<CARPETA_BASE>/Control_corrida.xlsx` — el control de la corrida:
    `Ejecucion`, `Alertas` y `Log`. Sale aparte a pedido del usuario, para
    que no ensucie la planilla de trabajo (`escribir_control()`).
- **Expone (funciones clave agregadas hasta ahora, además de las básicas
  de E/S y homologación):**
  - Diagnóstico de homologaciones: `_avisar_claves_sin_mapeo()` agrupa por
    valor los cruces faltantes o con dato vacío y los registra como
    `[AVISO]`; se usa para barra, Pmax (en las dos hojas), capacidad y
    eficiencia. Los cruces que no son un diccionario `clave -> valor`
    avisan en su propia función: `calcular_ae_af()` (Pmax presente pero
    en 0), `calcular_prorratas()` (central sin ninguna hora en la
    Prorrata SSCC) y `calcular_reservas_re545()` (central ausente de
    toda la hoja `Subastas`, con lo que `AC:AT` y `AU` quedan en 0).
  - Estructura del caso: `revisar_estructura(carpeta_base, aamm=None)` →
    `(rutas, filas)`, donde cada fila es un dict `{id, etiqueta, nivel,
    estado, detalle, ruta, es_carpeta, origen}` (`estado`:
    `ok`/`falta`/`pendiente`; `ruta` + `es_carpeta`: lo que la ventana
    convierte en link — ver `Balance_BESS.py`; `origen`: `None` o
    `{id, titulo, etiqueta}` de `origenes.py` cuando lo de esa fila viene
    de afuera del caso). El **nivel**
    (0 = raíz del caso, 1 = dentro de una carpeta/archivo, 2 = un nivel
    más) lo pone `nucleo` porque es estructura, no dibujo; el **id** es
    estable y es lo que la ventana usa para colgarle el botón que
    corresponda. Helpers: `_fila()`, `hojas_de(ruta)` (nombres de hoja de
    un Excel) y `hojas_con_datos(ruta)` (`{hoja: tiene datos}`, usado para
    el estado hoja por hoja de las dos salidas: una hoja preservada que
    nunca se generó queda con una sola celda vacía y tiene que verse
    PENDIENTE, no generada), `_filas_de_hojas()`.
  - De dónde viene lo que se trae (`origenes.py`, para los links de la
    ventana): `ORIGENES` (`id -> (titulo, etiqueta, resolver)`),
    `origen(id)` → `{id, titulo, etiqueta}` y `ruta_origen(id, aamm)` →
    la carpeta exacta. Los ids son `fd`, `cmg_csv`, `cmg_xlsx`,
    `subastas`, `fma_cpf`, `fma_csf`, `fma_ctf`. La ruta se resuelve al
    hacer click, no al pintar el diagrama: mirar una unidad de red
    desconectada tarda, y el diagrama se repinta en cada revisada.
  - Ofertas SSCC: `buscar_archivo_ofertas`, `construir_resumen_ofertas_sscc`,
    `cargar_resumen_en_medidores`, `calcular_r`, `calcular_s`,
    `construir_resumen_ventana_oferta`, `calcular_t`.
  - CMg/FD/Subastas: `buscar_archivo_sscc_desempeno`,
    `buscar_archivo_subastas`, `leer_cmg(ruta_cmg, registrar=print)` →
    `df_cmg`; `construir_fd(ruta_sscc, registrar=print)` →
    `(df_fd_csf, df_fd_cpf)`; `construir_subastas(ruta_subastas, registrar=print)`
    → `df_subastas`.
  - Calculo E Costos (etapa base): `_normaliza_cuarto(valor)` →
    texto (replica `NormalizaCuarto`); `construir_dic_cmg(df_cmg)` →
    `dict` clave `"BARRA|CUARTO"` → valor Q; `construir_mapa_barra(resumen_bess)`
    → `dict` nombre de central normalizado → barra de inyección;
    `construir_dic_resumen_factor(resumen_bess)` → `(dict` nombre de central
    normalizado → `Pmax (MW), umbral_soc_minimo)` (misma hoja `Resumen BESS`
    que `construir_mapa_barra`, ver plan §25.8);
    `construir_calculo_e_costos(df_medidores, mapa_barra, dic_cmg, registrar=print)`
    → `df_ecostos`; `escribir_pagos_bess(ruta_salida, df_ecostos, registrar=print)`
    escribe "Calculo E Costos"/"Calculo RE545" con `startrow=1` (fila 1 libre)
    y llama a `_escribir_encabezados_grupo()` para las celdas combinadas de
    `GRUPOS_CALCULO_E_COSTOS`/`GRUPOS_CALCULO_RE545` — los nombres de columna
    quedan en la fila 2 y los datos desde la fila 3 (antes: fila 1 y 2).
  - Calculo E Costos (etapa 2): `calcular_l(df_ecostos, df_subastas)`,
    `calcular_m(df_ecostos, umbral_soc_minimo)`, `calcular_n_o(df_ecostos)`,
    `calcular_r_ecostos(df_ecostos)` (sufijo `_ecostos` a propósito:
    Medidores ya tiene su propia `calcular_r()`, lógica no relacionada — no
    fusionarlas), `calcular_s_t_u(df_ecostos)`, `calcular_w_x(df_ecostos)`,
    `calcular_y_ab_ac_ad(df_ecostos)`, `calcular_ae_af(df_ecostos, dic_factor)`
    (usa `_calcular_asignacion_energia()`).
  - Calculo E Costos (etapa 3, plan §25.10): `construir_prorrata_sscc(df_subastas)`
    → tabla dinámica (pivot); `construir_dic_prorrata(tabla_prorrata, registrar=print)`
    → `dict` central+hora_mes → `(CPF, CSF)`; `calcular_prorratas(df_ecostos, dic_prorrata)`
    → `(AG, AH)`; `construir_dic_mapeo_diccionario(diccionario)` → `dict`
    (tercera lectura de `Diccionario`, distinta de `construir_homologacion`
    y `_mapas_homologacion_fge` — no fusionar); `_calcular_bloque(valor)`;
    `construir_dic_fd_bloque(df_fd, columna_id, columna_mas, columna_menos)`;
    `calcular_fd_prorrateado(df_ecostos, dic_mapeo, dic_fd_csf, dic_fd_cpf)`
    → `(AM, AN, AP, AQ)`; `_calcular_costo_ponderado(...)` +
    `calcular_as_at(df_ecostos)` → `(AS, AT)`; `calcular_au_av(df_ecostos)`
    → `(AU, AV)`.
  - Todas combinadas por
    `completar_calculo_e_costos_grupos(df_ecostos, df_subastas, dic_factor, umbral_soc_minimo, diccionario, df_fd_csf, df_fd_cpf, registrar=print)`
    → `df_ecostos` con L/M/N/O/R/S/T/U/W/X/Y/AB/AC/AD/AE/AF/AG/AH/AI/AJ/AK/AL/AM/AN/AO/AP/AQ/AR/AS/AT/AU/AV
    agregadas Y renombrada a nombres reales (`NOMBRES_CALCULO_E_COSTOS`,
    plan §25.9) — mismo patrón que `NOMBRES_FD_CSF`/`NOMBRES_SUBASTAS`;
    `AG:AL` y `AM:AR` comparten a propósito los mismos 6 nombres cortos
    (así es en el archivo real). `GRUPOS_CALCULO_E_COSTOS` (mismo criterio
    que `GRUPOS_CALCULO_RE545`, ver más abajo) documenta los encabezados de
    grupo reales (`Dia`, `Nombre`, `BESS`, `Componente 2`, `Prorratas (-)`,
    `Prorratas (+)`, `FD`, `Componente 1`), confirmados con
    `docs/Libro1_Subastas_real.xlsx` (hoja `E COSTOS`, la única con los
    merges de Excel intactos).
  - `construir_medidores(df_sae, df_soc, anio, mes, ruta_ofertas, diccionario, registrar=print)`
    → `(df_medidores, avisos, df_wxy, df_resumen_ventana)`.
  - `escribir_salida(df, ruta_salida, avisos, incidencias, df_wxy=None, df_resumen_ventana=None, df_cmg=None, df_fd_csf=None, df_fd_cpf=None, df_subastas=None, ruta_existente=None, hojas_regenerar=None, registrar=print)`
    — `hojas_regenerar=None` (por defecto) regenera las 5 hojas de datos;
    si es un `set` con algunos nombres de `_HOJAS_CONSOLIDADO`, las que NO
    estén en el set se copian tal cual desde `ruta_existente`
    (`_copiar_hoja_existente()`, copia cruda vía `openpyxl`, sin fórmulas ni
    formato) en vez de recalcularse.
  - `SECCIONES_SALIDA` — las dos mitades juntas, en el orden en que quedan
    las hojas del libro (`(prefijo_id, sección)`), con un `assert` que
    obliga a que ese orden sea exactamente `ORDEN_HOJAS_SALIDA`: si alguien
    agrega una hoja en un lado y se olvida del otro, revienta al importar.
  - `escribir_control(ruta_control, df_log=None, registro=None, ...)` —
    escribe `Control_corrida.xlsx` (`Ejecucion`, `Alertas`, `Log`),
    preservando las hojas que esa pasada no reescribe.
  - `SECCIONES_CONSOLIDADO` — tupla de `(id, etiqueta, descripción, hojas)`
    por cada hoja de ENTRADA de la planilla (una por botón).
    `"medidores"` y `"ofertas_sscc"` son ids SEPARADOS (una casilla cada
    uno). La dependencia entre las dos se dio vuelta: ahora `Ofertas SSCC`
    necesita `Medidores` (de ahí salen las centrales y las ventanas) y no
    al revés, así que tildar solo `"medidores"` NO abre el archivo
    `*OfertasSSCC*` ni lo exige; tildar `"ofertas_sscc"` recalcula
    Medidores en memoria pero no lo reescribe si no se pidió. `"cmg"`,
    `"fd"`, `"subastas"` siguen siendo independientes de punta a punta.
  - `generar_consolidado(carpeta_base, aamm, secciones_activas, registrar=print, progreso=None)`
    — genera/actualiza las hojas de entrada de `Balance_BESS.xlsx`
    recalculando solo las
    secciones pedidas; valida los archivos de entrada únicamente para esas
    secciones (si no se pide `"medidores"`, no exige
    Medidas_SAE/SoC/Centrales/Ofertas). Si el archivo no existe, se crea.
    Reemplaza a la vieja `ejecutar()`.
  - `generar_medidas_sae(carpeta_base, aamm, registrar=print, progreso=None)`
    — genera/actualiza `<CARPETA_BASE>/Medidas/Medidas_SAE.xlsx` corriendo
    los cuatro pasos seguidos (botón **Actualizar** de esa fila). El paso de
    la API de operación real es opcional: sin la hoja `Gen real` se escribe
    solo lo que viene de `homol`. Esa hoja la lee
    `Homologacion.leer_gen_real()`; `_resumir_diagnostico_medidas()` es lo
    que antes iba a `reporte_medidas_consolidadas.xlsx` y ahora va al log.
  - `_leer_hoja_con_encabezado(ruta, hoja, columnas_buscadas)` — lector
    genérico de hojas cuyo encabezado no está en la primera fila (las hojas
    reales traen un título arriba). `_leer_resumen_bess()` es ahora un caso
    particular de este.
  - `traer_csv_cmg(carpeta_base, aamm, registrar=print, progreso=None)` —
    copia el CSV 15-minutal del período de la unidad de red a
    `<CARPETA_BASE>/Cmg/` (botón **Traer**). Se copia en vez de
    leerlo directo de la red para que el caso quede autocontenido: una vez
    traído, `cmg.xlsx` se puede regenerar sin la unidad conectada y queda
    registrado con qué archivo se trabajó.
  - `generar_cmg(carpeta_base, aamm, ruta_csv=None, registrar=print, progreso=None)`
    — genera/actualiza `<CARPETA_BASE>/Cmg/cmg.xlsx` con el CSV que ya está
    en esa misma carpeta (botón **Generar**). Las barras salen de
    `barras_desde_resumen_bess(resumen_bess)`, que reusa
    `construir_mapa_barra()` — la MISMA fuente que `Calculo E Costos!Barra`,
    así que las dos puntas no se pueden desincronizar. El resto (ruta de
    red, formato del CSV, numeración del `Cuarto de Hora`, validación del
    layout) vive en `Script/Cmg/Extrae_CMG_barras.py`, ver su bloque más
    arriba.
  - `generar_pagos_bess(carpeta_base, registrar=print, progreso=None)` —
    genera/actualiza las hojas de cálculo de `Balance_BESS.xlsx`; lee
    `Medidores` Y `Subastas` de las hojas de entrada ya generadas de esa
    misma planilla (no los recalcula), y
    Centrales.xlsx/cmg.xlsx/`SSCC_Desempeño_*` frescos (este último, nuevo,
    para `AM:AR`). Sin `aamm` como parámetro: nada de las etapas 2/3 de
    Calculo E Costos lo necesita (todo sale de `Medidores`/`Subastas`, que
    ya traen Mes/Dia/Hora).
- **Parámetros fijos:** `INICIO_VENTANA = 10`, `UMBRAL_SOC = 0.06` (ver plan
  de migración §8).
- **Constantes de columnas:** `LETRA_A_CAMPO` (dict A→U de `Medidores`, su
  orden de inserción ES el orden final de columnas), `COLUMNAS_VACIAS`.
- **Depende de:** `pandas`, `openpyxl` (como engine de
  `pd.ExcelWriter`/`pd.read_excel`), `calendar` (stdlib, días del mes).

---

## Diferencias con el documento de dominio

_(vacío — no se detectaron diferencias entre `nucleo/`/`Balance_BESS.py`
y `docs/Plan_Traspaso_Python_Balance_BESS.md` al organizar el repositorio)._
