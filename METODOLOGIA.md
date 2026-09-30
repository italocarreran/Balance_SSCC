# METODOLOGIA.md — cómo se trabaja en este repositorio

> Documento vivo. Si en una sesión se descubre algo que la siguiente necesita
> saber, se agrega acá antes de cerrar.

---

## 0. Por qué existe este documento

Este repositorio no existe solo para "guardar el código". Existe para que un
asistente de IA pueda entender el sistema **sin leerlo entero**. Cada regla de
acá abajo sale de esa única necesidad: minimizar cuánto hay que leer para
hacer un cambio correcto.

Si el repositorio lo trabaja más de una persona o más de un asistente con
acceso de escritura, sin verse en tiempo real, este documento es además la
única red de seguridad que existe entre ellos. No son sugerencias.

---

## 1. Qué es este repositorio

Herramienta en Python que reemplaza, hoja por hoja, el cálculo hecho hoy en
`11_PAGOS_BESS_2607_Definitivo.xlsm` (Balance BESS / SSCC). Implementa hasta
ahora **Medidores** (incluida Ofertas SSCC), la carga de **CMg**, **FD** y
**Subastas**, y una primera etapa (base) de **Calculo E Costos**: dos
scripts interdependientes que comparten un `config.json` guardado junto al
código (no versionado, es de la herramienta, no del caso).

- `Balance_BESS.py` — ventana tkinter. Único punto de entrada para el
  usuario.
- `Script/` — paquete con todo el cálculo, sin dependencias de interfaz.
  `Balance_BESS.py` importa `Script.nucleo` y nunca al revés.
  - `Script/nucleo/` — el cálculo del caso, un módulo por etapa.
  - `Script/Cmg/Extrae_CMG_barras.py` — arma `cmg.xlsx` desde el CSV
    15-minutal.
  - `Script/Medidas/` — arma `Medidas_SAE.xlsx` desde las dos APIs del
    Coordinador (un módulo por cada uno de los cuatro scripts sueltos que
    había antes).

  Un módulo de etapa **no importa `nucleo`** (evita ciclos): recibe rutas y
  datos, y levanta su propia excepción (`ErrorCmg`, `ErrorMedidas`), que
  `nucleo` traduce a `ErrorEntrada`. Si necesita un helper de texto del
  núcleo (`normalizar`), se duplica: diez líneas valen menos que un ciclo
  de imports.

Cada caso a procesar vive en su propia carpeta ("carpeta base"), fuera del
repositorio, con la estructura fija documentada en
`docs/Plan_Traspaso_Python_Balance_BESS.md` §3.3. El repositorio no contiene
datos de ningún caso.

---

## 2. Regla de expansión de contexto

La regla que no se negocia, para cualquier sesión que toque este repositorio:

```
1. Leer MAPA.md                       (mapa corto: qué hace cada script/módulo)
2. Leer de este documento SOLO las secciones que hacen falta
3. Abrir completo ÚNICAMENTE el archivo que se va a modificar
```

**No abrir archivos vecinos "para tener contexto".** Si de verdad hace falta
uno más, pedirlo explícitamente y decir por qué. Leer de más es exactamente
el problema que esta metodología resuelve.

Para las reglas de negocio del cálculo (fórmulas de Excel replicadas, de
dónde sale cada dato, qué columnas siguen pendientes) el documento de
referencia es `docs/Plan_Traspaso_Python_Balance_BESS.md`. Es largo a
propósito y por eso mismo **no se lee entero**: se busca la sección puntual
con grep o por su encabezado (p. ej. "9.4" para la columna `Ventana`).

**Dónde vive el código**, con nombres exactos:

- `Balance_BESS.py` — módulo/entrada principal, se abre siempre primero.
  Vive en la raíz del repositorio.
- `Script/nucleo/` — paquete de cálculo compartido (fachada en su `__init__.py`).
- `Script/Cmg/Extrae_CMG_barras.py` — módulo de la etapa CMg.
- `Script/Medidas/*.py` — módulos de la etapa Medidas.
  Los nombres de archivo de los módulos usan guiones bajos, no espacios,
  para que sean importables.
- `Script/config.py` — el único módulo que lee y escribe `config.json`.
- `config.json` — estado por PC/usuario (última carpeta base elegida, AAMM) y
  la sección compartida `claves_api`. No se versiona (ver `.gitignore`).
- `config.ejemplo.json` — el formato de `claves_api` para copiar y pegar.
  Este sí se versiona; no tiene ninguna clave real adentro.
- `docs/Plan_Traspaso_Python_Balance_BESS.md` — documento de dominio.

No hay hoy scripts satélite. Si `Balance_BESS.py` se mueve de carpeta,
`Script/` debe moverse con él: la ventana hace `from Script import nucleo`,
y dentro del paquete los imports son relativos (`from .Cmg import ...`).

---

## 3. El set de documentos y para qué sirve cada uno

| Documento | Para qué | Quién lo edita y cuándo |
|---|---|---|
| `REGLAS.md` | Checklist obligatorio de inicio/cierre de sesión. Se lee **entero, primero**, antes que nada. | Se edita poco; es la red de seguridad, no debe crecer sin necesidad. |
| `METODOLOGIA.md` | Cómo se trabaja: regla de expansión, convenciones de código, trampas conocidas, decisiones ya tomadas. | Documento vivo — se actualiza cuando cambia algo estructural o se toma una decisión de diseño. |
| `BITACORA.md` | Registro de qué se hizo en cada sesión y qué quedó pendiente — lo que un `git log` no cuenta. | **Solo se agrega.** Nunca se edita ni borra una entrada vieja. La única excepción es la sección "Pendientes abiertos", que sí se edita porque es un estado, no un historial. |
| `MAPA.md` | Un bloque corto por script: qué hace · consume · produce · expone · depende de. Es la primera lectura de cualquier sesión. | Se actualiza cuando cambia la estructura o el rol de un script. |
| `README.md` | Puerta de entrada: instalación/uso y tabla de "querés X → leé Y". | Se actualiza si cambia la estructura de carpetas o el flujo de instalación. |
| `docs/Plan_Traspaso_Python_Balance_BESS.md` | Reglas de negocio del cálculo: fórmulas, de dónde sale cada dato, qué queda pendiente. Es largo a propósito y no se reescribe por comodidad. | Se corrige ahí mismo cuando el dominio cambia (p. ej. se confirma la lógica de una columna pendiente); nunca se duplica en otro archivo. |
| `docs/Estructura_Archivos_Reales.md` | Referencia rápida de CADA archivo Excel real (los 7 de entrada del proceso + los de validación/comparación): hoja, fila donde arrancan los encabezados, columnas reales en orden, trampas ya confirmadas. Se lee **antes de pedirle al usuario un archivo o una aclaración** que ya podría estar ahí. | Se agrega/corrige una fila cada vez que un archivo real (adjuntado por el usuario) confirma o corrige una estructura — mismo criterio que las correcciones de `docs/Plan_Traspaso...`. |

**Regla de prioridad de lectura:** `REGLAS.md` → `METODOLOGIA.md` →
`MAPA.md` → el archivo a modificar → (si hace falta) la sección puntual de
`docs/Plan_Traspaso_Python_Balance_BESS.md`.

**Si el usuario menciona o adjunta un archivo Excel real** (`Centrales.xlsx`,
un `SOC_*.xlsx`, una hoja de comparación contra la planilla 11, etc.): leer
`docs/Estructura_Archivos_Reales.md` primero. Ahí está la estructura real ya
confirmada de cada uno (hoja, fila de encabezados, columnas, trampas) — no
hace falta pedirle al usuario que la explique de nuevo ni volver a inspeccionar
un archivo ya documentado ahí, salvo que la sesión encuentre una diferencia
real (en cuyo caso se corrige ese documento, no se duplica la explicación en
otro lado).

**Donde el código y el documento de dominio difieren, manda el código.**
Las diferencias detectadas se anotan en una tabla al final de `MAPA.md`
("Diferencias con el documento de dominio"); al corregir el documento, se
borra la fila correspondiente.

No existe todavía un `INTERFACES.md` generado ni un generador de interfaces:
con dos archivos, `MAPA.md` alcanza. Si el proyecto crece a varios módulos,
reevaluar (ver §10 de la plantilla original de esta metodología).

---

## 4. Cómo se entregan los cambios

Por ahora este repositorio lo trabaja un asistente de IA con revisión
humana (no hay más de un asistente con push directo todavía). Si eso
cambia, esta sección debe actualizarse para documentar cómo se coordinan
entre sí (rama base compartida, `BITACORA.md` como único canal entre
sesiones que no se ven en tiempo real, prefijo de autor en los commits si
alguno autentica con el nombre de otra persona).

- Rama base única (la rama por defecto del repositorio remoto). Todo sale
  de ahí y todo vuelve ahí. Una rama de trabajo se fusiona y se borra; no
  se acumulan ramas vivas.
- Nunca reescribir el historial compartido (`push --force`, `commit --amend`
  sobre algo ya subido, `reset --hard` contra la rama remota).
- Si en algún momento se vuelve a un flujo sin push directo (el usuario sube
  y baja archivos a mano): entregar siempre el archivo completo, nunca un
  diff ni "reemplazá la línea X por esto" — quien lo recibe lo corre tal
  cual, y editar a mano es donde se cuelan los errores.

No hay todavía scripts de apoyo (`scripts/sincronizar.sh`,
`scripts/verificar.sh`). Mientras no existan, la verificación antes de
cerrar una sesión es manual: correr `python -m py_compile Balance_BESS.py
Script/nucleo/*.py Script/*/*.py`, `python -m unittest discover` y, si hay
un caso de prueba disponible, `nucleo.generar_consolidado(...)`/
`nucleo.generar_pagos_bess(...)` contra él. Lo que depende de una API
(`generar_medidas_sae`) se prueba monkeypatcheando las dos funciones de
descarga con DataFrames sintéticos con la forma de la respuesta real: el
resto del proceso es puro pandas y sí se puede verificar.

---

## 5. Convenciones de código establecidas

- **Interfaz de usuario:** una única ventana tkinter (patrón "carpeta base +
  Examinar + periodo AAMM + diagrama de la estructura del caso + log +
  barra de progreso + contador de tiempo"). El diagrama es un árbol de
  texto tipo consola (prefijos `├──`/`└──`/`│`, patrón tomado de un
  `Revisor_Reliquidacion.py` que el usuario dio como referencia): cada
  fila de `nucleo.revisar_estructura()` trae su **nivel** (0 = raíz del
  caso, 1 = adentro de una carpeta/archivo, 2 = un nivel más) y su **id**;
  la ventana traduce los niveles a prefijos (`_prefijos_arbol()`) y usa el
  id para decidir qué botón le cuelga a cada fila (`_boton_de_fila()`):
  `nucleo` sabe de estructura, no de árboles ni de botones. Los estados
  son exactamente tres: `ok` (verde), `falta` (rojo) y `pendiente`
  (ámbar). No agregar un cuarto estado sin actualizar `SIMBOLO` y
  `COLOR_ESTADO` en `Balance_BESS.py` a la vez.
  **Cada acción es un botón en su propia fila del diagrama**: no hay
  ventanas intermedias ni un botón "Ejecutar" único. Las columnas son
  `Estructura | Estado | Acción | Detalle`, con el botón a la izquierda
  del detalle en una celda de ancho fijo en píxeles (`ANCHO_ACCION`) para
  que el detalle arranque siempre en la misma columna. Las dos salidas se
  desglosan por hoja (una fila-hoja por sección de
  `nucleo.SECCIONES_CONSOLIDADO` / `SECCIONES_PAGOS`, igual que las hojas
  de `Centrales.xlsx`), y lo que no se actualiza se **conserva** tal cual
  estaba (no se recalcula ni se borra, ver
  `escribir_salida()`/`hojas_regenerar`). Todos los botones corren su
  función de `nucleo` en un hilo aparte reportando al log/barra de la
  ventana (helper `lanzar()`), y mientras algo corre quedan todos
  deshabilitados.
- **Credenciales:** las `user_key` de las APIs del Coordinador son **dos y
  distintas** (`prmte` para `medidas.coordinador.cl`, `generacion_real` para
  `operacion.coordinador.cl`) y **no viven en el código**: salen de la
  sección `claves_api` de `config.json`, que no se versiona. Se leen en el
  momento de usarlas (`Script/config.py` → `clave_api()`), no al importar,
  así se puede completar el archivo con el programa ya abierto. El valor no
  depende de quién corra el programa — es el mismo para todo el equipo —,
  pero como el archivo es local, cada uno lo pega una vez en su copia; el
  formato está en `config.ejemplo.json` y el error de `clave_api()` lo
  repite entero. (Antes estaban escritas en el código, en una sola
  constante `USER_KEY`; eso las dejaba versionadas y, al ser una sola,
  forzaba a que las dos APIs compartieran clave.)
- **Persistencia de configuración:** `config.json` junto al `.py`. Tiene
  dos clases de secciones: una **por PC/usuario** (`get_usuario()` =
  `hostname_usuario`, con la carpeta base y el AAMM recordados), para que
  varias personas puedan compartir la misma copia del script sin pisarse la
  carpeta recordada, y la sección **compartida** `claves_api` (nombre
  reservado), que es igual para todos. Todo el acceso al archivo pasa por
  `Script/config.py`; `Balance_BESS.py` solo le pide su propia sección, de
  modo que guardar la carpeta base nunca puede pisar las claves. Si el archivo existe pero no se puede leer/parsear, se
  ignora en silencio y se sigue — es preferible perder el ajuste recordado
  a que el programa no abra. La escritura reescribe el `config.json`
  completo (no es atómica todavía; ver §7).
- **Rutas de un caso:** nunca rutas absolutas ni dependientes de
  `Path(__file__).parent` para los archivos de un caso. Todo se deriva de la
  carpeta base vía `resolver_rutas()` en `nucleo/rutas.py`. `Path(__file__).parent`
  se usa únicamente para ubicar `config.json` (en `Script/config.py`).
- **Lectura de Excel:** con `pandas` (`read_excel`/`ExcelFile`) y escritura
  con `openpyxl` como engine. El SoC se extrae por **detección dinámica de
  bloques por encabezados** (`detectar_fila_nombres` → `detectar_bloques` →
  `extraer_soc`): nunca por letra de columna fija ni por offset constante
  entre el nombre de la central y sus columnas `Time Stamp`/`Value` (regla
  del plan, §6.1 de `docs/Plan_Traspaso_Python_Balance_BESS.md`).
- **Normalización de texto:** toda comparación de nombres de central/hoja
  pasa por `normalizar()` en `nucleo/utiles.py` (minúsculas, sin tildes, espacios
  colapsados). No reimplementar una variante local de esta función en otro
  archivo.
- **Homologación de nombres:** se resuelve siempre contra la hoja
  `Diccionario` de `Centrales.xlsx` vía `construir_homologacion()`. La
  primera réplica no corrige ni reinterpreta homologaciones aunque parezcan
  desplazadas; cualquier inconsistencia se reporta como aviso/incidencia,
  no se "arregla" en silencio.
- **Período del caso (AAMM):** lo ingresa el usuario en un campo de texto
  de la ventana (4 dígitos, ej. `2607`), no se infiere del nombre de ningún
  archivo. `nucleo.validar_aamm()` es la única función que valida el
  formato; todo lo demás (`buscar_soc`, `revisar_estructura`, `ejecutar`)
  recibe el AAMM ya como parámetro. El archivo de SoC dentro de `Medidas/`
  tampoco tiene un nombre fijo: solo debe contener "SOC" y el AAMM en
  cualquier posición del nombre (`_es_archivo_de_soc()`); el archivo en sí
  siempre es `.xlsx` (confirmado con el usuario — un CSV con "SOC"+AAMM en
  el nombre puede ser un archivo completamente distinto sin relación con
  el SoC, ver `METODOLOGIA.md` §7).
- **Columnas de `Medidores` (A:U):** el orden final de columnas sale de
  `LETRA_A_CAMPO`, cuyo **orden de inserción** es el orden de Excel.
  `COLUMNAS_VACIAS` (M, P, Q, U) son diseño confirmado, no trabajo
  pendiente. `V, W, X, Y, AB, AC, AD, AE` **no están en `LETRA_A_CAMPO`**:
  en la planilla original no son una columna por fila de `Medidores`, son
  tablas auxiliares de otro largo (central × día, central × ventana) que
  comparten esas letras de columna solo porque ahí había espacio libre. Se
  calculan y se escriben juntas en una sola hoja auxiliar
  (`HOJA_OFERTAS_SSCC = "Ofertas SSCC"`, una tabla al lado de la otra con
  su propio título vía `_escribir_tabla_con_titulo()`, que acepta tanto
  `fila_inicio` como `columna_inicio`) en vez de forzarlas a columnas
  `pd.NA` del mismo largo que A:U (ver plan de migración §20.1, §22, §24.3).
  El resumen intermedio equivalente a la hoja
  "Resumen Ofertas SSCC" del `.xlsm` original (con una columna por
  servicio `_RS`) es puramente auxiliar para construir la tabla W:Y — no
  se persiste en la planilla, solo vive en memoria dentro de
  `construir_medidores()`. `R`, `S`, `T` sí son columnas por fila y están
  implementadas: dependen de las macros de Ofertas SSCC
  (`Generar_Resumen_Ofertas_SSCC`, `Resumir_Medidores_Central_Ventana_
  Oferta_Completa`), replicadas fielmente a partir del código VBA y las
  fórmulas de Excel entregados (plan §20).
- **Ofertas SSCC:** archivo obligatorio (plan §17-18), se busca en
  `<CARPETA_BASE>/Ofertas/` con `buscar_archivo_ofertas()` — nombre debe
  contener "OfertasSSCC" (sin importar mayúsculas); a diferencia del SoC,
  si hay más de uno se toma el más reciente por fecha de modificación
  (replica exacta de `OSSCC_BuscarArchivoOfertas`, no una decisión nueva).
  Homologación de nombres para Ofertas SSCC usa específicamente
  `Diccionario!E/F/G` (índices 4/5/6 del DataFrame `header=None`) vía
  `_mapas_homologacion_fge()`/`_homologar_fge()` — es un mapeo DISTINTO del
  que usa `construir_homologacion()` para el SoC (que trata toda la fila
  como equivalencias simétricas); no confundir ni fusionar ambos.
- **Errores de entrada vs. errores inesperados:** un problema de datos de
  entrada (archivo faltante, ambigüedad de SOC, columnas faltantes, bloque
  sin `Time Stamp`/`Value`) se señaliza con `nucleo.ErrorEntrada`, con un
  mensaje explicativo para el usuario. No usar excepciones genéricas para
  esto: `Balance_BESS.py` distingue ambos casos para mostrar un mensaje
  distinto.
- **Columnas deliberadamente vacías:** se agregan igual al DataFrame de
  salida (como `pd.NA`, listadas en `COLUMNAS_VACIAS` para `Medidores`) en
  vez de omitirse, para que la forma de la hoja `Medidores` sea
  comparable con la planilla 11 aunque la columna no tenga valor.
- **CMg, FD, Subastas (plan §23):** replican únicamente las macros de
  *carga* de esas hojas (`Cargar_CMg_Desde_Archivo`,
  `Cargar_SSCC_Desempeno_En_FD`, `Cargar_Remuneracion_Subastas_Rapido`).
  Ninguna de las tres tiene un documento de dominio tan detallado como
  Medidores; sus nombres
  de columna (`NOMBRES_FD_CSF`, `NOMBRES_FD_CPF`, `NOMBRES_SUBASTAS`, plan
  §24) los confirmó el usuario contra un caso real, no se inventaron. Si
  aparece una columna sin ese respaldo, usar su letra de Excel tal cual
  (p. ej. `"N"`) en vez de inventarle un nombre de negocio no documentado —
  mismo criterio de "no adivinar" que el resto del proyecto. Ojo con
  nombres duplicados dentro de un mismo bloque (p. ej. `FD` repite "Hora
  Mes" en B y M, y en R y AE): se renombra con `set_axis()` recién al
  final, después de calcular todo con nombres de letra únicos — Python no
  prohíbe columnas duplicadas, pero indexar por ese nombre durante el
  cálculo sería ambiguo.
- **`FD` tiene el mismo patrón de "tablas de distinto largo compartiendo
  hoja" que Ofertas SSCC, pero por columnas en vez de por filas:** el
  bloque CSF (A:M, viene de `CSF Horario`) y el bloque CPF (Q:AE, viene de
  `CPF Horario`) se filtran y calculan por separado (pueden tener distinta
  cantidad de filas) y se escriben lado a lado (`escribir_salida()`, vía
  `startcol` en `df.to_excel()`), no una debajo de la otra.
- **Filtro BESS/SAE de FD y Subastas vs. el de Ofertas SSCC:**
  `_contiene_bess_o_sae_sin_bat()` (FD, Subastas) NO incluye "BAT";
  `_contiene_bess_o_sae()` (Ofertas SSCC) sí. Son dos filtros distintos que
  se parecen — no fusionarlos en una sola función aunque parezca tentador.
- **Subastas!N queda vacía a propósito:** su fórmula real depende de
  `'Calculo E Costos'!D/G/P`; la hoja ya existe (etapa base, plan §25)
  pero todavía no las columnas específicas que esa
  fórmula necesita. No se adivina su valor.
- **`Calculo E Costos` (plan §25), etapa base:** replica solo una parte de
  `Traspasar_Medidores_A_Calculos_Rapido` (traspaso A:G con D↔E
  invertidas, I/J según signo y `Ventana_No_Completa`, J→K, K→P) y de
  `Asignar_CMg_a_Calculos_Turbo` (columna Q, con `escribirR=False` — la
  hoja no incluye `Calculo RE545`). `H` (Barra) es fórmula
  (`VLOOKUP(G,Resumen!B:G,6,FALSE)` en el original) y se homologa por
  **nombre de columna** contra `Resumen BESS!Nombre activo`/`Barra
  inyección` (`construir_mapa_barra()`), no por posición: `Centrales.xlsx`
  no reproduce el layout `Resumen!B:G` del libro original. El resto de
  `Actualizar_Calculos_Columnas` (L, M, N, O, R, S, T, U, W, X, Y, AB:AF,
  AG:AX, AZ) se completó después, por etapas (plan §25.6 a §25.11);
  `Calculo RE545` completo sigue pendiente. Los nombres de columna, que
  en esa primera etapa eran placeholders derivados de los comentarios de
  la macro, ya están confirmados contra un archivo real (plan §25.9).

---

## 6. Generación de `INTERFACES.md`

No aplica todavía: no existe un generador de interfaces en este repositorio.
Con los módulos de hoy (`Balance_BESS.py`, `Script/nucleo/`,
`Script/Cmg/`, `Script/Medidas/`) alcanza con `MAPA.md`. Si se agregan
muchos más y esto deja de ser suficiente, documentar acá la decisión de
introducir un generador (o no) antes de empezar a usarlo.

---

## 7. Trampas conocidas

Tabla de solo agregar: cada vez que un bug cueste tiempo real de
investigación, se documenta acá con la regla que lo evita, **antes** de
cerrar la sesión que lo encontró. No se borran filas viejas salvo que la
causa raíz deje de existir en el código.

| Trampa | Regla |
|---|---|
| `Script/config.py` reescribe `config.json` entero sin escritura atómica. Un corte a mitad de escritura puede dejar el archivo corrupto. | Al tocar esa función, evaluar escritura atómica (escribir a un temporal y `rename`). Mientras tanto, `leer_config()` ya tolera un JSON corrupto devolviendo `{}`, así que el peor caso es perder la carpeta recordada, no romper el programa. |
| Si `Medidas/` tiene más de un archivo `SOC_AAMM.xlsx`, `buscar_soc()` lanza `ErrorEntrada` a propósito — no elige el más reciente. | No "arreglar" esto para que elija automáticamente por fecha de modificación: es una decisión deliberada del plan (§3.4) para no tomar en silencio el mes equivocado. |
| Las columnas `K, M, P, Q, R, S, T` de `Medidores` están en `COLUMNAS_PENDIENTES` como `pd.NA` porque su lógica exacta o su fuente (Ofertas SSCC) todavía no está definida. | No inventar una fórmula para completarlas "para que quede bonito". Cerrar primero la regla exacta en `docs/Plan_Traspaso_Python_Balance_BESS.md` §9, con el humano que conoce la planilla 11, y recién ahí implementar. |
| `calcular_ventana()` reinicia el contador por **bloque de filas consecutivas con la misma clave**, no por `groupby` sobre toda la central. | Si los datos de entrada no vienen ordenados por `clave` e `intervalo` antes de llamar a esta función, el resultado no coincide con la fórmula de Excel. `construir_medidores()` ya ordena con `sort_values(["clave", "intervalo"])` antes de calcularla; no quitar ese paso ni reordenar después. |
| El orden final de columnas usa `list(LETRA_A_CAMPO.values())` (el dict ya está declarado en el orden correcto de Excel). | Si en el futuro se necesitara reintroducir alguna letra de dos caracteres (AA, AB...) en `LETRA_A_CAMPO`, nunca ordenar sus claves con `sorted()`: "AA" < "B" como texto, lo que rompería el orden real de columnas de Excel. Ya pasó una vez en esta migración (ver `docs/Plan_Traspaso_Python_Balance_BESS.md` §19/§20). |
| La fórmula de `Medidores!V` usa `Diccionario!F` y `Diccionario!G` como alias hacia `Diccionario!E` (columnas 5,6,7 del sheet, índices 4,5,6 en el DataFrame `header=None`) — un mapeo posicional específico, distinto de `construir_homologacion()` (que usa toda la fila, sin posición fija). | No usar `construir_homologacion()` para resolver Ofertas SSCC ni `_mapas_homologacion_fge()` para el SoC: son dos bloques distintos de la misma hoja `Diccionario`, con reglas de lectura distintas. |
| `calcular_s()` no se reinicia por central: sigue siendo "igual a la fila anterior mientras `Ventana` no cambie" incluso cruzando de una central a otra. | Es fiel a la fórmula de Excel (`IF(L3=L2,S2,...)`, sin comparar `G`). Si dos centrales consecutivas terminan/empiezan con la misma `Ventana`, `S` no se reinicia — así es también en la planilla original, no es un bug a corregir. |
| Que un archivo se llame `SOC_2607.csv` (o cualquier nombre que contenga "SOC"+AAMM) no garantiza que sea el archivo de SoC de la etapa Medidores. Ya apareció un CSV con ese patrón de nombre que en realidad era un archivo de pagos/liquidación (columnas `Fecha_Hora, CONFIGURACION, Central, Pago, Tipo_pago, Bloque_15min`, sin ninguna columna de SoC), sin relación con `Medidores!J`. | El archivo de SoC real siempre es `.xlsx`, con la estructura de bloques horizontales `Status/Questionable/Time Stamp/Value` (ver `extraer_soc()`). Si un archivo que matchea el patrón de nombre no tiene esa estructura, **no asumir que el formato cambió**: es señal de que no es el archivo correcto. Preguntar antes de adaptar el parser a una estructura nueva. |
| El `Centrales.xlsx` real trae, en la hoja `Resumen BESS`, un título fusionado en la primera fila (`"Cuadro N° 1: Resumen BESS"`) **antes** de la fila de encabezados reales. Un primer intento leyó la hoja con `pd.read_excel(header=0)` (posición fija) y `construir_mapa_barra()` fallaba: no encontraba `'Nombre activo'`/`'Barra inyección'` porque esas columnas venían como `Unnamed: N`. | `leer_centrales()` ahora usa `_leer_resumen_bess()`, que detecta la fila de encabezados buscando los textos esperados (mismo criterio que `detectar_fila_nombres()` para el SoC), nunca por posición fija. Si en el futuro aparece otra hoja de `Centrales.xlsx` con un título similar, aplicar el mismo patrón, no asumir `header=0`. |
| **(Resuelto)** `nucleo.construir_calculo_e_costos()`/`completar_calculo_e_costos_grupos()` calculan TODO con nombres internos tipo letra/placeholder (`Mes`, `clave`, `Energia_Positiva`, `L`, `N`, `AB`...) y recién renombran a los nombres reales (`NOMBRES_CALCULO_E_COSTOS`) al final de `completar_calculo_e_costos_grupos()` — mismo patrón que `NOMBRES_FD_CSF`/`NOMBRES_SUBASTAS`. El usuario tardó 3 intentos en mandar el archivo correcto con la hoja "E COSTOS" (las dos primeras veces solo traía `FD`/`Subastas`). | Si se agrega una columna nueva a `Calculo E Costos`, calcularla con un nombre interno letra/placeholder y agregarla a `NOMBRES_CALCULO_E_COSTOS` al final, **nunca** usar el nombre real directamente en medio del cálculo (mismo motivo que FD: si dos columnas terminan compartiendo un nombre real, indexar por ese nombre a mitad de cálculo sería ambiguo). |
| `"OfertasSSCC"` tiene **tres** "s" seguidas al pasarlo a minúsculas (`"Ofertas"` termina en "s" + `"SSCC"` empieza con dos "s" más = `"...tas" + "sscc"` = `"...tasssc c"`). Un primer intento transcribió el literal a mano con solo dos "s" (`"ofertasscc"`) y `buscar_archivo_ofertas()` nunca encontraba ningún archivo real. | No transcribir a mano un literal derivado de un nombre con letras dobles/triples repetidas: calcularlo en tiempo de ejecución (`"OfertasSSCC".lower()`, constante `PATRON_NOMBRE_OFERTAS` en `nucleo/parametros.py`) y comparar contra eso. Se detectó con un test sintético antes de llegar a producción; si vuelve a fallar la detección del archivo de Ofertas, este es el primer sospechoso a descartar. |
| Al agregar `calcular_r()` para la etapa 2 de `Calculo E Costos`, se redefinió sin querer una función que YA existía con ese nombre (`calcular_r()` de Medidores, para `Oferta_Completa_Dia`) — Python no avisa: la segunda definición pisa a la primera en silencio, y como `construir_medidores()` llama a `calcular_r()` en tiempo de ejecución (no al definirse), el error solo aparece al correr esa parte, con un `TypeError` de argumentos que no dice nada sobre la causa real. Se detectó por un test sintético que corrió `generar_pagos_bess()` de punta a punta. | Antes de agregar una función nueva a `nucleo/`, buscar (`grep -rn "^def <nombre>(" Script/nucleo/`) si el nombre ya existe — la fachada re-exporta todo, así que dos módulos distintos con el mismo nombre se pisan igual que antes. Si dos hojas distintas tienen una columna con la misma letra pero lógica distinta (como el `R` de Medidores y el `R` de Calculo E Costos), usar un sufijo que distinga la hoja (`calcular_r_ecostos`, no `calcular_r`) en vez de reutilizar el nombre corto. |
| La hoja `Diccionario` de `Centrales.xlsx` se lee de **tres** formas distintas según qué la consume: `construir_homologacion()` (toda la fila como equivalencias simétricas, para el SoC), `_mapas_homologacion_fge()` (columnas E/F/G→E, para Ofertas SSCC) y `construir_dic_mapeo_diccionario()` (columna A→B, primera coincidencia gana, para `Calculo E Costos!AM:AR`). | No fusionar estas tres lecturas ni reusar una para lo que hace otra: son reglas de negocio distintas sobre la misma hoja, confirmadas en momentos distintos de la migración. Si aparece una CUARTA necesidad de homologación, no asumir que es igual a alguna de las tres — preguntar. |
| `NOMBRES_CALCULO_E_COSTOS` tiene valores DUPLICADOS a propósito: `AG:AL` ("Prorratas") y `AM:AR` ("FD") comparten los mismos 6 nombres cortos (`CPF(-)`, `CSF(-)`, `CTF(-)`, `CPF(+)`, `CSF(+)`, `CTF(+)`) porque así está en el archivo real (se distinguen por un encabezado de grupo en las filas 1-2 que no se replica en nuestro esquema de una sola fila). Indexar el DataFrame final por uno de esos nombres (`df["CPF(-)"]`) después de renombrar devuelve un DataFrame de 2 columnas, no una Series — un test que no lo espera falla con `ValueError: truth value of a Series is ambiguous`. | No es un bug: es el mismo patrón que el `"Hora Mes"` duplicado de `FD`. Si hay que acceder a una de las dos columnas después del rename (normalmente no hace falta, el rename es el último paso antes de escribir a Excel), usar posición (`df.columns.get_loc`/`.iloc`), nunca el nombre solo. |

---

## 8. Decisiones ya tomadas

Lista de solo agregar, para no volver a discutir lo mismo en cada sesión.

- **Primero replicar fielmente la lógica de la planilla 11; después
  simplificar.** No se reinterpretan reglas de negocio todavía no
  documentadas solo porque parezcan mejorables (plan §2, principio 1).
- **La carpeta base es la única selección manual del usuario.** Todo lo
  demás (`Medidas_SAE.xlsx`, `SOC_AAMM.xlsx`, `Centrales.xlsx`, la salida)
  se resuelve por ruta relativa desde ahí, para que mover
  `Balance_BESS.py` a otra ubicación no cambie qué entradas encuentra ni
  dónde escribe (criterio de portabilidad, plan §3.13).
- **Las columnas cuya regla no está confirmada se dejan pendientes
  (`pd.NA`) en vez de adivinarse.** Aplica hoy a `K, M, P, Q, R, S, T` de
  `Medidores`. `OfertasSSCC` en particular ni siquiera tiene ubicación de
  carpeta definida todavía (plan §3.6, §7).
- **El SoC se extrae por semántica de encabezados, nunca por posición
  fija.** Confirmado en el plan §16.1: ni letras de columna ni offsets
  constantes entre el nombre de la central y `Time Stamp`/`Value`, porque
  esa distancia varía entre archivos mensuales.
- **La planilla de salida es un artefacto de validación, no un archivo
  versionado.** Se genera por caso en la
  carpeta base del usuario y se ignora en git (ver `.gitignore`); el
  repositorio no guarda salidas de casos concretos. Se renombró porque ya
  no es solo la etapa Medidores: consolida varias entradas materializadas
  (Medidores, CMg, FD, Subastas) que alimentan las siguientes etapas del
  balance.
- **`V, W, X, Y, AB, AC, AD, AE` no son columnas de `Medidores` en Python.**
  Las fórmulas de Excel (`V3:V312`, no `V3:V26786`) muestran que son tablas
  auxiliares de otro largo que solo comparten letra de columna con
  `Medidores` porque ahí había espacio libre en la planilla. Forzarlas a
  columnas `pd.NA` del mismo largo que A:U (como se hizo antes de tener el
  código VBA) ya no es una aproximación razonable una vez que se pueden
  calcular de verdad: se escriben juntas en la hoja auxiliar
  `HOJA_OFERTAS_SSCC` (ver plan §20.1, §22). El resumen "Resumen Ofertas
  SSCC" del `.xlsm` original tampoco se persiste — es un paso intermedio
  que solo hace falta en memoria para construir la tabla W:Y.
- **El período AAMM lo escribe el usuario, no se adivina del nombre de un
  archivo.** `SOC_AAMM.xlsx` era solo un patrón conceptual en el plan
  original; en la práctica el archivo de SoC llega con nombres variables.
  Confiar en un regex sobre el nombre para extraer el período era frágil;
  pedirlo explícitamente en la ventana es la fuente de verdad y además
  sirve para validar el archivo de SoC encontrado (debe contener ese AAMM).
- **`Calculo E Costos` se implementa por etapas.** Vivió un tiempo en un
  archivo separado (`Pagos_BESS.xlsx`, nombre provisorio) a pedido del
  usuario; hoy vuelve a estar en la misma planilla que las entradas — ver
  la entrada "Una sola planilla de salida" más abajo. Las etapas: primero H (Barra) + CMg +
  traspaso base desde Medidores (elegido explícitamente por el usuario
  frente a la alternativa de traducir de una sola vez toda
  `Actualizar_Calculos_Columnas`, ~1500 líneas con dependencias profundas);
  después una etapa 2 con todo lo que NO dependía de la hoja `Resumen`
  (`L, N, O, R, S, T, U, W, X, Y, AB, AC, AD`, plan §25.6/25.7). Esa hoja
  `Resumen` del libro original resultó ser la MISMA tabla que
  `Centrales.xlsx!Resumen BESS` (confirmado con un archivo real, plan
  §25.8) — no hacía falta una hoja nueva, así que `M`, `AE` y `AF` se
  agregaron también. Una etapa 3 (`AG:AV`, plan §25.10) agregó las
  Prorratas (`AG:AL`, tabla dinámica **derivada de `Subastas`**, no un
  archivo externo), el FD homologado (`AM:AR`), el costo ponderado
  (`AS`/`AT`) y el ingreso/costo de Componente 1 (`AU`/`AV`) — el usuario
  confirmó que `CTF` (`AI/AL/AO/AR`) no existe y sale en 0 (coincide con
  el VBA original, que las deja hardcodeadas), y que el archivo de
  Subastas usado para nuestra hoja está corrido una columna respecto del
  original, lo que explicó por qué las letras que documentaba el VBA para
  `L` no coincidían con los encabezados reales. Una etapa 4 (plan
  §25.11) cerró la hoja con `AW`/`AX`/`AZ`, después de que el usuario
  entregara el documento de trazabilidad completo (hoy en
  `docs/Trazabilidad_11_PAGOS_BESS_2607_Definitivo.md`): el umbral de
  subida/bajada que las bloqueaba se deriva de `Subastas` +
  `Subastas!N`, y la "dependencia circular" que se había anotado no
  existía (`Subastas!N` depende de `Calculo E Costos!P`, que viene de
  `Medidores`, no de una columna calculada). Sigue pendiente toda la
  hoja `Calculo RE545`.
- **Cada salida tiene su propio botón "Generar" con casillas por sección,
  en vez de un único botón "Ejecutar" para todo.** Pedido explícito del
  usuario, con un archivo de referencia (`Revisor_Reliquidacion.py`) para
  el estilo de ventana (diagrama de carpetas + botón por fila). Al
  no pedir una sección, esa
  parte se **conserva** tal cual estaba (copia cruda de la hoja existente,
  no un recálculo ni un vaciado) — confirmado explícitamente con el
  usuario frente a las otras dos alternativas (recalcular todo siempre, o
  dejar vacío lo no tildado). Las 4 casillas de esa ventana no son 1:1 con
  cada archivo de entrada: `"medidores"` agrupa Medidas_SAE + SoC +
  Centrales + OfertasSSCC porque `construir_medidores()` los necesita
  siempre juntos, no se pueden actualizar por separado a ese nivel de
  detalle. Hoy cada hoja de las dos mitades tiene su propio botón en la
  fila de esa hoja, y no hay ninguna ventana intermedia de casillas.
- **Un período, una carpeta — y "no se sabe" no es "sí".** Regla del
  usuario: *"dos meses no pueden tener la misma carpeta"*. La ventana
  recuerda en `config.json` qué carpeta usó cada AAMM
  (`carpetas_por_periodo`), así cambiar de mes cambia de carpeta sola, y
  una carpeta ya anotada para un mes no se puede reusar para otro.

  Lo que estaba mal antes no era la falta de la pregunta, sino el
  **tercer caso**: cuando el nombre de la carpeta no traía ningún AAMM, el
  código decidía "no se puede saber → no molesto" y seguía trabajando
  sobre la carpeta del mes anterior en silencio. Por eso
  `carpeta_corresponde_al_periodo()` devuelve tres valores
  (`CARPETA_DEL_PERIODO` / `CARPETA_DE_OTRO_PERIODO` /
  `CARPETA_SIN_PERIODO`): con una duda se **pregunta**, no se supone.

  Y si el usuario cancela, se vuelve al período anterior: la ventana
  nunca queda con un mes escrito arriba apuntando a la carpeta de otro.

- **Una sola planilla de salida, ordenada de fin a inicio, y el control
  aparte.** Pedido explícito del usuario: *"combinar el consolidado
  entradas con pagos bess pero ordenados de fin a inicio, el fin es el
  resumen y el inicio las entradas"*, y *"que las hojas de Alertas,
  Ejecución y Log queden en una planilla diferente"*. Así que:
  `Consolidado_entradas.xlsx` + `Pagos_BESS.xlsx` = `Balance_BESS.xlsx`
  (`ARCHIVO_SALIDA`), con las hojas en `ORDEN_HOJAS_SALIDA` — `Resumen`
  arriba, `Medidores` abajo —, y `Control_corrida.xlsx` (`ARCHIVO_CONTROL`)
  al lado con `Ejecucion`, `Alertas` y `Log`.

  Las dos mitades se siguen **escribiendo por separado** (cada hoja tiene
  su botón y su cálculo detrás): lo que cambió es que la que escribe
  preserva las hojas de la otra (`_preservar_ajenas()`) y reordena el libro
  al final (`_ordenar_hojas()`). La alternativa —una sola función que
  escribiera las diez hojas de un viaje— hubiera obligado a recalcular todo
  en cada botón, que es justo lo que el usuario no quiere.

  Trampa que deja esto: el `ExcelWriter` **trunca** el archivo, así que el
  libro anterior se lee SIEMPRE antes de abrirlo (`_abrir_existente()`), no
  solo cuando se pide una actualización parcial. Si eso se saltea, escribir
  una hoja se lleva puestas las otras nueve.

- **Las hojas de salida no reproducen los auxiliares de la planilla
  original, y el orden de las hojas es el de lectura.** Pedido explícito
  del usuario ("ordena las planillas, quita los auxiliares innecesarios"),
  con el criterio que él mismo fijó: *"si es importante para que los
  coordinados entiendan y vean parte del cálculo no hay que sacarlo, si es
  trivial sí"*. Fuera quedan las columnas deliberadamente vacías, las
  claves auxiliares que no lee nadie y las copias exactas de otra columna
  (`Medidores!K`, `Calculo E Costos!X`, `Calculo RE545!BR`) más la columna
  sin nombre `Calculo RE545!BL`; se quedan todos los intermedios que dejan
  seguir el cálculo (curvas monótonas, `ranking cmg`, Componentes 1 y 2) y
  las CTF en 0. Nada de eso cambia el cálculo: las columnas se siguen
  calculando, sólo dejan de escribirse, y lo que alguna etapa posterior
  necesita se repone al leer (`reponer_auxiliares_medidores()`). Qué se
  escribe vive en una sola lista por hoja (`COLUMNAS_MEDIDORES_SALIDA`,
  `COLUMNAS_SALIDA_E_COSTOS`, `COLUMNAS_SALIDA_RE545`), que es también
  contra la que se ubican los encabezados de grupo: agregar una columna a
  `NOMBRES_CALCULO_*` sin mirar esa lista ya no alcanza.
- **El formato de las hojas (ancho, negrita, panel fijo, separador de
  miles) se aplica en un solo lugar, `Script/nucleo/formato.py`, al cerrar
  cada libro.** No se formatea al escribir cada tabla: así también quedan
  formateadas las hojas que se preservan de una corrida anterior, y
  cualquier hoja nueva sale igual sin tocar su escritura. La regla es que
  ese módulo no puede cambiar un valor: si lo hace, es un bug.
- **"Ejecutar todo" es un grafo, no un script que llama a todo en fila.**
  El orden, las dependencias y el paralelismo viven en
  `Script/nucleo/orquestador.py` (`TAREAS`/`GRUPOS`), separado de la
  ventana y de las funciones que hacen el trabajo: la ventana solo
  dibuja el plan que devuelve `planificar()`, y el orquestador no
  calcula nada propio, llama a las mismas funciones que los botones
  sueltos. Tres invariantes que no se pueden romper al tocarlo: dos
  tareas que escriben el mismo archivo comparten `recurso` (y por eso
  las hojas de cada salida se mandan JUNTAS en una sola llamada, que
  además escribe el libro una sola vez); una tarea corre solo si sus
  dependencias están al día o se rehacen en la misma corrida; y si una
  tarea falla no corre nada que dependa de ella. Agregar un paso nuevo
  es agregar una `Tarea` (con su fila de estado de
  `revisar_estructura`, sus requisitos y sus dependencias), no un
  `if` en la ventana.
- **Las entradas que pone la persona se dividen en dos.**
  `ENTRADAS_INICIALES` (Centrales con sus dos hojas, homologación,
  OfertasSSCC, SoC del período) bloquean el botón "Ejecutar todo": sin
  ellas no hay corrida posible y decirlo es mejor que ofrecer media.
  `ENTRADAS_TARDIAS` (el Excel de prorrata de retiros) no bloquean:
  las hojas que dependen de ellas quedan fuera del plan, a la vista y
  con el motivo. Mover una entrada de un grupo al otro es una decisión
  de producto, no un detalle: se decide con el usuario.
- **Lo que ya está escrito en la planilla se lee de ahí, no de
  su origen.** Vale para `Medidores`, `Ofertas SSCC`, `Subastas`, `FD`
  y ahora también `CMg` (antes la etapa de pagos reabría `cmg.xlsx`).
  Las hojas de entrada son la ÚNICA foto con la que se paga:
  si una entrada cambia después de generarlo, se regenera su hoja con
  su botón, no se lee el archivo nuevo por un costado. Y al leer varias
  hojas del mismo libro se abre UNA vez (`pd.ExcelFile`), porque cada
  `pd.read_excel(ruta, sheet_name=...)` vuelve a parsear el archivo
  entero.
- **El programa crea las CARPETAS de un caso, nunca sus archivos.** `crear_estructura_caso()` arma la carpeta del período y sus subcarpetas vacías, es idempotente y no borra ni mueve nada; los archivos de entrada los pone la persona (o los traen los botones "Traer"). Y no se impone una convención de nombre de caso: se propone el nombre del período anterior con el AAMM cambiado y el usuario lo edita antes de crear. Preguntar sólo cuando se puede saber (la carpeta no existe, o su nombre trae otro AAMM); si el nombre no dice nada del período, no se pregunta.
