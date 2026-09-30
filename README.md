# Balance_BESS

Herramienta en Python que reemplaza, hoja por hoja, el cálculo hecho hoy en
`11_PAGOS_BESS_2607_Definitivo.xlsm` (Balance BESS / SSCC). La planilla se
usa solo como referencia de validación; el proceso Python no depende de
información almacenada exclusivamente en ese libro.

Etapas implementadas: **Medidores** (incluye Ofertas SSCC), la carga de
**CMg**, **FD** y **Subastas**, las hojas **Calculo E Costos** y **Calculo
RE545**, y el cierre económico mediante **COMPENSACION_CENTRAL**,
**PRORRATA_RETIROS** y **Resumen**.

## Instalación

```bash
pip install -r requirements.txt
```

## Uso

```bash
python Balance_BESS.py
```

0. **Solo la primera vez:** copiar `config.ejemplo.json` como `config.json`
   (junto a `Balance_BESS.py`) y pegar adentro las dos claves de las APIs del
   Coordinador. Son **distintas** entre sí y el valor es el mismo para todo
   el equipo; `config.json` no se versiona, así que cada uno lo pega una vez
   en su copia. Solo hace falta para el botón que baja `Medidas_SAE.xlsx`:
   el resto del programa anda sin claves.

   ```json
   {
     "claves_api": {
       "prmte": "PEGAR_AQUI_LA_CLAVE",
       "generacion_real": "PEGAR_AQUI_LA_CLAVE"
     }
   }
   ```

   `prmte` es la de `medidas.coordinador.cl` y `generacion_real` la de
   `operacion.coordinador.cl`. Si falta alguna, el programa lo dice con el
   formato exacto para pegar.
1. Elegir la **carpeta base** del caso (ver estructura abajo). El programa
   recuerda, por PC/usuario y en el mismo `config.json` (en su propia
   sección: guardar la carpeta no pisa las claves), la última carpeta usada
   **y qué carpeta usó cada período** (`carpetas_por_periodo`), que es lo
   que le permite volver sola a la del mes que elijas.
2. Ingresar el **período (AAMM)** en el recuadro de la ventana: 4 dígitos,
   año+mes simplificado (ej. `2607` para julio de 2026). No se adivina del
   nombre de ningún archivo — es el dato con el que el programa ubica el
   SoC del período dentro de `Medidas/` y el CSV de CMg dentro de `Cmg/`.
3. La ventana detecta automáticamente las entradas y las dibuja como un
   diagrama de carpetas (`OK` / `FALTA` / `PENDIENTE` por cada una). El
   nombre de cada archivo y de cada carpeta es un **link**: el click abre
   esa carpeta en el explorador (la que contiene al archivo, si es un
   archivo — nunca se abre el archivo). Y las filas de lo que viene de
   afuera del caso dicen de dónde en su detalle, también como link:
   `Origen: DCO` en el `SSCC_Desempeño_*`, `Origen: CMg Reales` en el CSV,
   `Origen: progdiar_adjudicaSEN` en `DB subastas/`, y `Origen inputs: ...`
   en `cmg.xlsx` y en las tres de FMA, que no se traen hechas pero se arman
   con insumos que sí vienen de afuera.
   - **Cada período trabaja en su propia carpeta: dos meses nunca
     comparten una.** Al cambiar el AAMM:
     - Si ese período ya tuvo carpeta, la ventana **vuelve sola a ella**
       (se acuerda de qué carpeta usó cada mes, por PC/usuario, en
       `config.json`).
     - Si no, y la carpeta abierta es la de otro mes (lo dice su nombre, o
       que ya quedó anotada como la de ese otro mes), **pregunta**:
       **Examinar…** para elegir la carpeta de este período, **Crear la
       carpeta** para armarla nueva, o **Cancelar**, que vuelve al período
       anterior — el mes nuevo nunca se queda apuntando a la carpeta del
       mes viejo.
     - Si el nombre de la carpeta abierta no trae ningún AAMM y nunca se
       anotó, no hay con qué saberlo: pregunta lo mismo, más un tercer
       botón **"Esta es la de AAMM"** para decir que sí lo es (queda
       anotada y no vuelve a preguntar por ese mes).
   - **Crear la carpeta** propone el nombre de la carpeta que tenías
     abierta con el AAMM cambiado (`Balance BESS 2607` → `Balance BESS
     2608`) y dónde crearla, y las dos cosas se pueden editar antes de
     confirmar. Se crean la carpeta y sus subcarpetas vacías
     (`Medidas/`, `Auxiliares/`, `Ofertas/`, `Cmg/`, `FD y FMA/`,
     `Subastas/DB subastas/`, `Prorrata retiros/`) y el caso queda abierto
     en ese período. También está el botón **Crear carpeta del caso**,
     abajo de todo.
   - A un caso al que le falte alguna subcarpeta se la completa desde ahí
     mismo (al elegir la carpeta con **Examinar**, o al cambiar el
     período): se crean sólo las que faltan y no se toca nada de lo que ya
     hay.
4. **Cada acción es un botón en la fila que le corresponde** (y además está
   el botón **Ejecutar todo**, abajo de todo — ver más abajo):
   - `Medidas/Medidas_SAE.xlsx` → **Actualizar** (baja el mes completo de las
     dos APIs del Coordinador y arma el archivo).
   - `Cmg/cmg<AAMM>_def_15minutal.csv` → **Traer** (lo baja de la
     unidad de red).
   - `Cmg/cmg.xlsx` → **Generar** (lo arma con ese CSV).
   - `Balance_BESS.xlsx` —la única planilla de salida— se desglosa por hoja,
     igual que `Centrales.xlsx`, y cada hoja tiene su botón (**Actualizar**
     las de entrada, **Calcular** las de cálculo). El de la fila del
     archivo, **Actualizar**, las hace todas: primero las entradas y
     después el cálculo. Lo que no se actualiza se conserva tal cual
     estaba; si el archivo todavía no existe, se crea.
   - `COMPENSACION_CENTRAL` tiene **Resumir**: la compensación
     de `Calculo E Costos` (por central y `Ciclo de Carga del mes`) y la de
     `Calculo RE545` (por central y `Ventana de valorizacion`), más el total
     que recibe cada empresa según el `Propietario` de `Resumen BESS`.
   - `PRORRATA_RETIROS` tiene **Traer**: consume `Prorrata 15min`
     del Excel de `Prorrata retiros/` (A: cuarto de hora, B: suministrador,
     C: prorrata) y reparte, cuarto de hora por cuarto de hora, el monto a
     compensar de `Calculo E Costos` + `Calculo RE545` entre las empresas
     que retiraron en ese mismo cuarto, según su peso. La hoja queda con
     tres cuadros: el monto de cada cuarto, el reparto de ese monto entre
     las empresas, y el total del mes de cada empresa. `Resumen` tiene
     **Asignar** y muestra cuánto `RECIBE`, `PAGA` y el `NETO` de
     cada empresa.

   Los botones dicen **solo el verbo** (`Traer`, `Generar`, `Resumir`,
   `Actualizar`, `Calcular`, `Asignar`): de qué se trata ya lo dice la fila
   en la que está el botón. Todos miden lo mismo y quedan en una columna
   recta, igual que la de `Estado`.

### El botón "Ejecutar todo"

Abre una ventana con el **plan de la corrida**: una fila por paso, tildada si
hay que hacerlo, con su estado (`se genera` / `se rehace` / `al dia` /
`bloqueada` / `sin tildar`) y, si no se puede, el motivo.

- **Se hace solo lo que falta.** Lo que ya está al día no se rehace, salvo
  que lo tildes; al tildarlo se tilda solo todo lo que sale de ahí (rehacer
  `Medidores` sin rehacer los pagos dejaría el libro mezclado entre dos
  corridas).
- **El botón se bloquea si falta una entrada inicial** — `Centrales.xlsx` con
  sus dos hojas, el Excel de homologación, el `*OfertasSSCC*`, el SoC del
  período, o el período `AAMM` —, y dice cuál falta. El Excel de prorrata de
  retiros, que suele llegar después, **no** bloquea: sus dos hojas quedan
  fuera del plan (a la vista y con el motivo) y se tildan a mano cuando
  aparece.
- **Respeta las dependencias y aprovecha lo que puede ir en paralelo**: las
  cuatro bajadas (medidas SAE, CSV de CMg, FD, subastas) y la generación de
  FMA corren juntas; `cmg.xlsx` espera su CSV; las hojas de entrada esperan
  a todo eso y se escriben **una sola vez**; las de cálculo esperan a las de
  entrada. Dos pasos que escriben en el mismo lugar nunca
  corren a la vez, y si un paso falla no se corre nada que dependa de él.

El grafo (quién depende de quién), el plan y la corrida viven en
`Script/nucleo/orquestador.py`; la ventana solo los dibuja. No hay ningún
cálculo nuevo ahí: llama a las mismas funciones que los botones sueltos.

Durante el cálculo, el registro muestra líneas con la **severidad** y el **id
del control** (`[ALTA] MAE-001: ...`) cuando una homologación no encuentra
correspondencia: central sin barra, capacidad,
eficiencia o Pmax en `Resumen BESS` (o con Pmax en 0, que deja `AE`/`AF`
vacías); central ausente del `Diccionario`; claves de FD o de CMg
inexistentes; central sin ninguna hora en la Prorrata SSCC; central que no
aparece en ninguna fila de `Subastas` (sus reservas y `AU` quedan en 0). El
cálculo conserva el comportamiento de la planilla —algunos faltantes quedan
vacíos y otros se rellenan con cero—, pero ahora informa la causa, las
centrales/claves afectadas y el impacto antes de continuar.

Cada corrida deja además, **en una planilla aparte** (`Control_corrida.xlsx`,
al lado de `Balance_BESS.xlsx`), las hojas de control — a pedido del usuario,
para que no ensucien la planilla de trabajo:

- **`Alertas`** — una fila por alerta, con su id, severidad, etapa, central y
  clave. Es el detalle **completo**: la pantalla muestra un resumen con hasta 15
  ejemplos, el archivo los guarda todos.
- **`Ejecucion`** — el estado de la corrida (`APROBADA` / `APROBADA CON
  ADVERTENCIAS` / `NO APROBADA - REQUIERE REVISION`), qué hojas se recalcularon
  en esta pasada, la conciliación de energía (`Medidores = Calculo E Costos +
  Calculo RE545`) con su tolerancia, y el manifiesto de las entradas: nombre,
  tamaño, fecha y `sha256` de cada archivo que alimentó el cálculo.
- **`Log`** — los avisos e incidencias de la generación de las hojas de
  entrada (antes era una hoja más de `Consolidado_entradas.xlsx`).

El catálogo de controles, con lo que está implementado y lo que falta, está en
`docs/Alertas_y_Controles_Traspaso_Python_BESS.md`.

Las pruebas de estos avisos se corren con `python -m unittest discover`
desde la raíz del repositorio.

## Estructura del repositorio

```text
Balance_BESS.py            <- la ventana (lo único que se ejecuta)
Script/
    nucleo/                <- el cálculo del caso, una etapa por módulo
        __init__.py            <- la fachada: `nucleo.<lo que sea>`
        parametros.py  utiles.py  avisos.py  rutas.py  estructura.py
        lectura.py  soc.py  ofertas_sscc.py  hojas_entrada.py
        subastas_accdb.py  fma.py  diccionarios.py
        columnas_compartidas.py
        ecostos*.py            <- las cuatro etapas de Calculo E Costos
        re545*.py              <- las cuatro partes de Calculo RE545
        medidores.py  escritura.py  proceso.py  traer.py  medidas_sae.py
    Cmg/
        Extrae_CMG_barras.py   <- arma cmg.xlsx desde el CSV 15-minutal
    Medidas/
        Homologacion.py        <- punto de medida + canal -> clave
        Descarga_PRMTE.py      <- API de medidas, por punto de medida
        Claves_Balance.py      <- calendario de cuartos + agrupación por clave
        Generacion_Real.py     <- API de operación real (hoja "Gen real")
```

`nucleo/__init__.py` es solo una fachada: re-exporta todo, así que
`nucleo.lo_que_sea` funciona igual que cuando era un único archivo.
Ver `MAPA.md` para qué hace cada módulo.

## Estructura de carpeta de un caso

```text
<CARPETA_BASE>/
├── Medidas/
│   ├── Medidas_SAE.xlsx     (botón "Actualizar")
│   ├── _trabajo/            (lotes descargados; no se muestra en la ventana)
│   └── <algún archivo .xlsx cuyo nombre contenga "SOC" y el AAMM,
│        ej. SOC_2607.xlsx, "resumen soc julio 2607.xlsx">
├── Auxiliares/
│   ├── Centrales.xlsx       (hojas "Resumen BESS" y "Diccionario")
│   │                        El "Diccionario" acepta dos formatos: el
│   │                        viejo (varias tablas lado a lado) y el nuevo,
│   │                        una sola tabla con encabezados
│   │                        Balance_BESS | FD | Subastas | Ofertas |
│   │                        FMA_CPF. El nuevo es el único que permite
│   │                        homologar la nomenclatura de FMA CPF.
│   └── <algún archivo Excel cuyo nombre contenga "Homologacion">
│                            (hojas "homol" y "Gen real")
├── Ofertas/
│   └── <algún archivo Excel cuyo nombre contenga "OfertasSSCC">
├── Cmg/
│   ├── cmg<AAMM>_def_15minutal.csv   (botón "Traer")
│   └── cmg.xlsx                      (botón "Generar")
├── SSCC_Desempeño/
│   └── <algún archivo Excel cuyo nombre empiece con "SSCC_Desempeño_">
├── Subastas/
│   └── DB subastas/                  (botón "Traer")
│       └── OfertasSSCCAdj*.accdb     <- el origen real de las subastas
├── Prorrata retiros/
│   └── Prorrata_Retiros_<AAMM>_pre.xlsx o _def.xlsx
│       (hoja "Prorrata 15min")
├── Balance_BESS.xlsx            <- LA salida: entradas y cálculo en un
│                                   solo libro, ordenado de fin a
│                                   inicio (Resumen, PRORRATA_RETIROS,
│                                   COMPENSACION_CENTRAL, Calculo RE545,
│                                   Calculo E Costos, Subastas, FD, CMg,
│                                   Ofertas SSCC, Medidores). Una fila
│                                   por hoja en la ventana, cada una con
│                                   su botón
└── Control_corrida.xlsx         <- Ejecucion, Alertas y Log de la
                                    última corrida
```

Ningún archivo (salvo `cmg.xlsx`) sigue un nombre fijo:

- **SoC**: cualquier `.xlsx` en `Medidas/` cuyo nombre contenga "SOC" y el
  AAMM ingresado en la ventana. Si hay más de un archivo que cumple la
  condición, el programa se detiene y pide dejar solo el del período
  correspondiente (no elige por fecha de modificación).
- **OfertasSSCC** y **SSCC_Desempeño_\***: cualquier archivo Excel
  (`.xlsx`/`.xlsm`/`.xlsb`/`.xls`) en su carpeta correspondiente cuyo
  nombre contenga (Ofertas) o empiece con (el otro) ese texto. Si hay más
  de uno, a diferencia del SoC, se toma automáticamente el más reciente por
  fecha de modificación — así lo hacen las macros originales de la
  planilla. La planilla `3_REMUNERACIÓN_SUBASTAS_E_ID_*` **ya no se usa**:
  las subastas salen de los Access de `Subastas/DB subastas/`.
- **Medidas_SAE.xlsx**: tampoco se arma a mano. El botón **Actualizar** de esa
  fila corre los cuatro pasos de un viaje:

  1. lee el Excel de homologación de `Auxiliares/` (hoja `homol`:
     `Punto de Medida` + `Canal` → `clave` + `Flujo`);
  2. baja las medidas de cada punto de medida del mes
     (`medidas.api.coordinador.cl`), por lotes y **reanudable**: si se corta,
     la corrida siguiente retoma donde quedó;
  3. arma el calendario de cuartos de hora del mes y agrupa por `clave`;
  4. **agrega** las centrales listadas en la hoja `Gen real` del **mismo**
     Excel de homologación, cuya medida viene de la API de operación real
     (`operacion.api.coordinador.cl`) y no de la API por punto de medida.

  Es el proceso más lento del programa (miles de llamadas a la API). Los
  archivos intermedios van a `Medidas/_trabajo/` y no aparecen en la ventana.

  La hoja **`Gen real`** tiene las **mismas cuatro columnas que `homol`**,
  con una lectura propia de cada una:

  | clave | Punto de Medida | Canal | Flujo |
  |---|---|---|---|
  | `SAE-ANDES-III` | `SAE PFV Andes Solar III (Inyección)` | | `1` |
  | `SAE-ANDES-III` | `SAE PFV Andes Solar III (Retiro de central)` | | `-1` |

  - `clave`: con qué nombre tiene que aparecer en `Medidas_SAE.xlsx`, igual
    que en `homol`. Dos filas pueden apuntar a la misma clave: se suman.
  - `Punto de Medida`: acá va el **`topologyName` exacto** de la API de
    operación real — es lo que identifica a la central en esa API, que no
    tiene el concepto de punto de medida.
  - `Canal`: acá es la **unidad** de la medida, y es lo único que decide
    si el valor se multiplica por mil o no. La API de operación real
    devuelve **MWh** y todo el balance trabaja en **kWh**. Se lee el
    **principio** del texto, así que sirve tal cual viene en el archivo
    real (`MWhD`, `MWhR`, `kWhD`, `kWhR`): la `D` y la `R` del final no
    entran en la cuenta. Vacía, o con un texto que no empieza con `MWh`
    ni con `kWh`, vale MWh.
  - `Flujo`: `1` / `-1`, igual que en `homol`. Si se deja vacío vale `1`.

  La hoja entera es opcional: si no existe, no se agrega ninguna central por
  ese camino y el resto del proceso corre igual.

  **`user_key`**: las dos APIs piden la misma clave. Vive en el código, en
  `Script/Medidas/comun.py` (constante `USER_KEY`) — un solo lugar para las
  dos, en vez de repetida en cada script como estaba antes. Si el
  Coordinador la cambia, se cambia ahí y nada más. Tené presente que, al
  estar en el código, queda versionada: cualquiera con acceso al
  repositorio la tiene.

- **cmg.xlsx**: única excepción con nombre literal fijo, dentro de `Cmg/`.
  Tampoco hay que armarlo a mano, y son dos pasos, cada uno con su botón en
  esa misma carpeta del diagrama:

  1. **Traer cmg_15min** copia el CSV 15-minutal oficial del período,

     ```
     T:\CMgReales 15MIN\<AAAA>\<AAMM>\Mensual\CMg\Cmg para balance\cmg<AAMM>_def_15minutal.csv
     ```

     a `<CARPETA_BASE>/Cmg/`, al lado de `cmg.xlsx`. Se copia (en vez de
     leerlo directo de la red) para que el caso quede autocontenido: se
     puede regenerar `cmg.xlsx` después sin la unidad conectada, y queda
     registrado con qué archivo se trabajó. La raíz `T:` está en una sola
     constante (`Extrae_CMG_barras.RAIZ_CMG_REALES`) por si cambia de letra.
  2. **Generar** arma `cmg.xlsx` con ese CSV, filtrando las barras que trae
     la columna `Barra inyección` de la hoja `Resumen BESS` de
     `Centrales.xlsx` — no hay ninguna lista de barras escrita en el
     código: se agregan o se sacan editando `Centrales.xlsx`.

La carpeta base puede estar en cualquier ubicación (disco local, red,
OneDrive); moverla o mover `Balance_BESS.py` a otro lugar no cambia el
resultado, siempre que la carpeta base seleccionada sea la misma.

## Documentación

| Querés... | Leé |
|---|---|
| **Saber qué estructura real tiene cada archivo Excel que arma el usuario (o que manda para validar), sin tener que pedírselo de nuevo ni que te lo vuelva a explicar** | `docs/Estructura_Archivos_Reales.md` |
| Entender cómo se trabaja en este repo (para un asistente de IA o alguien nuevo) | `METODOLOGIA.md` |
| El checklist obligatorio de inicio/cierre de sesión | `REGLAS.md` |
| Qué hace cada script, en dos líneas | `MAPA.md` |
| Historial de sesiones y pendientes abiertos | `BITACORA.md` |
| Reglas de negocio del cálculo y el plan completo de migración | `docs/Plan_Traspaso_Python_Balance_BESS.md` |
| El código VBA original, las fórmulas del `.xlsm` y de dónde sale cada dato | `docs/Trazabilidad_11_PAGOS_BESS_2607_Definitivo.md` |
| La hoja `Calculo RE545` real (recortada), con sus nombres de columna y fórmulas | `docs/Calculo_RE545_reducido_para_IA.xlsx` |
| La hoja `Subastas` real, con encabezados y fórmulas (fuente de la corrección de `NOMBRES_SUBASTAS`); también trae la hoja `E COSTOS` con los encabezados de grupo reales (celdas combinadas), fuente de `GRUPOS_CALCULO_E_COSTOS` | `docs/Libro1_Subastas_real.xlsx` |
| `Centrales.xlsx` y `SOC_AAMM.xlsx` reales (fuente de las correcciones de `detectar_fila_nombres()` y `construir_homologacion()`) | `docs/Centrales_real.xlsx`, `docs/SOC_real_2607.xlsx` |
| Primera comparación real vs Python de `Calculo E Costos` fila a fila (fuente de la corrección de la Prorrata SSCC) | `docs/Pagos_BESS_comparacion_real.xlsx` |

## Estado actual

`Balance_BESS.xlsx` es **la** planilla de salida: antes eran dos
(`Consolidado_entradas.xlsx` y `Pagos_BESS.xlsx`) y el usuario pidió
combinarlas. Sus hojas quedan ordenadas **de fin a inicio** — arriba el
`Resumen`, que es lo primero que se mira, y abajo las entradas de las que
sale todo:

```
Resumen | PRORRATA_RETIROS | COMPENSACION_CENTRAL | Calculo RE545 |
Calculo E Costos | Subastas | FD | CMg | Ofertas SSCC | Medidores
```

Las hojas de control (`Ejecucion`, `Alertas`, `Log`) viven aparte, en
`Control_corrida.xlsx`.

Las dos mitades se siguen escribiendo por separado (cada hoja tiene su
botón), y la que escribe **preserva** las hojas de la otra: actualizar
`Medidores` no toca el `Resumen`, y calcular el `Resumen` no toca
`Medidores`.

### Las hojas de entrada

- `Medidores` — 12 columnas: A:J de entrada más `Ventana` e `Indicador_SoC`
  calculadas. **Sin los auxiliares** de la planilla original (las cuatro
  columnas vacías M/P/Q/U, la clave auxiliar N y la copia K de la Ventana:
  ver `COLUMNAS_AUXILIARES_MEDIDORES`); la copia K la repone al vuelo
  `reponer_auxiliares_medidores()` cuando las hojas de cálculo la piden como
  "Ciclo de Carga del mes". **No depende de Ofertas SSCC**: su botón "Actualizar" no abre el archivo `*OfertasSSCC*`
  ni lo exige.
- `Ofertas SSCC` — las tablas auxiliares equivalentes a `Medidores!W:Y`
  ("Ofertas SSCC por dia") y `Medidores!AB:AE` ("Resumen ventana oferta"),
  una al lado de la otra. De ahí salen las columnas R
  (`Oferta_Completa_Dia`), S (`Indicador_Ventana_Oferta`) y T
  (`Ventana_No_Completa`) de la planilla original, que **ya no se escriben
  en `Medidores`**: se reconstruyen en memoria cuando hacen falta (T es la
  que reparte cada fila entre las dos hojas de cálculo). El resumen
  intermedio equivalente a la hoja "Resumen Ofertas SSCC" del `.xlsm`
  original es puramente auxiliar y no se persiste.
- `CMg` — copia ordenada de `Cmg/cmg.xlsx` (replica
  `Cargar_CMg_Desde_Archivo`).
- `FD` — datos de `CPF Horario`/`CSF Horario` filtrados por BESS/SAE, más
  sus columnas calculadas, con sus nombres reales de columna (replica
  `Cargar_SSCC_Desempeno_En_FD`). Los bloques CSF (A:M) y CPF (Q:AE) son dos
  tablas de distinto largo, lado a lado en la misma hoja.
- `Subastas` — datos de los Access de `Subastas/DB subastas/` filtrados por
  "Configuración" contiene BESS/SAE (funcionalmente equivalente a filtrar
  por Propietario: los nombres de central BESS empiezan con "SAE-"), más la
  columna "Clave horaria" calculada, con sus nombres reales de columna
  (`Concepto`, `Control`, `Sub_Baj`, ..., `Energía SSCC`, `FD`, `FMA`). Sale
  en el orden del origen (no se ordena por `Hora_mes`). La columna "Ciclo"
  queda vacía: depende de `Calculo E Costos`.

### Las hojas de cálculo

El libro abre por el `Resumen` (quién paga y quién recibe) y sigue hacia
atrás con el detalle: `PRORRATA_RETIROS`, `COMPENSACION_CENTRAL`,
`Calculo RE545` y `Calculo E Costos`.

- `Calculo E Costos` — traspaso desde `Medidores`, asignación de CMg, y casi
  toda `Actualizar_Calculos_Columnas` (`L, M, N, O, R, S, T, U, W, X, Y, AB,
  AC, AD, AE, AF, AG:AX, AZ`), con **nombres reales de columna** (confirmados
  contra un archivo real, hoja "E COSTOS"): `Configuracion`, `Barra`,
  `Descarga kWh`/`Carga kWh`, `SoC %`, `CMg`, `Adj SSCC`, `SoC sobre el
  minimo`, `ranking cmg`, `Valorizacion Descarga`/`Carga`, `Bloque ordenado`,
  `Curva monotona CMg Descarga`/`Carga`, `Energía descargada`/
  `cargada`, las Prorratas y el FD homologado (`CPF(±)`/`CSF(±)`/`CTF(±)`,
  este último siempre en 0 — confirmado que no existe), `Ingreso descarga`,
  `Costo carga`, `Descuento FD`, `Total` y `Monto a compensar`, entre otros.
  El FD homologado sale de la hoja `FD` de la misma planilla, no de releer
  el `SSCC_Desempeño_*`. La hoja
  queda **completa** (`A:AZ`, sin la columna `AY`, que
  la macro original tampoco escribe).
- `Calculo RE545` — la hoja hermana, también **completa** (`A:CE`): mismo
  traspaso desde `Medidores` (la energía se reparte entre las dos hojas según
  `Ventana_No_Completa`), las reservas por subasta (`AC:AU`), el resumen por
  central+ventana (`AW:BG`, una tabla de otro largo que se escribe al lado) y
  los Componentes 1 y 2 (`BI:CE`), hasta el `Monto a compensar`.

Las dos hojas de cálculo llevan además, arriba de los nombres de
columna, los encabezados de grupo con celdas combinadas del archivo real
(`Dia`, `Nombre`, `BESS`, `Prorratas (-)/(+)`, `FD`, `Subastas`, `FMA`,
`Componente 1`/`Componente 2` — `GRUPOS_CALCULO_E_COSTOS`/
`GRUPOS_CALCULO_RE545`); como la salida no reproduce la letra de Excel real
(solo el orden y el contenido), cada grupo cae en la columna que le toca en
**nuestro** orden, no en la del archivo original.

Las dos hojas de cálculo tampoco escriben sus auxiliares: `X` ("Ciclo") en
`Calculo E Costos`, que repetía el "Ciclo de Carga del mes" ya presente, y
`BL` (la columna sin nombre de la que `BM` saca su k-ésimo mayor) y `BR`
(repetía `T`, "Ventana de valorizacion") en `Calculo RE545`. Se siguen
calculando igual: sólo dejaron de ocupar una columna. Los intermedios que sí
dejan seguir el cálculo (las curvas monótonas, el `ranking cmg`, las
energías con FD, todo el paso a paso de los Componentes 1 y 2) se quedan.

Todas las hojas salen formateadas para leer (`formato.py`):
nombres de columna en negrita, panel inmovilizado bajo el encabezado, ancho
de columna según el contenido y separador de miles en las columnas
numéricas. Es sólo aspecto: no toca un valor.

Cada etapa abre **solo** lo que necesita, y lo que ya está escrito en la
planilla sale de ahí:

- Las hojas de cálculo leen `Medidores`, `Ofertas SSCC`, `CMg`, `Subastas` y
  `FD` **con una sola apertura del archivo** (antes cada hoja volvía a
  parsear el libro entero), y ya no reabren `cmg.xlsx`: el CMg es el de la
  hoja `CMg`, que es el mismo dato — así no puede pasar que los pagos usen
  un CMg distinto del que quedó en la foto de las entradas.
- Recalcular solo `PRORRATA_RETIROS` o el `Resumen` no lee las hojas de
  entrada ni `Centrales.xlsx` de más: esas dos hojas salen de las dos hojas
  de cálculo ya escritas (el `Resumen` sí necesita `Centrales.xlsx` para el
  propietario de cada central).
- `Centrales.xlsx` se abre una vez por corrida, aunque lo necesiten dos
  secciones, y la ventana no reabre las planillas grandes en cada repintado
  (se acuerda de lo leído mientras el archivo no cambie).

Las macros de Ofertas SSCC, CMg, FD y Subastas replicadas son solo las de
**carga** de esas hojas.

Validado con casos sintéticos (no con datos reales todavía): ver
`BITACORA.md` → "Pendientes abiertos" para lo que falta antes de dar por
cerrada cada etapa (validación contra un caso real y contra la planilla 11).
