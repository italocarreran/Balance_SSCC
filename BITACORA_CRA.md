# BITACORA_CRA.md — registro de sesiones del CRA (`Balance_CRA.py` / `Script/cra/`)

Solo se agrega. Nunca se edita ni borra una entrada vieja. La única
excepción es la sección "Pendientes abiertos", que sí se edita porque es un
estado, no un historial.

Desde 2026-10-06 la bitácora está separada en dos: esta, del CRA (`Balance_CRA.py` / `Script/cra/`), y
`BITACORA_BESS.md` del BESS. Lo que toca a los dos (reglas, estructura común,
`Script/nucleo` cuando lo usan ambos) se anota en las dos.

---

## Pendientes abiertos

Referencia: `docs/Trazabilidad_CRA_Periodo_Generico_v5_Auditoria_Formulas.md`. Nada de esto
se adivinó: lo que depende de una respuesta quedó como constante en
`Script/cra/parametros.py` o como hoja PENDIENTE en el árbol.

- **Maestros que el Excel tiene escritos a mano (bloquean CO, FP y
  `CÁLCULO_CRA`).** La v5 confirma por fórmula `FP → CO → CO_Barra_Propia`,
  pero esas fórmulas consultan tablas que no son fórmula ni vienen de
  ningún archivo de entrada:
  - `CO!T:W` — lista de configuraciones con su barra/embalse
    (`CO!G`, `H`, `J` la usan);
  - `FP!G:H` — `BARRA BALANCE ↔ BARRA POLITICA` (homologa la barra del
    CO con la del `fp_`);
  - `CO!F2` — el día con cambio de hora (la hora día llega a 25);
  - `CO!AA:AB` — unidad → embalse; `CO!X:Y` — lo usa `CONDICION_EMBALSE!H`;
  - `dict_SCCO` (`SC y CO!AT8:AV49`) — Central → Config. InfoTécnica →
    Nombre CRA; en agosto le falta `ANGOSTURA-3` (está en la lista de
    embalses) y todos los `PE-*`;
  - unidades candidatas (`CÁLCULO_CRA!AQ:AS`), `RENDIMIENTOS`.
  - (`EMPRESAS` ya está: hoja `empresas` de `centrales_cra.xlsx`.)
  Hay que decidir de dónde los lee Python (¿un `Maestros_CRA.xlsx` en el
  caso, como el `Centrales.xlsx` del BESS?).
- **CO: de dónde sale `CO!L`** — resuelto en la v5 (6.4): `L = I × K`,
  CO base × FP por barra + día + hora. Falta solo lo de los maestros.
- **TC:** pendiente (usuario, 2026-09-30).
- **`fp_` / `cvar_cra_` desde la red** (`Script/Politicas`): confirmado por
  el usuario (2026-10-02) que `centrales_cra.xlsx` va en `Auxiliares/`, que
  la carpeta de `progdiar_SEN` es el año de dos dígitos y que, por ahora,
  si falta la política de un día se corta. Falta solo correrlo una vez en
  Windows contra `nas-cen1`.
- **RENDIMIENTOS, unidades candidatas, COTAS,
  CONDICION_EMBALSE:** fuera de foco por ahora (usuario: "enfocarnos en lo
  que tenemos").
- **FD_CPF/CSF/CTF:** la v5 (6.7) confirma la forma: clave en `A`
  (Unidad + Fecha/Hora horaria) y el FD en `I`/`H`/`I`, igual que las hojas
  horarias del `SSCC_Desempeño`, con todas las unidades (82.584 / 95.976 /
  369.024 filas en agosto). Resuelto; falta validar contra un archivo
  real.
- **Nombres de unidad:** el Reporte_CRA real trae `CANUTILLAR-1`, el
  `CÁLCULO_CRA` usa `CANUTILLAR_U1`: el diccionario SC/CO → Nombre CRA
  hace falta para `CÁLCULO_CRA` (no para cargar la hoja).
- Falta un `Cálculo_SobrecostosSSCC_*.xlsm` real para validar la parte SC.
- **Lista de centrales de embalse** (`CENTRALES_EMBALSE`): hoy está en el
  código, copiada del `Actualiza_SC_CO.py` del usuario. Evaluar si pasa a
  un maestro (junto con EMPRESAS / unidades candidatas) cuando se definan.
- **PRORRATA_RETIROS:** solo el bloque 1 (matriz período × empresa, en
  orden alfabético). Los cuadros 2 y 3 esperan a `CÁLCULO_CRA`.
- **Cambio de hora (92/100 períodos):** la `Clave_Bloque` de SC replica
  el `*96` del Excel y la cobertura de la prorrata avisa si no son
  `días × 96`; revisar con un mes con cambio de hora.
- Validar todas las hojas contra los archivos reales de un período (hoy
  solo hay pruebas sintéticas).
- Abrir `Balance_CRA.py` en Windows (acá no hay tkinter: solo se compiló
  y pasó `pyflakes`).

---

## 2026-09-30 — Inicio del CRA: `Balance_CRA.py` y `Script/cra/` (hojas de entrada)

**Pedido del usuario:** pasar a Python una segunda planilla,
`5_REMUNERACIÓN_CRA_<AAMM>_Definitivo.xlsx`, en el mismo repositorio que el
BESS, con su propia ventana `Balance_CRA.py` y compartiendo código; empezar
por las hojas con sus entradas resueltas y preguntar lo que no sea obvio.
Entregó la trazabilidad, que quedó versionada en
`docs/Trazabilidad_CRA_Periodo_Generico_v3.md`.

**Qué se hizo:**

- `Script/cra/`, paquete hermano de `nucleo` (el del BESS). Usa de
  `nucleo` solo las piezas genéricas (`normalizar`, `ErrorEntrada`,
  `formatear_hoja`); `nucleo` no importa nada de `cra`.
- Las cuatro hojas de entrada cuya carga está definida en la trazabilidad,
  escritas en `Balance_CRA.xlsx` (una fila por hoja en la ventana, cada una
  con su botón **Actualizar**; las que no se tocan se conservan):
  - `ENERGIA` (traz. 6.1): B:F, `HORADIA = HORA(G)`,
    `PERIODO DE CALCULO = J*4 + (MINUTO(G)+15)/15`, kWhD/kWhR. Se valida
    que Año/Mes de todas las filas sean los del período.
  - `FP` (6.4.1) y `CO` (6.4.2): copia de sus bloques A2:D.
  - `SC y CO` (6.6): Reporte_CRA (CO) apilado arriba de Sobrecostos (SC),
    11 columnas; las sumas de tres componentes y las dos `Clave_Bloque`
    replican las fórmulas entregadas (incluido el `*96`). CPF/CSF/CTF
    quedan en memoria (`participacion_por_servicio()`), no se escriben.
- Lectura por **letra y fila real de Excel** con openpyxl
  (`lectura.leer_columnas`), porque así define la trazabilidad cada carga.
- `Script/arbol.py`: los prefijos del árbol (`├──`/`└──`) salieron de
  `Balance_BESS.py` para que las dos ventanas usen la misma función.

**Qué NO se hizo, a propósito:** TC, COTAS, RENDIMIENTOS,
CONDICION_EMBALSE, FD_*, EMPRESAS, el `Neto` de ENERGIA, `CO_Barra_Propia`
y el cálculo (`CÁLCULO_CRA` → `PRORRATA_RETIROS` → `RESUMEN`). Aparecen en el
árbol como PENDIENTE con el motivo; las preguntas están en "Pendientes
abiertos → CRA".

**Verificación:** `py_compile` de todo, 160 pruebas (9 nuevas en
`tests/test_cra_entradas.py`, con archivos sintéticos armados con la forma
de la trazabilidad). La ventana no se pudo abrir: el entorno no tiene
tkinter.

---

## 2026-09-30 (2) — CRA: respuestas del usuario, FD_* y matriz de PRORRATA_RETIROS

**Respuestas del usuario** a las preguntas de la entrada anterior:

1. La hoja de `Cálculo_SobrecostosSSCC_*.xlsm` es `SOBRECOSTOS`.
   `Reporte_CRA` es un **.csv** (`Reporte_CRA_15min_AAMM`): se toma tal
   cual (si llegara en Excel, la primera hoja).
2. La estructura de carpetas (una por hoja) queda.
3. `fp_` lleva siempre el AAMM → se busca `fp_<AAMM>*`. El de energía es
   `..._Hidro_Agosto2026` → se busca con el mes en palabras (también
   `Setiembre`).
4. ENERGIA `K`, `L`, `M` son fórmulas que no están en la trazabilidad →
   pregunta abierta (ver pendientes).
5. No quedó claro qué se preguntaba sobre `CO_Barra_Propia` → se
   reformuló en pendientes.
6. TC pendiente. 7. Enfocarse en lo que hay.
8. FD y prorrata de retiros son las mismas fuentes del BESS → se
   reutiliza su código.

**Qué se hizo:**

- `lectura.py` lee **CSV** con la misma regla de letra/fila (separador
  detectado, coma decimal con `;`, fechas en texto sin mes-día-año), la
  hoja por posición (`0` = primera) y encuentra la hoja aunque cambien
  mayúsculas/tildes.
- `SC y CO` avisa (no corta) si hay fechas fuera del período.
- Hojas nuevas en `Balance_CRA.xlsx`, en `fuentes_bess.py`:
  - `FD_CPF`, `FD_CSF`, `FD_CTF`: copia de `CPF/CSF/CTF Horario` del
    `SSCC_Desempeño_*` (rangos de `Script/Fd/Desempeno_Horario.py`), con
    sus encabezados reales. Carpeta `FD/` del caso, con botón **Traer**
    que usa la misma función del BESS (`Indicadores_DCO.traer_fd`).
  - `PRORRATA_RETIROS`, bloque 1: la matriz período × empresa desde
    `Prorrata 15min`, leída con `nucleo.prorrata_retiros`. Avisa si los
    períodos no son exactamente `1..días×96`.

**Verificación:** `py_compile`, `pyflakes` del CRA, 164 pruebas (13 del
CRA). La ventana sigue sin poder abrirse acá (sin tkinter).

---

## 2026-09-30 (3) — CRA: `Reporte_CRA` real validado; un solo mes por archivo

**Respuestas del usuario:** ENERGIA sigue solo con datos (las fórmulas de
K/L/M las revisa él); CO se ve después; Sobrecostos trae **un solo mes**;
entregó un `Reporte_CRA_15min_2512.csv` real; sobre la forma de `FD_*`
dice que debería estar en la trazabilidad (no lo está: ver pendientes).

**Qué se hizo:**

- El CSV real quedó en `docs/Reporte_CRA_15min_2512_real.csv` y su
  estructura en `docs/Estructura_Archivos_Reales.md` §D.1. Con el lector
  tal como estaba se leyó bien: 6.867 filas, `Clave_Bloque` `1#38`..`16#92`,
  y las participaciones calculadas (`CPF = CPF(+)+CPF(-)`, etc.)
  coinciden con las columnas P/Q/R que trae el propio archivo (diferencia
  máxima 1e-9). Tres pruebas nuevas contra ese archivo.
- Como cada archivo de SC y CO es de un solo mes, una fecha fuera del
  período ya no es un aviso sino `ErrorEntrada` (archivo equivocado o
  fecha mal leída).
- Los enteros del CSV quedan enteros (`2512`, no `2512.0`).

**Verificación:** 167 pruebas (16 del CRA), `pyflakes` del CRA limpio.

---

## 2026-09-30 (4) — CRA: `SC y CO` solo con centrales de embalse, SC arriba

**Pedido del usuario:** en `SC y CO` hay que filtrar las centrales de
embalse; entregó su `Actualiza_SC_CO.py` (xlwings, sobre la planilla) para
ver la lógica, aclarando que el origen de CO ahí es otro (`Calculo_CO`) y
no hay que fijarse en eso.

**Qué se tomó de ese script (y qué no):**

- La lista `CENTRALES_EMBALSE` (27 unidades, nombre exacto del origen) →
  `Script/cra/parametros.py`. Se filtra en las dos fuentes, comparando
  normalizado.
- Los avisos: centrales de la lista sin ninguna fila, y centrales que
  terminan en `-número` sin estar en la lista (posible unidad nueva).
- El **orden** de la hoja: SC arriba y CO abajo. Estaba al revés.
- **No** se tomó: el origen `Calculo_CO` / hoja `PRORRATA CO` ni sus
  letras (`AU:BJ`, `BM:CB` de 16 columnas de prorrata), porque el usuario
  dijo que en este caso el origen es otro; se sigue la trazabilidad
  (`Reporte_CRA` + las sumas de 3 componentes de `SOBRECOSTOS`).
  Tampoco la Clave Año_Mes copiada de SC a CO: `Reporte_CRA` ya la trae.

**Resultado con el CSV real:** 5.686 de 6.867 filas (salen `PE-TOLPANSUR`
y `PE-SANGABRIEL`); aviso de 13 centrales de la lista sin filas en el mes
(RALCO, COLBUN, PANGUE, ANTUCO, ANGOSTURA, CIPRESES-2, RAPEL-5).

**Verificación:** 167 pruebas, `pyflakes` limpio.

---

## 2026-09-30 (5) — CRA: trazabilidad v5 (auditoría de fórmulas); ENERGIA con `Neto`

**Entrega del usuario:** `Trazabilidad_CRA_Periodo_Generico_v5_Auditoria_
Formulas.md`, que **reemplaza** a la v3 en `docs/` (se movió con `git mv`
para conservar la historia; las referencias se actualizaron salvo en las
entradas viejas de esta bitácora, que no se editan).

**Qué trae de nuevo, y qué se hizo con cada cosa:**

- **ENERGIA `K`, `L`, `M`** (6.1) → implementado: `corregir_energia()`
  transcribe rama por rama las dos fórmulas `SI(...)` y `Neto = L + K`;
  celda vacía = 0 (como la fórmula). Se comprobó que coinciden con la
  "tabla funcional" del documento en una grilla de casos (negativos,
  cero, positivos, iguales). La hoja `ENERGIA` ahora trae
  `kWhD corregido`, `kWhR corregido` y `Neto`.
- **Clave de `FD_*`** (6.7) → cada hoja FD empieza con `Fecha Hora`
  (= `Fecha + Hora` horas), que junto con `Unidad` es la clave del libro
  (`A = D & (B + TIMEVALUE(C & " :00"))`). No se copia el texto pegado del
  Excel (número de serie como texto): el cruce será por las dos columnas.
- **Control `AI` de `SC y CO`** → aviso si se repite Tipo + Unidad +
  Clave_Bloque.
- **`FP → CO → CO_Barra_Propia`, `CONDICION_EMBALSE!A:I`, `dict_SCCO`** →
  confirmados por fórmula pero dependen de tablas escritas a mano en el
  libro; quedaron como pregunta en "Pendientes abiertos → CRA".
- Libro en `calcMode = manual` (13.6): al validar contra el Excel, hacerlo
  con una copia recalculada.

**Verificación:** 18 pruebas del CRA, suite completa en verde.

---

## 2026-10-02 — CRA: `fp_<AAMM>` y `cvar_cra_<AAMM>` los arma el programa

**Pedido del usuario:** con `entradas_sscc.py` + `archivo_de_configuracion.yaml`
(el script que se usa hoy para preparar las entradas del CRA y de otros
cálculos) como base solo para rutas y lógica, armar `fp_AAMM` y
`cvar_cra_AAMM` integrado al proyecto.

**Qué se tomó del script (solo el paso `costo_variable`):** las rutas
(`progdiar_SEN/<AA>/PO*.csv|xlsx`, `DPID/.../PID_CDC_<HH>/PO*_<HH>.csv`,
`CMgReales/<AAMM>/Politicas/PRG*_<HH>.xlsx`), la lectura de cada archivo
(despivoteo, tercera hoja B:Z para FP), la regla de reemplazo PID
(horas ≥ HH, de la 1 a la 23) y el filtro con `centrales_cra.xlsx`.
Nombres de salida iguales: `fp_<AAMM>_1_<N>.xlsx`,
`cvar_cra_<AAMM>_1_<N>.xlsx`. **No** se tomó: el YAML (el período sale de
la ventana; las rutas son constantes del módulo), las carpetas
`input/output` por versión, el hack `dia == 41`, ni el resto de los pasos.

**Cambios respecto del original:** la columna de configuración y la de
barra se toman por posición (el encabezado de la primera es la fecha, el
de la segunda trae espacios); si falta la política de un día se corta
con la lista completa (el original fallaba en el primero y lo tapaba con
un `except` general); se avisan PID con programa pero sin política y
configuraciones de `centrales_cra` sin datos.

**Integración:** paquete hermano `Script/Politicas/` (sin `nucleo`,
`ErrorPoliticas`), envuelto por `cra.generar_politicas`; carpeta nueva
`Auxiliares/` en el caso CRA con `centrales_cra.xlsx`; botones
**Generar** en las filas de `FP/` y `CO/`. Se comprobó que la hoja `FP` y
la hoja `CO` leen tal cual lo que se genera.

**Verificación:** 175 pruebas (6 nuevas en `tests/test_cra_politicas.py`,
con un árbol de red falso: mes completo, una reprogramación PID, un PID
sin política, un día faltante, dos archivos del período). Las rutas
reales de `nas-cen1` no se pudieron probar desde acá.

---

## 2026-10-02 (2) — Confirmaciones de `fp_`/`cvar_cra_` y fusión a `main`

El usuario confirmó las tres decisiones de la entrada anterior
(`Auxiliares/centrales_cra.xlsx`; año de dos dígitos en `progdiar_SEN`;
cortar si falta un día) y pidió fusionar a `main`. Sin cambios de código.

---

## 2026-10-02 (3) — `centrales_cra.xlsx` con formato nuevo; FD solo de las unidades del CRA

**Pedido del usuario:** `centrales_cra.xlsx` pasa a un formato nuevo, que
además trae la homologación con los FD; adaptar el script y que al traer
los FD se queden solo los de esas centrales "y no todo el mamarracho".

**Formato nuevo** (archivo real en `docs/centrales_cra_real.xlsx`,
documentado en `docs/Estructura_Archivos_Reales.md` §D.2): tres hojas
`centrales_cra`, `empresas`, `diccionario`, con título arriba y
encabezados en la fila 8, columna B.

**Qué se hizo:**

- `Script/cra/maestros.py`: un solo lector para las tres hojas (busca el
  encabezado por texto; separa las listas `;` del diccionario).
- `cvar_cra`: `Costos_Variables.construir()` ya no lee el Excel; recibe
  la lista `configuraciones` que lee `cra.maestros` (hoja
  `centrales_cra`). Sin el maestro, `Generar` de `CO/` se detiene.
- `FD_CPF/CSF/CTF`: solo las filas cuya `Unidad` está en esa columna del
  diccionario (comparación normalizada); avisa las unidades del
  diccionario sin filas. El maestro pasa a ser obligatorio para esas
  hojas.
- `EMPRESAS` queda resuelto como maestro (`leer_empresas`), todavía sin
  uso hasta `CÁLCULO_CRA`.
- **Rendimiento:** el archivo real tiene un `styles.xml` de ~11 MB y
  openpyxl tardaba ~9 s por apertura. `maestros.py` lee los valores
  directo del XML del zip (sin estilos; openpyxl de respaldo): < 1 s, y
  se comprobó que da exactamente lo mismo que openpyxl en las tres hojas.

**Verificación:** suite completa en verde; pruebas nuevas contra el
maestro real (56 configuraciones, 81 empresas, 27/25/27 unidades FD) y
del filtro de FD (incluida una unidad escrita con otro formato que igual
entra, y una del diccionario sin filas que se avisa).

---

## 2026-10-02 (4) — Diccionario de FD con una sola columna

El usuario cambió la hoja `diccionario` de `centrales_cra.xlsx`: en vez de
`FD_CPF`/`FD_CSF`/`FD_CTF` deja una sola columna `FD`, porque eran
iguales para los tres. `leer_unidades_fd(ruta)` ya no recibe la sección:
la misma lista (27 unidades) filtra las tres hojas `FD_*`. Se reemplazó
`docs/centrales_cra_real.xlsx` por el archivo nuevo
(`centrales_cra_diccionario_FD_ordenado.xlsx`). Único efecto visible: en
el archivo anterior ANTUCO no tenía CSF; ahora entra a las tres y, si no
tiene filas en `CSF Horario`, `FD_CSF` lo avisa. Suite en verde.

## 2026-10-06 — Bitácora separada en `BITACORA_BESS.md` y `BITACORA_CRA.md`

Pedido del usuario. `BITACORA.md` se partió moviendo el texto **tal
cual**: los pendientes del CRA y las entradas desde 2026-09-30 quedaron
acá; el resto en `BITACORA_BESS.md`. `REGLAS.md` ahora pide leer la
bitácora de la herramienta que se toca (las dos si el cambio toca código
compartido) y anotar en la que corresponda. En el CRA no cambió código en
esta sesión (los arreglos del día fueron del BESS: ver su bitácora).
