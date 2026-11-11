# Plan: ventanas del Balance BESS más rápidas y fluidas

Auditoría del 2026-11-11 sobre `Balance_BESS.py` y lo que la ventana usa de
`Script/`. Este documento es para **quien ejecute los cambios** (otro modelo
u otra persona): no supone haber leído nada más que `desarrollo/REGLAS.md`.

- **Objetivo:** que la ventana abra más rápido, que no se congele al
  repintar y que el desplazamiento y los diálogos se sientan fluidos.
- **Lo que no cambia:** ningún resultado del cálculo, ningún archivo de
  salida, ningún texto ni botón que el usuario ve (salvo el cartel
  "Cargando…" de la Fase 2).
- **Alcance:** la ventana del BESS. El cálculo ya se optimizó (bitácora
  2026-11-03 (15), (16) y (19)) y no se toca. `Balance_CRA.py` queda fuera
  (ver §7).

---

## 1. Cómo trabajar este plan

1. Leer `desarrollo/REGLAS.md` entero y cumplirlo (bitácora, pruebas, no
   reescribir historial).
2. **Una fase = un commit.** Las fases son independientes salvo donde se
   dice; si una se complica, se deja sin hacer y se sigue con la otra.
3. Antes de empezar y después de cada fase, correr las dos cosas:

   ```bash
   xvfb-run -a python -m unittest discover
   ```

   ```bash
   xvfb-run -a python desarrollo/medir_ventana.py
   ```

   (En Windows, sin `xvfb-run -a`.) Las pruebas de ventana se saltean sin
   tkinter o sin pantalla: **si el total dice "skipped", no sirve como
   verificación de este plan** — hay que correrlas con pantalla.
4. Línea base de las pruebas al escribir este plan: **437 pruebas, OK, 0
   salteadas** (Python 3.12 + tkinter + xvfb). El total no puede bajar.
5. Cada fase trae su **criterio de aceptación**. Si no se cumple, la fase no
   está terminada; si al medir no hay mejora, se revierte y se anota en la
   bitácora por qué.
6. Lo que este plan marca como **decisión del usuario** (§6) no se hace sin
   preguntar.

### Entorno para las pruebas de ventana (Linux)

El Python del sistema puede no traer tkinter. Lo que funcionó acá:

```bash
apt-get install -y python3-tk xvfb
```

```bash
/usr/bin/python3.12 -m pip install --break-system-packages -r requirements.txt
```

y después correr todo con `/usr/bin/python3.12` (el que tiene tkinter).
`pyodbc` no hace falta para las pruebas.

---

## 2. Resultado de la auditoría

### 2.1 Lo que está bien (no tocar)

- El cálculo corre en un hilo aparte (`lanzar`), el registro llega por una
  cola y se escribe por tandas cada 100 ms, sin `update_idletasks` por línea.
- La rueda del mouse acumula los deltas chicos del touchpad y desplaza de a
  píxeles (`yscrollincrement=1`).
- "Ejecutar todo" arma su lista una vez y al tildar solo cambia valores.
- `revisar_estructura` cachea por versión de archivo (`_CACHE_HOJAS`) qué
  hojas tiene cada libro: no reabre las planillas en cada repintado.
- 437 pruebas, de ellas 5 archivos manejan la ventana real.

### 2.2 Mediciones (línea base)

Medido con `desarrollo/medir_ventana.py` en Linux, disco local, Python 3.12,
caso sintético completo. **No se pudo medir en Windows ni sobre una unidad de
red**, que es donde corre el programa: los números de abajo son el piso, y la
latencia de red está **simulada** (se dice dónde).

| Qué | Línea base | De dónde sale |
|---|---|---|
| `import Balance_BESS` (antes de que exista la ventana) | 490–760 ms | pandas 373 ms, openpyxl 102 ms, pyarrow 37 ms, el resto de `Script.nucleo` ~200 ms. tkinter: 6 ms |
| `main()` hasta el primer dibujo | 120–190 ms | incluye la primera revisión de la carpeta |
| Llamadas al disco por cada `revisar()` | **877** (caso con 60 Access) | 721 en `accdb_presentes`, 70 en `vigencia._mtime`, ~85 el resto |
| `revisar()` con disco local | 12–35 ms | — |
| `revisar()` simulando 3 ms por llamada | **2,9 s con la ventana congelada** | simulación; con 1 ms son ~1 s, con 10 ms ~9 s |
| Widgets del árbol | 396 (≈ 9 por fila) | se destruyen y se crean todos en cada repintado |
| Botón "Actualizar" (revisar + pintar) | 66 ms | — |
| Abrir "Ejecutar todo" | 42 ms | vuelve a revisar la carpeta entera |

Lectura: en disco local la ventana ya es rápida una vez abierta. Lo que el
usuario percibe como lento sale de **tres** cosas, en este orden:

1. **El arranque**: nada se dibuja hasta que terminan de importarse pandas,
   openpyxl y los ~60 módulos de `Script.nucleo`. En Windows, en frío y con
   antivirus, eso es varias veces lo medido acá.
2. **Cada repintado pregunta ~800 veces al disco**, en el hilo de la
   ventana. Sobre una unidad de red o una carpeta sincronizada eso son
   segundos de ventana congelada, y pasa al abrir, al cambiar de mes, al
   elegir carpeta, al apretar "Actualizar" y al terminar cada botón.
3. **El árbol se destruye y se vuelve a crear entero** (396 widgets) en cada
   repintado, y los diálogos se muestran mientras todavía se están armando.

### 2.3 Hallazgos

Gravedad: **A** = se nota siempre · **B** = se nota en red o en casos grandes
· **C** = prolijidad / robustez.

| # | Grav. | Dónde | Qué pasa |
|---|---|---|---|
| H1 | A | `Balance_BESS.py:116` | `from Script import nucleo` arriba del archivo: 0,5–0,8 s (más en Windows) sin ninguna ventana a la vista. Cualquier `import Script.nucleo.<algo>` ejecuta la fachada `nucleo/__init__.py`, que importa todo. |
| H2 | B | `Script/Subastas/Ofertas_Adjudicadas.py:216` `accdb_presentes` | Pregunta `is_file()` por cada nombre **posible** del mes (24 × días = 720–744), existan o no. Es el 82 % de las llamadas al disco de cada repintado. Una sola lectura de la carpeta da lo mismo. |
| H3 | B | `Script/nucleo/vigencia.py:79-82` | Un `stat` más por cada Access presente, para sacar el más nuevo. Con el mes completo son cientos. |
| H4 | B | `Balance_BESS.py:1213` `revisar` | Toda la revisión de la carpeta corre en el hilo de la ventana. Con una unidad de red lenta o desconectada la ventana queda "No responde". |
| H5 | A | `Balance_BESS.py:1033` `pintar_arbol` | Destruye y recrea las ~42 filas (396 widgets) aunque no haya cambiado nada: parpadeo en Windows. "Ejecutar todo" ya resolvió esto mismo (`armar`/`repintar`). |
| H6 | B | `Balance_BESS.py:472` `ajustar` | Atado a `<Configure>` del canvas **y** del contenedor; en cada evento recalcula `bbox("all")` y reconfigura el ancho, lo que dispara otro `<Configure>`. |
| H7 | B | `Balance_BESS.py:528-588` rueda | Un `yview_scroll` por cada evento de rueda. Un touchpad de precisión manda decenas por segundo y cada uno redibuja el árbol. |
| H8 | B | `Balance_BESS.py:1165` `vaciar_cola_log` | Vacía la cola entera de una vez: una ráfaga de miles de líneas congela la ventana mientras dura. Y `see("end")` siempre: no se puede leer hacia arriba durante una corrida. |
| H9 | A | Los 5 `Toplevel` y `Script/dialogos.py` | Se crean visibles y se van llenando: se ven armarse. Tampoco se centran sobre la ventana principal. |
| H10 | C | `Balance_BESS.py:2357` | "Ejecutar todo" crea el `Toplevel` y **después** llama a `revisar_estructura` sin `try`: si falla (unidad caída) queda una ventana vacía abierta y el error va a la consola. |
| H11 | C | `ventana_ejecutar_todo`, `ventana_antecedentes` | No son modales ni de instancia única: dos clics abren dos ventanas, cada una con la carpeta y el mes de cuando se abrió. |
| H12 | C | `Balance_BESS.py:2357, 2365` | Dentro de `ventana_ejecutar_todo` las variables `filas_arbol` y `COLOR_ESTADO` tapan a las de afuera (el `Frame` del árbol y el dict de colores del módulo). Hoy no rompe nada; confunde. |
| H13 | C | `usar_carpeta`, `aamm_cambiado`, arranque | `config.json` (compartido, posiblemente en red) se lee 3 veces y se escribe 2 veces con `fsync` por cada cambio de carpeta; en cada arranque se escribe aunque no haya cambiado nada. |
| H14 | C | `Script/nucleo/estructura.py` `revisar_estructura` | Repite el mismo `is_file()` 2–3 veces por archivo (`medidas_sae`, `cmg_csv`, `salida`, `consolidado`). |
| H15 | C | `Balance_BESS.py:1739-1966` | Ocho acciones casi idénticas (`caso_listo` + `validar_aamm` + `lanzar`): ~150 líneas repetidas. |
| H16 | C | `Balance_BESS.py:1625-1640` y docstring | `TITULO_DE_FUNCION` todavía dice "Balance_BESS.xlsx (entradas)" en el encabezado del registro; la planilla se llama `11_PAGOS_BESS_AAMM.xlsx` desde 2026-11-04. |
| H17 | C | `Script/nucleo/estructura.py:103` | `pd.ExcelFile(ruta)` sin motor: pandas avisa (`Pandas4Warning`) que el motor por defecto va a cambiar. `xlsx_rapido.nombres_de_hojas` ya da lo mismo sin abrir el libro. |
| H18 | C | `Balance_BESS.py:424-2795` | `main()` es una sola función de 2.370 líneas con todo adentro como clausuras. No bloquea este plan; **no se reestructura acá** (ver §5). |

---

## 3. Fases

Orden recomendado: de más ganancia y menos riesgo a menos ganancia y más
riesgo.

| Fase | Qué | Ataca | Riesgo |
|---|---|---|---|
| 1 | Menos llamadas al disco por repintado | H2, H3, H14 | bajo |
| 2 | La ventana aparece antes de importar lo pesado | H1 | medio |
| 3 | El árbol se actualiza sin destruirse | H5 | medio |
| 4 | Desplazamiento y registro | H6, H7, H8 | bajo |
| 5 | Diálogos y configuración | H9–H13, H16 | bajo |
| 6 | Revisar la carpeta fuera del hilo de la ventana | H4 | **alto** — decisión del usuario |
| 7 | Opcionales | H15, H17 | bajo |

---

### Fase 1 — Menos llamadas al disco por repintado

**Archivos:** `Script/Subastas/Ofertas_Adjudicadas.py`,
`Script/nucleo/estructura.py`, `Script/nucleo/vigencia.py`.

**1.1 `accdb_presentes` con una sola lectura de la carpeta.** Reemplazar el
`is_file()` por nombre con un `os.scandir` de `DB subastas/` y buscar cada
nombre posible en ese conjunto:

```python
def _archivos_de(carpeta):
    """{nombre: DirEntry} de los archivos de la carpeta. {} si no se puede leer."""
    try:
        with os.scandir(carpeta) as entradas:
            return {_clave(e.name): e for e in entradas if e.is_file()}
    except OSError:
        return {}
```

- `_clave` es `str.casefold` en Windows (`os.name == "nt"`) y la identidad en
  el resto: así el resultado es **el mismo** que daba `is_file()` en cada
  sistema.
- El resultado de `accdb_presentes` no cambia de forma: lista de
  `(dia, hora, carpeta / nombre)` en el orden de `nombres_del_periodo`.
- El módulo no importa `os` todavía: agregarlo.
- `accdb_presentes` también lo usa el consolidado de subastas (línea ~398):
  sale ganando sin tocarlo.

**1.2 La fecha del Access más nuevo sale de esa misma lectura.** Agregar en
`Ofertas_Adjudicadas.py` una función que devuelva también el `st_mtime` más
alto de los presentes (`entrada.stat().st_mtime`: en Windows no cuesta otra
llamada). En `revisar_estructura` guardarlo en
`rutas["accdb_subastas_mtime"]` (`None` si no hay ninguno), y en
`vigencia._fechas` usarlo cuando la clave esté; si no está, seguir como hoy
(`_mtime` de cada ruta). `rutas["accdb_subastas"]` se sigue llenando igual.

**1.3 No repetir el mismo `is_file()`** dentro de `revisar_estructura`:
guardar en una variable local el de `rutas["medidas_sae"]`,
`rutas["puntos_fallidos"]`, `rutas["cmg_csv"]`, `rutas["cmg"]`,
`rutas["salida"]` y `consolidado`, y usarla en las 2–3 veces que se pregunta.
Solo eso: no cambiar el orden ni el contenido de las filas.

**No hacer en esta fase:** cachear entre repintados qué archivos existen (el
botón "Actualizar" existe justamente para ver lo recién pegado), ni tocar
`_migrar_archivos_antiguos` o `sellos.py` (son ~25 llamadas entre todos).

**Pruebas nuevas** (`desarrollo/tests/test_accdb_presentes.py`):
- carpeta con algunos Access del período, uno de otro mes, un `.txt` y una
  subcarpeta con nombre de Access → devuelve exactamente los del período, en
  orden `(dia, hora)`;
- carpeta que no existe → `[]`;
- con `os.scandir` envuelto por un contador: **una** lectura de carpeta y
  ningún `os.stat` por nombre posible;
- la fecha más nueva coincide con `max(os.stat(r).st_mtime)` de los
  presentes.

**Criterio de aceptación** (`medir_ventana.py`, caso sintético):
- "llamadas al sistema de archivos" pasa de 877 a **≤ 120**;
- "con 3 ms por llamada" pasa de 2,9 s a **≤ 0,4 s**;
- las 437 pruebas siguen en verde, en especial `test_vigencia`,
  `test_orquestador`, `test_ventana_*`.

---

### Fase 2 — La ventana aparece antes de importar lo pesado

**Archivo:** `Balance_BESS.py` (y una prueba nueva).

Hoy: `import pandas…` (0,5 s o más) → recién ahí `tk.Tk()`. Después: la
ventana con un cartel "Cargando…" aparece enseguida y lo pesado se importa
con la ventana ya a la vista y respondiendo.

**2.1 Sacar `nucleo` de los imports de arriba.** Dejar
`from Script import config` (5 ms) y los de `arbol` y `dialogos` (no importan
nada pesado). `nucleo` pasa a ser un nombre del módulo que se llena en
`main()`:

```python
nucleo = None   # se carga en main(), con la ventana ya a la vista


def _cargar_nucleo():
    global nucleo
    if nucleo is None:
        from Script import nucleo as _nucleo
        nucleo = _nucleo
    return nucleo
```

Todo el resto del archivo sigue escribiendo `nucleo.algo`: son clausuras que
buscan el nombre al ejecutarse, cuando ya está cargado.

**2.2 `ARCHIVO_REGISTRO`** (línea 152) hoy se copia de `nucleo` al importar,
y una prueba lo lee como `app.ARCHIVO_REGISTRO`. Resolverlo con un
`__getattr__` de módulo (PEP 562) que llame a `_cargar_nucleo()`, y dentro
de `main()` usar `nucleo.ARCHIVO_REGISTRO` directo. No importar
`Script.nucleo.parametros` arriba: ejecuta la fachada entera.

**2.3 En `main()`, antes de armar nada:**

```python
root = tk.Tk()
root.title("Balance BESS / SSCC")
root.geometry("1180x820")
cartel = tk.Label(root, text="Cargando…", font=("Segoe UI", 11), fg=COLOR_NEUTRO)
cartel.pack(expand=True)
root.update()                      # la ventana ya se ve

if not _importar_con_la_ventana_viva(root):
    return                         # la cerraron mientras cargaba
cartel.destroy()
```

`_importar_con_la_ventana_viva(root)`: lanza `_cargar_nucleo` en un hilo y,
mientras el hilo vive, hace `root.update()` + `time.sleep(0.01)`. Así
`main()` sigue siendo **síncrona** (las pruebas hacen `app.main()` y esperan
la ventana armada al volver) y la ventana se puede mover y no queda "No
responde".
- Si el import levanta una excepción en el hilo: guardarla, y en el hilo
  principal mostrarla con `mostrar_error` y cerrar.
- Si `root.update()` levanta `tk.TclError` (cerraron la ventana): devolver
  `False`.
- `config.migrar_tokens()` y `leer_config()` van después de `root.update()`
  (leen de `__config__`, que puede estar en red).

**Plan B** si el import en un hilo diera cualquier problema en Windows:
importar en el hilo principal, justo después del `root.update()`. La ventana
aparece igual de rápido; queda quieta mientras carga. Anotar en la bitácora
cuál de las dos quedó.

**2.4 Prueba nueva** (`desarrollo/tests/test_arranque_liviano.py`), en un
subproceso para no depender de lo que ya importó la corrida de pruebas:

```python
codigo = "import sys, Balance_BESS; print('pandas' in sys.modules, 'openpyxl' in sys.modules)"
```

tiene que imprimir `False False`. (Importar `Balance_BESS` necesita tkinter:
saltearla con el mismo `_hay_ventana()` de las otras.)

**Criterio de aceptación:**
- `python -X importtime -c "import Balance_BESS"` no muestra `pandas`, y el
  acumulado de `Balance_BESS` es **≤ 60 ms** (hoy 490–760 ms);
- al abrir el programa, la ventana con "Cargando…" se ve de inmediato y
  después aparece el árbol;
- `python Balance_BESS.py`, cerrar la ventana durante "Cargando…": termina
  sin traceback;
- las pruebas de ventana pasan sin cambios.

**Riesgos a mirar:**
- Las pruebas parchean `app.leer_config`, `app.guardar_config`,
  `app.messagebox` y `app.threading`: esos nombres tienen que seguir
  existiendo en el módulo.
- `proceso.py` e `Indices_FMA.py` usan `ProcessPoolExecutor`; en Windows
  (`spawn`) los procesos hijos vuelven a importar `Balance_BESS.py`. Con este
  cambio ese import sale más barato; el `if __name__ == "__main__"` ya evita
  que abran otra ventana. Verificar con `test_paralelo` y `test_fma_paralelo`.
- `Script/nucleo/antecedentes.py` arma el recalculador siguiendo imports con
  `ast` **dentro de `Script/`**: esta fase no toca `Script/`, así que no lo
  afecta. Confirmar con `test_antecedentes_sae`.

**Para el PC del usuario** (no es código; anotarlo como pedido en la
bitácora): correr `python -X importtime Balance_BESS.py 2> importtime.txt` y
mirar si existe `Script/nucleo/__pycache__/`. Si la carpeta del programa es
de solo lectura o está en red, Python recompila los ~80 módulos en cada
arranque; se arregla con `python -m compileall Script` o instalando el
programa en disco local.

---

### Fase 3 — El árbol se actualiza sin destruirse

**Archivo:** `Balance_BESS.py` (`_fila_arbol`, `pintar_arbol`).

Mismo criterio que ya tiene "Ejecutar todo": la lista se arma una vez y
después solo cambian valores.

**3.1** `_fila_arbol` devuelve, además del `Frame`, las piezas que pueden
cambiar: la etiqueta del nombre, la etiqueta del estado y la celda de origen
(donde va el símbolo de aviso). Guardarlas en un dict por `id` de fila
(`filas_vivas`), junto con lo último que se pintó.

**3.2** `pintar_arbol(filas)`:
- calcula la **forma** del árbol:
  `tuple((f["id"], f["nivel"], prefijo, f.get("origen") and f["origen"]["id"]))`;
- si la forma es **igual** a la pintada: por cada fila cuyo contenido cambió
  (`etiqueta`, `estado`, `ruta`, `es_carpeta`, `aviso`), actualizar con
  `config` el texto, el color, la fuente (subrayado si hay ruta), el cursor y
  volver a atar el `<Button-1>` con la ruta nueva; crear o destruir solo el
  símbolo de aviso. Los botones no se tocan (dependen solo del `id`);
- si la forma **cambió** (apareció `Puntos_fallidos.xlsx`, cambió el mes y
  hay otras filas): reconstruir todo como hoy;
- el caso `filas == []` ("Selecciona una carpeta base.") y la fila de error
  siguen yendo por la reconstrucción completa;
- al final, como hoy: `if corriendo["activo"]: habilitar_botones(False)`.

**3.3** Al terminar un proceso (`terminar`) los botones quedan
deshabilitados y hoy se "rehabilitan" porque el árbol se recrea. Con filas
reutilizadas eso ya no pasa solo: llamar a `habilitar_botones(True)` en
`terminar` antes de `revisar()`.

**Pruebas nuevas** (en `desarrollo/tests/test_ventana_actualizar.py` o un
archivo hermano):
- apretar "Actualizar" dos veces sin cambiar nada → la lista de widgets del
  árbol (`[str(w) for w in self._todos(...)]`) es idéntica antes y después
  (mismo patrón que `test_tildar_no_vuelve_a_armar_la_lista`);
- pegar un archivo que cambia el **estado** de una fila sin cambiar la forma
  → el estado cambia y los widgets son los mismos;
- después de un proceso simulado (el `HiloEnElMismo` que ya usa
  `test_medidas_arranca_sin_advertencia`) los botones del árbol quedan en
  `normal`.

**Criterio de aceptación:** las tres pruebas nuevas y todas las de
`test_ventana_*`; "boton Actualizar" de `medir_ventana.py` baja (hoy 66 ms);
a la vista, apretar "Actualizar" ya no hace parpadear el árbol.

**Cuidado:** las pruebas encuentran una fila como
`etiqueta.master.master` (etiqueta → celda de estructura → fila). **No
cambiar la jerarquía de widgets de la fila** en esta fase.

---

### Fase 4 — Desplazamiento y registro

**Archivo:** `Balance_BESS.py`.

**4.1 `ajustar` partido en dos.** `canvas <Configure>` → solo ajustar el
ancho de la ventana interna, y **solo si cambió** (`event.width` distinto del
último). `contenedor <Configure>` → solo `scrollregion`. Aplicar lo mismo a
`ajustar_lista` de "Ejecutar todo".

**4.2 Un desplazamiento por cuadro, no por evento.** En
`_desplazar_ventana`, en vez de llamar a `canvas.yview_scroll` en cada
evento: sumar los píxeles a un pendiente y, si no hay uno programado,
programar **un** `root.after_idle(aplicar)` que haga un solo `yview_scroll`
con el total. Tiene que ser `after_idle` y no `after(N)`: las pruebas de
`test_ventana_rueda.py` mandan eventos, hacen `root.update()` y miran la
posición en el acto. Igual para `desplazar_lista` de "Ejecutar todo". El
registro (`Text`) se deja como está.

Esta mejora **solo se puede juzgar en Windows con un touchpad**. Si el
usuario no nota diferencia, se revierte: el código de hoy es más simple.

**4.3 El registro por tandas con tope.** `vaciar_cola_log(todo=False)`:
- saca como mucho ~400 mensajes por pasada; si quedaron, se reprograma en
  10 ms en vez de 100;
- baja al final (`see("end")`) solo si el registro **ya estaba al final**
  antes de insertar (`txt_log.yview()[1] >= 0.999`): así se puede leer hacia
  arriba durante una corrida;
- `terminar()` tiene que llamar a `vaciar_cola_log(todo=True)`: ahí se
  necesita la cola vacía de verdad **antes** de escribir la línea de cierre y
  de `guardar_registro()`, que guarda lo que hay en el widget.

No recortar el contenido del registro: `Registro_ultima_corrida.txt` sale de
ahí.

**Criterio de aceptación:** `test_ventana_rueda`, `test_ventana_registro` y
`test_ventana_ejecutar_todo` sin cambios; prueba nueva: 5.000 mensajes en la
cola + `terminar` → están los 5.000 en el archivo del registro, en orden.

---

### Fase 5 — Diálogos y configuración

**Archivos:** `Balance_BESS.py`, `Script/dialogos.py`.

**5.1 Los diálogos aparecen ya armados y centrados.** Un ayudante local:

```python
def presentar(top, modal=False, ancho=None, alto=None):
    """Muestra un Toplevel ya construido, centrado sobre la ventana principal."""
    top.update_idletasks()
    ancho = ancho or top.winfo_reqwidth()
    alto = alto or top.winfo_reqheight()
    x = root.winfo_rootx() + (root.winfo_width() - ancho) // 2
    y = root.winfo_rooty() + (root.winfo_height() - alto) // 3
    top.geometry(f"+{max(x, 0)}+{max(y, 0)}")
    top.deiconify()
    if modal:
        top.wait_visibility()
        top.grab_set()
```

Una ventana oculta todavía no tiene tamaño real (`winfo_width()` da 1): por
eso se usa el tamaño pedido (`winfo_reqwidth`), y los diálogos de tamaño fijo
pasan el suyo (`presentar(top, ancho=900, alto=640)`).

En cada diálogo: `top.withdraw()` apenas se crea, se construye entero, y al
final `presentar(top, modal=...)`. Son cinco: `ventana_aviso`,
`ventana_carpeta_del_periodo`, `ventana_nuevo_caso`, `ventana_ejecutar_todo`,
`ventana_antecedentes`; y `mostrar_error` en `Script/dialogos.py` (ahí con su
propia copia del centrado: ese módulo no depende de la ventana).

Orden obligatorio: **`deiconify` → `wait_visibility` → `grab_set`**. Un
`grab_set` sobre una ventana oculta levanta `TclError: grab failed: window
not viewable`. Los que hoy fijan tamaño (`top.geometry("900x640")`) lo
conservan: fijar el tamaño antes y solo la posición en `presentar`.
`ventana_carpeta_del_periodo.con_crear` ya usa `withdraw` por su cuenta: no
romperlo. `Herramientas/` tiene su propia copia de `dialogos`: no se toca.

**5.2 "Ejecutar todo" revisa antes de abrir (H10).** Mover
`nucleo.revisar_estructura(ruta, aamm)` **antes** de `tk.Toplevel(root)`,
dentro de un `try` que muestre el error con `mostrar_error` y vuelva (como ya
hace `ventana_antecedentes`).

**5.3 Una sola instancia (H11).** Para "Ejecutar todo" y "Generar
antecedentes SAE": guardar el `Toplevel` abierto; si ya hay uno vivo
(`winfo_exists()`), traerlo al frente (`lift()` + `focus_set()`) en vez de
crear otro.

**5.4 Nombres que se tapan (H12).** En `ventana_ejecutar_todo`:
`filas_arbol` → `filas_plan` y `COLOR_ESTADO` → `COLOR_ESTADO_PLAN`.

**5.5 Menos idas a `config.json` (H13).** Sin cachear nada (el archivo es
compartido con otros programas):
- `recordar_carpeta` no escribe si esa carpeta ya está anotada para ese
  período;
- `usar_carpeta` junta sus dos escrituras (`carpeta_base` y
  `carpetas_por_periodo`) en una sola llamada a `guardar_config`.

`leer_config` y `guardar_config` siguen siendo las dos únicas puertas (las
pruebas las parchean).

**5.6 Texto viejo (H16).** En `TITULO_DE_FUNCION`, "Balance_BESS.xlsx" →
"Planilla 11_PAGOS_BESS" en las tres entradas que lo dicen, y actualizar el
docstring del módulo. Revisar antes con `grep` que ninguna prueba compare ese
texto.

**Criterio de aceptación:** todas las pruebas; prueba nueva: dos `invoke()`
de "Ejecutar todo" dejan **un** `Toplevel`; prueba nueva: con
`nucleo.revisar_estructura` parcheado para que falle, "Ejecutar todo" no deja
ningún `Toplevel` abierto. A la vista: los diálogos aparecen de una vez,
sobre la ventana principal.

---

### Fase 6 — Revisar la carpeta fuera del hilo de la ventana

**Decisión del usuario (§6).** Es la única fase que cambia **cuándo** pasan
las cosas, y la de más riesgo. Hacerla solo si, con la Fase 1 hecha, el
usuario sigue viendo la ventana congelarse (típico: carpeta del caso en una
unidad de red lenta o que se desconecta). Si trabaja en disco local, **no
hacerla**.

**Diseño:**
- `revisar()` deja de devolver con el árbol pintado: **pide** una revisión.
  Toma `ruta` y `aamm` en el hilo principal, suma un número de pedido y lanza
  un hilo que hace `Path(ruta).is_dir()` + `nucleo.revisar_estructura(...)` y
  entrega `(numero, resultado o error)` al hilo principal.
- Al recibirlo: si `numero` no es el último pedido, se descarta (llegó
  tarde). Si es el último, se pinta como hoy (`pintar_arbol`, `var_estado`,
  `btn_abrir_salida`, color de `lbl_base`).
- Una revisión a la vez: si se pide otra mientras corre una, se anota y se
  lanza al terminar la que está (con los valores más nuevos).
- Mientras se revisa: `var_estado.set("Revisando la carpeta…")`. El árbol
  anterior queda a la vista (con la Fase 3 no parpadea).
- Una variable Tk con nombre, `tk.BooleanVar(root, name="revisando")`, dice
  si hay una revisión en curso. Es la forma de que las pruebas esperen sin
  meterse en las clausuras:

  ```python
  def _esperar_revision(self):
      while self.root.getvar("revisando") in (1, "1", True):
          self.root.update()
          time.sleep(0.005)
  ```

  Las pruebas de ventana que miran el árbol después de una acción tienen que
  llamar a ese ayudante (va en `TestVentanaActualizar`). Es un cambio
  legítimo de las pruebas: lo que comprueban no cambia.
- `Script/nucleo/estructura.py`: `_con_cache` muta un dict de módulo; ponerle
  un `threading.Lock`, porque ahora puede revisarse desde dos hilos (el de
  la revisión y el del proceso, que llama a `planificar`).
- Entregar el resultado con la misma cola que ya sondea el registro, o con
  `root.after(0, ...)` como hoy hacen `lanzar` y `abrir_origen`. Elegir una y
  decirlo en la bitácora.

**Quiénes llaman a `revisar()` y qué esperan** (revisar uno por uno):
`usar_carpeta` (después llama a `faltan_subcarpetas`, que no depende del
árbol), `actualizar_vista`, `volver_al_periodo_anterior`, `aamm_cambiado`
(cuatro lugares), `terminar` (después abre el cuadro "Listo"),
`crear_carpetas_en`, y el arranque. Ninguno usa el resultado de `revisar()`;
todos toleran que el árbol se pinte un instante después.

`ventana_ejecutar_todo` y `ventana_antecedentes` llaman a
`nucleo.revisar_estructura` directo y **se dejan síncronas**: con la Fase 1
esa llamada ya es barata.

**Criterio de aceptación:**
- prueba nueva: con `nucleo.revisar_estructura` parcheado para tardar 0,5 s,
  `actualizar.invoke()` vuelve en menos de 50 ms y, tras esperar, el árbol
  está pintado;
- prueba nueva: dos pedidos seguidos con resultados distintos → queda
  pintado el del **último**;
- todas las pruebas de ventana, con `_esperar_revision()` donde haga falta;
- manual: con la carpeta del caso en una unidad de red, desconectarla y
  apretar "Actualizar": la ventana se sigue pudiendo mover.

---

### Fase 7 — Opcionales (solo si sobra tiempo; cada uno en su commit)

**7.1 Nombres de hoja sin abrir el libro (H17).** En `estructura.hojas_de`,
para `.xlsx`/`.xlsm` usar `xlsx_rapido.nombres_de_hojas(ruta)` y, si falla,
caer a `pd.ExcelFile` como hoy (con `engine="openpyxl"` explícito para
apagar el aviso). Para `hojas_con_datos` se puede leer el `<dimension ref>`
de los primeros 2 KB de cada hoja (`xlsx_rapido._hojas` da la parte de cada
una); medido: 0,7 ms contra 13 ms en un libro de 64 MB. **Solo** con una
prueba que compare contra el camino actual en: libro escrito por pandas,
libro escrito por el programa (`relleno_rapido`), hoja vacía de una sola
celda, hoja sin `<dimension>`. Si hay cualquier diferencia, no se hace: de
esto depende qué filas salen OK/PENDIENTE.

**7.2 Acciones repetidas (H15).** Un ayudante
`_accion(funcion, etiqueta_salida, **extra)` que haga `caso_listo` +
`validar_aamm` + `lanzar`, y las ocho funciones (`traer_cmg_15min`,
`traer_fd`, `traer_subastas`, `consolidar_subastas`, `generar_csv_sae`,
`generar_reporte`, `generar_cmg`, `actualizar_medidas_sae`) quedan en una
línea cada una. Sin cambio de comportamiento; no mejora la velocidad.

---

## 4. Verificación final (después de la última fase hecha)

1. `python -m py_compile Balance_BESS.py Balance_CRA.py Script/*.py Script/nucleo/*.py Script/*/*.py`
2. `xvfb-run -a python -m unittest discover` → OK, **0 salteadas**, total ≥
   437 + las pruebas nuevas.
3. `xvfb-run -a python desarrollo/medir_ventana.py` → pegar la salida en la
   bitácora, al lado de la línea base de §2.2.
4. Tabla de antes/después en la entrada de la bitácora.
5. Pedirle al usuario que pruebe en Windows lo que acá no se puede ver:
   arranque en frío, parpadeo del árbol, rueda con touchpad, diálogos, y una
   corrida completa de "Ejecutar todo" (que el registro y
   `Registro_ultima_corrida.txt` salgan completos). Dejarlo en "Pendientes
   abiertos" hasta que conteste.

Metas globales con las fases 1 a 5 hechas:

| Qué | Hoy | Meta |
|---|---|---|
| Tiempo hasta ver la ventana | import completo (≥ 0,5 s; más en Windows) | lo que tarda `tk.Tk()` (< 0,1 s acá) |
| Llamadas al disco por repintado | 877 | ≤ 120 |
| Repintado con red simulada (3 ms) | 2,9 s | ≤ 0,4 s |
| Widgets recreados al "Actualizar" sin cambios | 396 | 0 |
| Diálogos | se ven armarse | aparecen armados y centrados |

---

## 5. Qué no hacer

- **No tocar el cálculo ni la escritura de planillas** (`ecostos*`,
  `re545*`, `escritura.py`, `relleno_rapido.py`, `xlsx_rapido.pegar_hojas`,
  `proceso.py`): ya se optimizaron y cada cambio ahí exige comparar celda por
  celda.
- **No hacer `import` perezosos dentro de `Script/nucleo/`.**
  `antecedentes.py` arma el recalculador que se entrega a los coordinados
  siguiendo los imports de esos módulos; moverlos puede dejar un módulo
  afuera sin que ninguna prueba de ventana lo note.
- **No reestructurar `main()`** en clases o módulos (H18). Es una deuda
  real, pero mezclarla con este plan hace imposible revisar qué cambió.
- **No cambiar la jerarquía de widgets de las filas** ni pasar el árbol a
  `ttk.Treeview`: los botones por fila y las columnas rectas son pedidos
  explícitos del usuario, y las pruebas dependen de esa jerarquía.
- **No cachear entre repintados qué archivos existen.**
- **No activar "DPI awareness" en Windows.** Dejaría el texto más nítido,
  pero los anchos del árbol están en píxeles fijos (`ANCHO_ESTRUCTURA`…) y
  con la pantalla al 125–150 % los nombres se cortarían.
- **No quitar el cuadro "Listo"** del final de cada proceso ni ninguna otra
  ventana de confirmación: son decisiones del usuario.
- No saltear, desactivar ni aflojar pruebas para llegar al verde.

---

## 6. Decisiones del usuario

| # | Pregunta | Si no contesta |
|---|---|---|
| D1 | ¿La carpeta del caso y `__config__` están en una unidad de red o sincronizada (OneDrive)? | Hacer fases 1–5; dejar la 6 sin hacer. |
| D2 | ¿Hacer la Fase 6 (revisar en segundo plano)? Cambia que el árbol se pinta un instante después de la acción. | No hacerla. |
| D3 | El cartel "Cargando…" de la Fase 2: ¿ese texto u otro? | "Cargando…". |
| D4 | ¿Que los diálogos queden centrados sobre la ventana principal (Fase 5.1)? Hoy salen donde Windows los ponga. | Centrados. |
| D5 | ¿Repetir las fases 3 y 4 en `Balance_CRA.py`? | No: fuera de alcance. |

---

## 7. Fuera de alcance, para otra vez

- **`Balance_CRA.py`** tiene los mismos patrones en chico: revisa la carpeta
  en cada `<FocusOut>` y cada Enter del AAMM, destruye y recrea el árbol, y
  manda cada línea del registro con un `root.after(0, …)` propio en vez de
  una cola. Lo de las fases 3 y 4.3 se puede copiar tal cual.
- **Arranque real más corto** (no solo percibido): para que el árbol se
  pinte sin esperar a pandas haría falta que `estructura.py`, `rutas.py`,
  `origenes.py`, `vigencia.py` y `orquestador.py` no lo importen ni directa
  ni indirectamente (hoy llegan por `externos.py`). Es una reorganización de
  `Script/nucleo/` que choca con el recalculador de antecedentes; evaluarla
  solo si, con la Fase 2 hecha, el `importtime` del PC del usuario muestra
  varios segundos.
