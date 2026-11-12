# -*- coding: utf-8 -*-
r"""
Sacar de PI AF cualquier dato de una o varias centrales (generacion,
SoC, potencia, lo que tenga el elemento) en un rango de tiempo, cada
15 minutos o cada hora, a un Excel.

Usa la conexion, la busqueda y la lista de elementos SAE* de
Buscar_Rutas_PI.py: los dos archivos tienen que estar en la MISMA
carpeta. Necesita pythonnet, el AF SDK de OSIsoft, pandas y openpyxl.

Como se usa:

  1. Se busca la central (lista SAE*, "Buscar en toda la base" o la
     pestaña Navegar, igual que en Buscar_Rutas_PI).
  2. A la derecha salen sus datos (los atributos que tienen serie de
     tiempo, con su valor actual y unidad). Se tildan los que se
     quieren y "Agregar al pedido". Se puede repetir con otras
     centrales.
  3. Abajo: desde, hasta, intervalo (15 minutos u horario) y que valor
     por intervalo:
       - Promedio: promedio ponderado en el tiempo de cada intervalo
         (lo que corresponde para potencia o generacion). La fecha es
         el INICIO del intervalo: 00:00 = de 00:00 a 00:15 (o a 01:00).
       - Instantaneo: el valor interpolado justo en cada fecha (lo que
         hace el SoC del programa).
  4. "Extraer a Excel..." deja una hoja 'Datos' (fecha, fecha UTC y una
     columna por central y dato) y una hoja 'Pedido' con de donde salio
     cada columna.

Las fechas son la hora local del PC. Los dias de cambio de hora salen
como en PI: la hora que se repite en abril viene dos veces (por eso
esta la columna fecha UTC, que no se repite).
"""

import sys
import threading
from datetime import datetime, timedelta
from pathlib import Path

AQUI = Path(__file__).resolve().parent
if str(AQUI) not in sys.path:
    sys.path.insert(0, str(AQUI))

import Buscar_Rutas_PI as base  # noqa: E402

FORMATO_FECHA = "%Y-%m-%d %H:%M"
INTERVALOS = {"15 minutos": "15m", "Horario": "1h"}
TIPOS = {
    "Promedio del intervalo": "promedio",
    "Valor instantaneo (interpolado)": "instantaneo",
}
# PI corta las consultas muy grandes: se pide de a tramos.
DIAS_POR_TRAMO = 31
PROFUNDIDAD_ATRIBUTOS = 3


# ------------------------------------------------------------
# Sin AF (se prueba en cualquier maquina)
# ------------------------------------------------------------

def leer_fecha(texto):
    """'2026-10-01 00:00' (o sin la hora) -> datetime. ValueError si no."""

    texto = str(texto or "").strip()
    for formato in (FORMATO_FECHA, "%Y-%m-%d %H:%M:%S", "%Y-%m-%d",
                    "%d-%m-%Y %H:%M", "%d-%m-%Y"):
        try:
            return datetime.strptime(texto, formato)
        except ValueError:
            continue
    raise ValueError(
        f"No entiendo la fecha '{texto}'. Usa AAAA-MM-DD HH:MM, "
        f"ej. 2026-10-01 00:00.")


def mes_anterior(hoy=None):
    """(primero del mes anterior 00:00, primero de este mes 00:00)."""

    hoy = hoy or datetime.now()
    fin = datetime(hoy.year, hoy.month, 1)
    inicio = (fin - timedelta(days=1)).replace(day=1)
    return inicio, fin


def partir_rango(desde, hasta, dias=DIAS_POR_TRAMO):
    """[(inicio, fin), ...] de a lo mas 'dias' dias, contiguos."""

    if hasta <= desde:
        raise ValueError("'Hasta' tiene que ser despues de 'Desde'.")
    tramos = []
    inicio = desde
    while inicio < hasta:
        fin = min(inicio + timedelta(days=dias), hasta)
        tramos.append((inicio, fin))
        inicio = fin
    return tramos


def nombre_columna(pedido):
    unidad = f" [{pedido['unidad']}]" if pedido.get("unidad") else ""
    return f"{pedido['central']} - {pedido['atributo']}{unidad}"


def armar_tabla(series):
    """
    series: [(nombre_columna, [(fecha_utc, fecha_local, valor), ...])].
    DataFrame fecha | fecha UTC | una columna por serie, ordenado por la
    fecha UTC (la que no se repite en el cambio de hora).
    """

    import pandas as pd

    locales = {}
    columnas = {}
    for nombre, valores in series:
        datos = {}
        for utc, local, valor in valores:
            datos[utc] = valor
            locales.setdefault(utc, local)
        # Dos pedidos con el mismo nombre no se pisan.
        clave, n = nombre, 2
        while clave in columnas:
            clave, n = f"{nombre} ({n})", n + 1
        columnas[clave] = datos

    utcs = sorted(locales)
    tabla = pd.DataFrame({
        "fecha": pd.to_datetime([locales[u] for u in utcs]),
        "fecha UTC": pd.to_datetime(utcs),
    })
    for nombre, datos in columnas.items():
        tabla[nombre] = [datos.get(u) for u in utcs]
    return tabla


def a_numero(valor):
    """El valor de PI como numero si se puede; si no, como texto."""

    if valor is None:
        return None
    try:
        return float(valor)
    except (TypeError, ValueError):
        return str(valor)


# ------------------------------------------------------------
# AF
# ------------------------------------------------------------

class AFDatos(base.AF):
    """base.AF mas los atributos con serie de tiempo y su lectura."""

    @staticmethod
    def atributos_con_datos(elem):
        """
        [{"atributo", "unidad", "valor"}] de los atributos (y
        sub-atributos, 'Padre|Hijo') que tienen referencia de datos.
        """

        salida = []

        def recorrer(atributos, prefijo, nivel):
            for atributo in atributos:
                nombre = f"{prefijo}{atributo.Name}"
                if atributo.DataReference is not None:
                    unidad = ""
                    try:
                        if atributo.DefaultUOM is not None:
                            unidad = str(atributo.DefaultUOM.Abbreviation)
                    except Exception:  # noqa: BLE001 - sin unidad
                        unidad = ""
                    try:
                        actual = atributo.GetValue()
                        valor = (f"{actual.Value}" if actual.IsGood
                                 else f"({actual.Value})")
                    except Exception:  # noqa: BLE001 - se muestra igual
                        valor = "(no se pudo leer)"
                    salida.append({"atributo": nombre, "unidad": unidad,
                                   "valor": valor})
                if nivel < PROFUNDIDAD_ATRIBUTOS and atributo.HasChildren:
                    recorrer(atributo.Attributes, f"{nombre}|", nivel + 1)

        recorrer(elem.Attributes, "", 1)
        salida.sort(key=lambda a: a["atributo"].upper())
        return salida

    def atributo(self, ruta_elemento, nombre_atributo):
        camino = self.por_ruta(ruta_elemento)
        if not camino:
            raise RuntimeError(f"No se encontro {ruta_elemento} en AF.")
        partes = nombre_atributo.split("|")
        atributo = camino[-1].Attributes[partes[0]]
        for parte in partes[1:]:
            if atributo is None:
                break
            atributo = atributo.Attributes[parte]
        if atributo is None:
            raise RuntimeError(
                f"{camino[-1].Name} no tiene el atributo {nombre_atributo}.")
        return atributo

    def leer(self, atributo, desde, hasta, intervalo, tipo):
        """[(fecha_utc, fecha_local, valor)] de un tramo."""

        from OSIsoft.AF.Time import AFTimeRange, AFTimeSpan

        rango = AFTimeRange(desde.strftime("%Y-%m-%d %H:%M:%S"),
                            hasta.strftime("%Y-%m-%d %H:%M:%S"))
        paso = AFTimeSpan.Parse(intervalo)

        if tipo == "promedio":
            from OSIsoft.AF.Data import (
                AFCalculationBasis, AFSummaryTypes, AFTimestampCalculation,
            )
            resumen = atributo.Data.Summaries(
                rango, paso, AFSummaryTypes.Average,
                AFCalculationBasis.TimeWeighted,
                AFTimestampCalculation.Auto)
            valores = resumen[AFSummaryTypes.Average]
        else:
            valores = atributo.Data.InterpolatedValues(
                rango, paso, None, None, False)

        salida = []
        for v in valores:
            utc = str(v.Timestamp.UtcTime.ToString("yyyy-MM-dd HH:mm:ss"))
            local = str(v.Timestamp.LocalTime.ToString("yyyy-MM-dd HH:mm:ss"))
            salida.append((utc, local, a_numero(v.Value) if v.IsGood else None))
        return salida


def extraer(af, pedidos, desde, hasta, intervalo, tipo, avisar=print):
    """[(nombre_columna, valores)] de todos los pedidos, de a tramos."""

    tramos = partir_rango(desde, hasta)
    series = []
    for numero, pedido in enumerate(pedidos, start=1):
        atributo = af.atributo(pedido["ruta"], pedido["atributo"])
        valores = []
        vistos = set()
        for i, (inicio, fin) in enumerate(tramos, start=1):
            avisar(f"{numero}/{len(pedidos)} {nombre_columna(pedido)}: "
                   f"tramo {i}/{len(tramos)}...")
            for utc, local, valor in af.leer(atributo, inicio, fin,
                                             intervalo, tipo):
                # Los tramos se tocan en la punta: no repetir.
                if utc in vistos:
                    continue
                # Un promedio que empieza en 'hasta' es del intervalo
                # siguiente: no va.
                if tipo == "promedio" and local >= hasta.strftime(
                        "%Y-%m-%d %H:%M:%S"):
                    continue
                vistos.add(utc)
                valores.append((utc, local, valor))
        series.append((nombre_columna(pedido), valores))
    return series


def escribir_excel(ruta, tabla, pedidos, desde, hasta, intervalo_texto,
                   tipo_texto):
    import pandas as pd

    hoja_pedido = pd.DataFrame([
        {"columna": nombre_columna(p), "central": p["central"],
         "atributo": p["atributo"], "unidad": p.get("unidad", ""),
         "ruta": p["ruta"]} for p in pedidos])
    hoja_pedido["desde"] = desde.strftime(FORMATO_FECHA)
    hoja_pedido["hasta"] = hasta.strftime(FORMATO_FECHA)
    hoja_pedido["intervalo"] = intervalo_texto
    hoja_pedido["valor"] = tipo_texto

    with pd.ExcelWriter(ruta, engine="openpyxl") as writer:
        tabla.to_excel(writer, sheet_name="Datos", index=False)
        hoja_pedido.to_excel(writer, sheet_name="Pedido", index=False)
        for nombre in ("Datos", "Pedido"):
            hoja = writer.book[nombre]
            hoja.freeze_panes = "A2"
            for columna in hoja.columns:
                letra = columna[0].column_letter
                largo = max(len(str(c.value or "")) for c in columna[:200])
                hoja.column_dimensions[letra].width = min(max(12, largo + 2),
                                                          60)
        for celda in writer.book["Datos"]["A"][1:] + writer.book["Datos"]["B"][1:]:
            celda.number_format = "yyyy-mm-dd hh:mm"


# ------------------------------------------------------------
# La ventana
# ------------------------------------------------------------

def main():
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    root = tk.Tk()
    root.title("Extraer datos de PI AF")
    root.geometry("1280x800")

    prefs = base.leer_preferencias()
    estado = {"af": None, "elementos": [], "nodos": {}, "ocupado": False,
              "elegido": None, "atributos": [], "pedidos": []}
    var_servidor = tk.StringVar(value=prefs.get("servidor", base.SERVIDOR_AF))
    var_base = tk.StringVar(value=prefs.get("base", base.BASE_AF))
    var_texto = tk.StringVar()
    var_info = tk.StringVar()
    inicio, fin = mes_anterior()
    var_desde = tk.StringVar(value=inicio.strftime(FORMATO_FECHA))
    var_hasta = tk.StringVar(value=fin.strftime(FORMATO_FECHA))
    var_intervalo = tk.StringVar(value="15 minutos")
    var_tipo = tk.StringVar(value="Promedio del intervalo")

    def avisar(mensaje):
        root.after(0, var_info.set, mensaje)

    # ---- arriba ------------------------------------------------------
    arriba = tk.LabelFrame(root, text="PI AF", padx=10, pady=6)
    arriba.pack(fill="x", padx=14, pady=(10, 4))
    tk.Label(arriba, text="Servidor:").pack(side="left")
    tk.Entry(arriba, textvariable=var_servidor, width=18).pack(side="left",
                                                              padx=4)
    tk.Label(arriba, text="Base:").pack(side="left", padx=(8, 0))
    tk.Entry(arriba, textvariable=var_base, width=8).pack(side="left", padx=4)
    boton_lista = tk.Button(arriba, text="Actualizar lista SAE*")
    boton_lista.pack(side="left", padx=8)

    medio = tk.LabelFrame(root, text="Central (nombre aproximado)",
                          padx=10, pady=6)
    medio.pack(fill="x", padx=14, pady=4)
    entrada = tk.Entry(medio, textvariable=var_texto,
                       font=("Consolas", 11, "bold"))
    entrada.pack(side="left", fill="x", expand=True)
    boton_base = tk.Button(medio, text="Buscar en toda la base")
    boton_base.pack(side="left", padx=8)
    tk.Label(medio, text="ej. arena, tocopilla, cardones", fg="#444444",
             font=("Segoe UI", 8)).pack(side="left", padx=4)

    # ---- centro: elegir central | sus datos -------------------------
    centro = tk.PanedWindow(root, orient="horizontal", sashwidth=5)
    centro.pack(fill="both", expand=True, padx=14, pady=4)

    pestanas = ttk.Notebook(centro)
    centro.add(pestanas, width=640)

    marco_tabla = tk.Frame(pestanas)
    columnas = (("nombre", "Elemento", 200), ("pcp", "Nombre PCP/PID", 180),
                ("parecido", "Parecido", 65), ("ruta", "Ruta", 380))
    tabla = ttk.Treeview(marco_tabla, columns=[c for c, _, _ in columnas],
                         show="headings", selectmode="browse")
    for col, titulo, ancho in columnas:
        tabla.heading(col, text=titulo)
        tabla.column(col, width=ancho,
                     anchor="center" if col == "parecido" else "w")
    sb = ttk.Scrollbar(marco_tabla, orient="vertical", command=tabla.yview)
    tabla.configure(yscrollcommand=sb.set)
    sb.pack(side="right", fill="y")
    tabla.pack(side="left", fill="both", expand=True)
    pestanas.add(marco_tabla, text="Buscar")

    marco_arbol = tk.Frame(pestanas)
    arbol = ttk.Treeview(marco_arbol, show="tree", selectmode="browse")
    arbol.tag_configure("sae", font=("Segoe UI", 9, "bold"))
    sb2 = ttk.Scrollbar(marco_arbol, orient="vertical", command=arbol.yview)
    arbol.configure(yscrollcommand=sb2.set)
    sb2.pack(side="right", fill="y")
    arbol.pack(side="left", fill="both", expand=True)
    pestanas.add(marco_arbol, text="Navegar")

    marco_attr = tk.LabelFrame(centro, text="Datos del elemento",
                               padx=4, pady=4)
    attrs = ttk.Treeview(marco_attr, columns=("valor", "unidad"),
                         show="tree headings", selectmode="none")
    attrs.heading("#0", text="Dato (clic para tildar)")
    attrs.heading("valor", text="Valor actual")
    attrs.heading("unidad", text="Unidad")
    attrs.column("#0", width=290)
    attrs.column("valor", width=150)
    attrs.column("unidad", width=60, anchor="center")
    sb3 = ttk.Scrollbar(marco_attr, orient="vertical", command=attrs.yview)
    attrs.configure(yscrollcommand=sb3.set)
    barra_attr = tk.Frame(marco_attr)
    barra_attr.pack(side="bottom", fill="x", pady=(4, 0))
    boton_agregar = tk.Button(barra_attr, text="Agregar al pedido",
                              bg="#1f5fa8", fg="white",
                              font=("Segoe UI", 9, "bold"), state="disabled")
    boton_agregar.pack(side="right")
    sb3.pack(side="right", fill="y")
    attrs.pack(side="left", fill="both", expand=True)
    centro.add(marco_attr, width=560)

    # ---- el pedido ----------------------------------------------------
    marco_pedido = tk.LabelFrame(root, text="Pedido", padx=6, pady=4)
    marco_pedido.pack(fill="x", padx=14, pady=4)
    pedido = ttk.Treeview(marco_pedido, columns=("central", "atributo",
                                                 "unidad", "ruta"),
                          show="headings", height=5, selectmode="extended")
    for col, titulo, ancho in (("central", "Central", 200),
                               ("atributo", "Dato", 260),
                               ("unidad", "Unidad", 60),
                               ("ruta", "Ruta", 600)):
        pedido.heading(col, text=titulo)
        pedido.column(col, width=ancho)
    pedido.pack(side="left", fill="x", expand=True)
    botones_pedido = tk.Frame(marco_pedido)
    botones_pedido.pack(side="left", padx=6)
    boton_quitar = tk.Button(botones_pedido, text="Quitar", width=10)
    boton_quitar.pack(pady=2)
    boton_vaciar = tk.Button(botones_pedido, text="Vaciar", width=10)
    boton_vaciar.pack(pady=2)

    # ---- rango, intervalo, extraer -------------------------------------
    abajo = tk.LabelFrame(root, text="Que sacar", padx=10, pady=6)
    abajo.pack(fill="x", padx=14, pady=4)
    tk.Label(abajo, text="Desde:").pack(side="left")
    tk.Entry(abajo, textvariable=var_desde, width=17,
             font=("Consolas", 10)).pack(side="left", padx=4)
    tk.Label(abajo, text="Hasta:").pack(side="left", padx=(8, 0))
    tk.Entry(abajo, textvariable=var_hasta, width=17,
             font=("Consolas", 10)).pack(side="left", padx=4)
    boton_mes = tk.Button(abajo, text="Mes anterior")
    boton_mes.pack(side="left", padx=4)
    tk.Label(abajo, text="Intervalo:").pack(side="left", padx=(12, 0))
    ttk.Combobox(abajo, textvariable=var_intervalo, values=list(INTERVALOS),
                 state="readonly", width=11).pack(side="left", padx=4)
    tk.Label(abajo, text="Valor:").pack(side="left", padx=(8, 0))
    ttk.Combobox(abajo, textvariable=var_tipo, values=list(TIPOS),
                 state="readonly", width=28).pack(side="left", padx=4)
    boton_extraer = tk.Button(abajo, text="Extraer a Excel...",
                              bg="#2d7a2d", fg="white",
                              font=("Segoe UI", 9, "bold"))
    boton_extraer.pack(side="right")
    tk.Label(root, textvariable=var_info, anchor="w", fg="#444444").pack(
        fill="x", padx=14, pady=(0, 8))

    botones_af = (boton_lista, boton_base, boton_extraer)

    # ---- AF en segundo plano ------------------------------------------
    def con_af(trabajo, despues, texto):
        estado["ocupado"] = True
        for boton in botones_af:
            boton.config(state="disabled")
        var_info.set(texto)
        root.config(cursor="watch")

        def hilo():
            try:
                if estado["af"] is None:
                    avisar(f"Conectando a {var_servidor.get()}...")
                    estado["af"] = AFDatos(var_servidor.get().strip(),
                                           var_base.get().strip())
                    base.guardar_preferencias(
                        servidor=var_servidor.get().strip(),
                        base=var_base.get().strip())
                resultado = trabajo(estado["af"])
                root.after(0, fin_af, despues, resultado, None)
            except Exception as error:  # noqa: BLE001 - se muestra
                root.after(0, fin_af, despues, None, error)

        threading.Thread(target=hilo, daemon=True).start()

    def fin_af(despues, resultado, error):
        estado["ocupado"] = False
        root.config(cursor="")
        for boton in botones_af:
            boton.config(state="normal")
        if error is not None:
            estado["elegido"] = None  # se puede volver a intentar
            var_info.set("No se pudo.")
            base.mostrar_error("PI AF", str(error))
            return
        despues(resultado)

    # ---- buscar la central ------------------------------------------
    def llenar_tabla(resultados):
        tabla.delete(*tabla.get_children())
        for elem, nota in resultados:
            tabla.insert("", "end", values=(
                elem["nombre"], elem["pcp"],
                "" if nota is None else f"{nota} %", elem["ruta"]))

    def mostrar(*_):
        elementos = estado["elementos"]
        if not elementos or not base.limpiar(var_texto.get()):
            return
        resultados = base.buscar(elementos, var_texto.get())
        llenar_tabla(resultados)
        var_info.set(
            f"{len(resultados)} parecido(s). Si no esta (ej. una central "
            f"que no es SAE), 'Buscar en toda la base' o Navegar."
            if resultados else
            "Ninguno se parece: prueba 'Buscar en toda la base' o Navegar.")

    def actualizar_lista():
        def guardar(elementos):
            estado["elementos"] = elementos
            try:
                import json
                from datetime import date
                base.ruta_lista().write_text(json.dumps(
                    {"servidor": var_servidor.get(), "base": var_base.get(),
                     "fecha": date.today().isoformat(),
                     "elementos": elementos},
                    ensure_ascii=False, indent=1), encoding="utf-8")
            except OSError:
                pass
            mostrar()
            var_info.set(f"{len(elementos):,} elementos SAE* leidos de AF.")

        con_af(lambda af: af.lista_sae(avisar=avisar), guardar,
               "Leyendo los elementos SAE*...")

    def en_toda_la_base():
        texto = var_texto.get().strip()
        if not texto:
            var_info.set("Escribe parte del nombre del elemento.")
            return

        def mostrar_base(elementos):
            notas = {id(e): n for e, n in base.buscar(elementos, texto,
                                                      limite=10 ** 6)}
            elementos.sort(key=lambda e: -notas.get(id(e), 0))
            llenar_tabla([(e, notas.get(id(e))) for e in elementos])
            pestanas.select(0)
            var_info.set(f"{len(elementos)} elemento(s) con '{texto}' en el "
                         f"nombre.")

        con_af(lambda af: af.en_toda_la_base(texto, avisar=avisar),
               mostrar_base, f"Buscando '*{texto}*' en toda la base...")

    # ---- los datos del elemento ---------------------------------------
    def elegir(ruta, elem_af=None):
        if not ruta or estado["ocupado"] or estado["elegido"] == ruta:
            return
        estado["elegido"] = ruta
        attrs.delete(*attrs.get_children())
        boton_agregar.config(state="disabled")
        nombre = ruta.split("\\")[-1]
        marco_attr.config(text=f"Datos de {nombre}")

        def leer(af):
            elem = elem_af
            if elem is None:
                camino = af.por_ruta(ruta)
                if not camino:
                    raise RuntimeError(f"No se encontro {ruta} en AF.")
                elem = camino[-1]
            return af.atributos_con_datos(elem)

        def pintar(lista):
            if estado["elegido"] != ruta:
                return
            estado["atributos"] = lista
            for i, dato in enumerate(lista):
                attrs.insert("", "end", iid=str(i),
                             text="☐ " + dato["atributo"],
                             values=(dato["valor"], dato["unidad"]))
            boton_agregar.config(state="normal" if lista else "disabled")
            var_info.set(
                f"{nombre}: {len(lista)} dato(s) con serie de tiempo. Tilda "
                f"los que quieras y 'Agregar al pedido'." if lista else
                f"{nombre} no tiene datos con serie de tiempo. Prueba con "
                f"el elemento de arriba o de abajo en Navegar.")

        con_af(leer, pintar, f"Leyendo los datos de {nombre}...")

    def tildar(evento):
        iid = attrs.identify_row(evento.y)
        if not iid:
            return
        texto = attrs.item(iid, "text")
        nuevo = ("☑ " if texto.startswith("☐") else "☐ ") + texto[2:]
        attrs.item(iid, text=nuevo)

    def agregar():
        ruta = estado["elegido"]
        tildados = [estado["atributos"][int(i)] for i in attrs.get_children()
                    if attrs.item(i, "text").startswith("☑")]
        if not tildados:
            var_info.set("Tilda al menos un dato (clic sobre el nombre).")
            return
        central = ruta.split("\\")[-1]
        ya = {(p["ruta"], p["atributo"]) for p in estado["pedidos"]}
        nuevos = 0
        for dato in tildados:
            if (ruta, dato["atributo"]) in ya:
                continue
            item = {"central": central, "ruta": ruta,
                    "atributo": dato["atributo"], "unidad": dato["unidad"]}
            estado["pedidos"].append(item)
            pedido.insert("", "end", values=(central, dato["atributo"],
                                             dato["unidad"], ruta))
            nuevos += 1
        for i in attrs.get_children():
            attrs.item(i, text="☐ " + attrs.item(i, "text")[2:])
        var_info.set(f"{nuevos} dato(s) agregado(s). Pedido: "
                     f"{len(estado['pedidos'])}.")

    def quitar():
        seleccion = pedido.selection()
        indices = sorted((pedido.index(i) for i in seleccion), reverse=True)
        for i in indices:
            del estado["pedidos"][i]
        pedido.delete(*seleccion)

    def vaciar():
        estado["pedidos"].clear()
        pedido.delete(*pedido.get_children())

    # ---- navegar --------------------------------------------------------
    def insertar_nodo(padre, elem):
        nombre = str(elem.Name)
        iid = arbol.insert(padre, "end", text=nombre,
                           tags=("sae",) if base.limpiar(nombre).startswith(
                               "SAE") else ())
        estado["nodos"][iid] = {"elem": elem, "ruta": str(elem.GetPath()),
                                "cargado": False}
        if elem.HasChildren:
            arbol.insert(iid, "end", text="...")

    def al_abrir(*_):
        iid = arbol.focus()
        nodo = estado["nodos"].get(iid)
        if nodo is None or nodo["cargado"]:
            return
        nodo["cargado"] = True
        arbol.delete(*arbol.get_children(iid))
        for hijo in sorted(base.AF.hijos(nodo["elem"]),
                           key=lambda e: str(e.Name)):
            insertar_nodo(iid, hijo)

    def al_cambiar_pestana(*_):
        if pestanas.index("current") == 1 and not arbol.get_children(""):
            con_af(lambda af: af.raices(),
                   lambda elems: [insertar_nodo("", e) for e in
                                  sorted(elems, key=lambda e: str(e.Name))],
                   "Leyendo la raiz de la base...")

    def al_elegir_en_arbol(*_):
        sel = arbol.selection()
        if sel and sel[0] in estado["nodos"]:
            nodo = estado["nodos"][sel[0]]
            elegir(nodo["ruta"], nodo["elem"])

    def al_elegir_en_tabla(*_):
        sel = tabla.selection()
        if sel:
            elegir(tabla.item(sel[0], "values")[3])

    # ---- extraer ----------------------------------------------------------
    def poner_mes_anterior():
        desde, hasta = mes_anterior()
        var_desde.set(desde.strftime(FORMATO_FECHA))
        var_hasta.set(hasta.strftime(FORMATO_FECHA))

    def extraer_excel():
        if not estado["pedidos"]:
            messagebox.showwarning("Pedido vacio",
                                   "Agrega al menos un dato al pedido.")
            return
        try:
            desde = leer_fecha(var_desde.get())
            hasta = leer_fecha(var_hasta.get())
            partir_rango(desde, hasta)
        except ValueError as error:
            messagebox.showwarning("Fechas", str(error))
            return
        intervalo_texto, tipo_texto = var_intervalo.get(), var_tipo.get()
        centrales = sorted({p["central"] for p in estado["pedidos"]})
        sugerido = (f"PI_{centrales[0] if len(centrales) == 1 else 'centrales'}"
                    f"_{desde:%Y%m%d}_{hasta:%Y%m%d}_"
                    f"{INTERVALOS[intervalo_texto]}.xlsx").replace(" ", "_")
        destino = filedialog.asksaveasfilename(
            title="Guardar como", defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")], initialfile=sugerido,
            initialdir=prefs.get("carpeta_salida", ""))
        if not destino:
            return
        base.guardar_preferencias(carpeta_salida=str(Path(destino).parent))
        pedidos = list(estado["pedidos"])

        def trabajo(af):
            series = extraer(af, pedidos, desde, hasta,
                             INTERVALOS[intervalo_texto], TIPOS[tipo_texto],
                             avisar=avisar)
            tabla_datos = armar_tabla(series)
            avisar(f"Escribiendo {destino}...")
            try:
                escribir_excel(destino, tabla_datos, pedidos, desde, hasta,
                               intervalo_texto, tipo_texto)
            except PermissionError as error:
                raise RuntimeError(f"No se pudo escribir {destino}: cerralo "
                                   f"en Excel y volve a intentar.") from error
            return len(tabla_datos)

        def listo(filas):
            var_info.set(f"Listo: {filas:,} fila(s) en {destino}")
            if messagebox.askyesno("Listo", f"Se guardo:\n{destino}\n\n"
                                   f"¿Abrirlo?"):
                import os
                try:
                    os.startfile(destino)  # solo Windows
                except (AttributeError, OSError):
                    pass

        con_af(trabajo, listo, "Extrayendo...")

    # ---- arranque -------------------------------------------------------
    try:
        import json
        guardada = json.loads(base.ruta_lista().read_text(encoding="utf-8"))
        estado["elementos"] = guardada["elementos"]
        var_info.set(f"{len(estado['elementos']):,} elementos SAE* (lista del "
                     f"{guardada.get('fecha', '?')}). Escribe una central.")
    except (OSError, ValueError, KeyError):
        var_info.set("Aprieta 'Actualizar lista SAE*' (solo la primera vez), "
                     "busca en toda la base o usa Navegar.")

    boton_lista.config(command=actualizar_lista)
    boton_base.config(command=en_toda_la_base)
    boton_agregar.config(command=agregar)
    boton_quitar.config(command=quitar)
    boton_vaciar.config(command=vaciar)
    boton_mes.config(command=poner_mes_anterior)
    boton_extraer.config(command=extraer_excel)
    tabla.bind("<<TreeviewSelect>>", al_elegir_en_tabla)
    arbol.bind("<<TreeviewOpen>>", al_abrir)
    arbol.bind("<<TreeviewSelect>>", al_elegir_en_arbol)
    attrs.bind("<Button-1>", tildar)
    pedido.bind("<Delete>", lambda e: quitar())
    pestanas.bind("<<NotebookTabChanged>>", al_cambiar_pestana)
    entrada.bind("<Return>", lambda e: en_toda_la_base()
                 if not tabla.get_children() else None)
    var_texto.trace_add("write", mostrar)
    entrada.focus_set()
    root.mainloop()


if __name__ == "__main__":
    main()
