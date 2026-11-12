# -*- coding: utf-8 -*-
r"""
Buscar en PI AF la ruta de una central (la que va en la columna SoC_PI
del Diccionario de Centrales.xlsx).

Funciona guardado en cualquier carpeta: no usa nada del resto del
programa. Necesita Python con pythonnet (pip install pythonnet), el AF
SDK de OSIsoft instalado (el PC donde se usa PI System Explorer o
DataLink) y, para la lista de centrales, openpyxl.

Tres formas de llegar a la ruta:

  1. Buscar: se baja UNA vez la lista de elementos "SAE*" de la base
     (nombre, atributo "Nombre PCP/PID", si tiene "Porc Carga
     Disponible" y su ruta) y se guarda en __config__. Despues se
     escribe un nombre aproximado ('manzano', 'arica 2', 'lile') y la
     lista se filtra mientras se escribe, sin conectarse.
     "Buscar en toda la base" busca ademas por nombre de elemento en
     toda la base (para lo que no empieza con SAE).
  2. Centrales.xlsx: se elige el archivo y aparecen sus centrales,
     marcadas con o sin ruta en SoC_PI. Clic en una muestra su ruta (si
     ya tiene) o la busca y deja elegida la mas parecida (posible
     ruta). "Guardar en Centrales.xlsx" escribe la ruta elegida en la
     columna SoC_PI de esa central (pregunta antes; si la columna no
     existe, la agrega).
  3. Navegar: el arbol de la base, carpeta por carpeta, como en PI
     System Explorer. "Ver en el arbol" abre el arbol en el elemento
     elegido en la busqueda, para ver los que estan a su lado.

Al elegir un elemento se ven sus atributos (con "Nombre PCP/PID" y
"Porc Carga Disponible" arriba). "Copiar ruta" deja en el portapapeles
la ruta lista para pegar en SoC_PI:

    \\SRV-SCADA-AF2\SEN\Generación\SEN\03 Región II\SAE-ARENA
"""

import difflib
import json
import re
import sys
import threading
import unicodedata
from datetime import date
from pathlib import Path

AQUI = Path(__file__).resolve().parent

SERVIDOR_AF = "SRV-SCADA-AF2"
BASE_AF = "SEN"
ATRIBUTO_SOC = "Porc Carga Disponible"
ATRIBUTO_NOMBRE = "Nombre PCP/PID"
FILTRO_SAE = "Name:SAE*"
RUTA_AF_SDK = r"C:\Program Files (x86)\PIPC\AF\PublicAssemblies\4.0"
MAXIMO_RESULTADOS = 80
MAXIMO_BUSQUEDA_BASE = 300
MAXIMO_ATRIBUTOS = 80

# Encabezados del Diccionario (normalizados con limpiar()).
TITULO_CENTRAL = "BALANCE BESS"
TITULOS_RUTA = ("SOC PI", "RUTA SOC")


# ------------------------------------------------------------
# Donde guardar / leer (sirve desde cualquier carpeta)
# ------------------------------------------------------------

def carpeta_config():
    """El __config__ mas cercano hacia arriba, o esta misma carpeta."""

    for carpeta in [AQUI, *AQUI.parents]:
        if (carpeta / "__config__").is_dir():
            return carpeta / "__config__"
    return AQUI


def ruta_lista():
    return carpeta_config() / "elementos_sae_pi.json"


def ruta_preferencias():
    return carpeta_config() / "buscar_rutas_pi.json"


def leer_preferencias():
    try:
        return json.loads(ruta_preferencias().read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}


def guardar_preferencias(**cambios):
    datos = leer_preferencias()
    datos.update(cambios)
    try:
        ruta_preferencias().write_text(
            json.dumps(datos, ensure_ascii=False, indent=1), encoding="utf-8")
    except OSError:
        pass


# ------------------------------------------------------------
# La busqueda (sin AF: se prueba en cualquier maquina)
# ------------------------------------------------------------

def limpiar(texto):
    texto = unicodedata.normalize("NFKD", str(texto or ""))
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    return " ".join(re.sub(r"[^0-9A-Za-z]+", " ", texto.upper()).split())


# Palabras que casi todos los nombres tienen: no ayudan a distinguir y,
# si no estan en el elemento, no deberian bajar el parecido.
PALABRAS_COMUNES = {"SAE", "CRCA", "BESS", "PFV", "PE", "BAT", "FV"}

ROMANOS = {"I": "1", "II": "2", "III": "3", "IV": "4", "V": "5"}


def _palabras(texto):
    palabras = []
    for palabra in limpiar(texto).split():
        palabras.append(ROMANOS.get(palabra, palabra))
    return palabras


def buscar(elementos, texto, limite=MAXIMO_RESULTADOS):
    """
    [(elemento, parecido 0..100)]: todas las palabras (en cualquier
    orden) arriba; despues lo que se parece aunque este mal escrito.
    Se busca en el nombre, el Nombre PCP/PID y la ruta.
    """

    palabras = _palabras(texto)
    utiles = [p for p in palabras if p not in PALABRAS_COMUNES] or palabras
    if not utiles:
        return []

    salida = []
    for elem in elementos:
        suyas = _palabras(
            f"{elem.get('nombre', '')} {elem.get('pcp', '')} "
            f"{elem.get('ruta', '').split(chr(92))[-1]}"
        )
        if not suyas:
            continue
        notas = []
        for palabra in utiles:
            if palabra in suyas:
                notas.append(1.0)
            elif any(palabra in s or (len(s) > 2 and s in palabra)
                     for s in suyas):
                notas.append(0.9)
            else:
                notas.append(max(
                    difflib.SequenceMatcher(None, palabra, s).ratio()
                    for s in suyas))
        nota = sum(notas) / len(notas)
        # Coincide justo el Nombre PCP/PID: va primero.
        if limpiar(texto) and limpiar(texto) == limpiar(elem.get("pcp")):
            nota = 1.01
        if nota >= 0.6:
            salida.append((elem, min(100, round(100 * nota))))

    salida.sort(key=lambda x: (-x[1], x[0].get("ruta", "")))
    return salida[:limite]


def partes_de_ruta(ruta):
    r"""[elemento, ...] de \\SERVIDOR\BASE\A\B (sin el atributo |...)."""

    ruta = str(ruta or "").split("|")[0].strip()
    partes = [p for p in ruta.lstrip("\\").split("\\") if p.strip()]
    return partes[2:]


# ------------------------------------------------------------
# Centrales.xlsx: que centrales tienen ruta en SoC_PI
# ------------------------------------------------------------

def centrales_del_diccionario(ruta_centrales):
    """
    [(central, ruta_o_vacio)] de la hoja Diccionario (formato de una
    sola tabla, con encabezado Balance_BESS). Lista vacia si la hoja
    no tiene ese encabezado.
    """

    from openpyxl import load_workbook

    libro = load_workbook(ruta_centrales, read_only=True, data_only=True)
    try:
        hoja = next(
            (libro[n] for n in libro.sheetnames
             if limpiar(n) == "DICCIONARIO"), None)
        if hoja is None:
            return []

        filas = list(hoja.iter_rows(values_only=True))
    finally:
        libro.close()

    for indice, fila in enumerate(filas[:10]):
        titulos = [limpiar(v) for v in fila]
        if TITULO_CENTRAL not in titulos:
            continue
        col_central = titulos.index(TITULO_CENTRAL)
        col_ruta = next(
            (titulos.index(t) for t in TITULOS_RUTA if t in titulos), None)
        salida = []
        for resto in filas[indice + 1:]:
            if col_central >= len(resto) or not limpiar(resto[col_central]):
                continue
            ruta = ""
            if col_ruta is not None and col_ruta < len(resto):
                ruta = str(resto[col_ruta] or "").strip()
            salida.append((str(resto[col_central]).strip(), ruta))
        return salida

    return []


class CentralesAbierto(Exception):
    """Centrales.xlsx no se puede escribir (lo tiene abierto Excel)."""


def guardar_ruta_en_centrales(ruta_centrales, central, ruta_pi):
    """
    Escribe ruta_pi en la columna SoC_PI de la fila de 'central' en la
    hoja Diccionario. Si la columna no existe, la agrega al final de los
    encabezados (con el formato del encabezado de al lado). Devuelve la
    ruta que habia antes ("" si no habia).

    ValueError si no se encuentra la hoja, el encabezado o la central;
    CentralesAbierto si el archivo esta abierto en Excel.
    """

    from copy import copy

    from openpyxl import load_workbook

    libro = load_workbook(ruta_centrales)
    hoja = next((libro[n] for n in libro.sheetnames
                 if limpiar(n) == "DICCIONARIO"), None)
    if hoja is None:
        raise ValueError("Centrales.xlsx no tiene la hoja Diccionario.")

    fila_titulos = None
    for fila in range(1, min(hoja.max_row, 10) + 1):
        titulos = {limpiar(hoja.cell(fila, c).value): c
                   for c in range(1, hoja.max_column + 1)}
        if TITULO_CENTRAL in titulos:
            fila_titulos = fila
            break
    if fila_titulos is None:
        raise ValueError(
            "El Diccionario no tiene la fila de encabezados con Balance_BESS.")

    col_central = titulos[TITULO_CENTRAL]
    col_ruta = next((titulos[t] for t in TITULOS_RUTA if t in titulos), None)

    if col_ruta is None:
        ultima = max(c for t, c in titulos.items() if t)
        col_ruta = ultima + 1
        titulo = hoja.cell(fila_titulos, col_ruta)
        titulo.value = "SoC_PI"
        titulo._style = copy(hoja.cell(fila_titulos, ultima)._style)
        letra = titulo.column_letter
        hoja.column_dimensions[letra].width = 90

    fila_central = next(
        (f for f in range(fila_titulos + 1, hoja.max_row + 1)
         if limpiar(hoja.cell(f, col_central).value) == limpiar(central)),
        None)
    if fila_central is None:
        raise ValueError(f"La central {central} no esta en el Diccionario.")

    celda = hoja.cell(fila_central, col_ruta)
    anterior = str(celda.value or "").strip()
    if not celda.has_style:
        celda._style = copy(hoja.cell(fila_central, col_central)._style)
    celda.value = str(ruta_pi).split("|")[0].strip()

    try:
        libro.save(ruta_centrales)
    except PermissionError as error:
        raise CentralesAbierto(
            f"No se pudo guardar {ruta_centrales}: cerralo en Excel y "
            f"volve a apretar el boton.") from error
    return anterior


# ------------------------------------------------------------
# AF (solo en el PC con el AF SDK)
# ------------------------------------------------------------

def cargar_af_sdk():
    try:
        import clr  # pythonnet
    except ImportError as error:
        raise RuntimeError(
            "Falta pythonnet. Instalalo con:\n\npip install pythonnet"
        ) from error

    if RUTA_AF_SDK not in sys.path:
        sys.path.append(RUTA_AF_SDK)
    try:
        clr.AddReference("OSIsoft.AFSDK")
        from OSIsoft.AF import PISystems
        from OSIsoft.AF.Search import AFElementSearch
    except Exception as error:  # el error real es de .NET
        raise RuntimeError(
            f"No se pudo cargar el AF SDK de OSIsoft (se busca en "
            f"{RUTA_AF_SDK}). Esta herramienta solo funciona en un PC con "
            f"PI AF instalado.\n\nDetalle: {error}"
        ) from error
    return PISystems, AFElementSearch


class AF:
    """La base de AF, con lo que la ventana necesita."""

    def __init__(self, servidor, base):
        PISystems, self._busqueda = cargar_af_sdk()
        sistema = PISystems()[servidor]
        if sistema is None:
            raise RuntimeError(
                f"El servidor {servidor} no esta registrado en este PC "
                f"(PI System Explorer -> Servidores).")
        sistema.Connect()
        self.base = sistema.Databases[base]
        if self.base is None:
            raise RuntimeError(f"{servidor} no tiene la base '{base}'.")

    @staticmethod
    def a_dict(elem):
        nombre_pcp = ""
        atributo = elem.Attributes[ATRIBUTO_NOMBRE]
        if atributo is not None:
            try:
                nombre_pcp = str(atributo.GetValue().Value or "")
            except Exception:  # noqa: BLE001 - valor ilegible
                nombre_pcp = ""
        return {
            "nombre": str(elem.Name),
            "pcp": nombre_pcp,
            "soc": elem.Attributes[ATRIBUTO_SOC] is not None,
            "ruta": str(elem.GetPath()),
        }

    def buscar(self, filtro, maximo, avisar=print):
        busqueda = self._busqueda(self.base, "BuscarRutasPI", filtro)
        salida = []
        for elem in busqueda.FindObjects(0, True, 1000):
            salida.append(self.a_dict(elem))
            if len(salida) % 25 == 0:
                avisar(f"Leyendo elementos... {len(salida)}")
            if len(salida) >= maximo:
                break
        return salida

    def lista_sae(self, avisar=print):
        return self.buscar(FILTRO_SAE, 100000, avisar=avisar)

    def en_toda_la_base(self, texto, avisar=print):
        texto = str(texto).strip().replace("'", "")
        return self.buscar(f"Name:'*{texto}*'", MAXIMO_BUSQUEDA_BASE,
                           avisar=avisar)

    def raices(self):
        return list(self.base.Elements)

    @staticmethod
    def hijos(elem):
        return list(elem.Elements)

    def por_ruta(self, ruta):
        """[elemento raiz, ..., elemento] de la ruta (None si se corta)."""

        camino = []
        coleccion = self.base.Elements
        for nombre in partes_de_ruta(ruta):
            elem = coleccion[nombre]
            if elem is None:
                return None
            camino.append(elem)
            coleccion = elem.Elements
        return camino

    @staticmethod
    def atributos(elem):
        """[(nombre, valor)] con Nombre PCP/PID y el SoC primero."""

        filas = []
        for atributo in list(elem.Attributes)[:MAXIMO_ATRIBUTOS]:
            try:
                valor = atributo.GetValue()
                texto = f"{valor.Value}"
                if atributo.Name == ATRIBUTO_SOC:
                    texto += f"   ({valor.Timestamp.LocalTime})"
            except Exception:  # noqa: BLE001 - se muestra igual
                texto = "(no se pudo leer)"
            filas.append((str(atributo.Name), texto))
        primero = {ATRIBUTO_NOMBRE: 0, ATRIBUTO_SOC: 1}
        filas.sort(key=lambda f: primero.get(f[0], 2))
        return filas


# ------------------------------------------------------------
# Cuadro de error con boton "Copiar" (copia de Script/dialogos.py: esta
# herramienta no importa nada del programa).
# ------------------------------------------------------------

def mostrar_error(titulo, mensaje, parent=None):
    import tkinter as tk

    raiz = parent or tk._default_root
    ventana = tk.Toplevel(raiz)
    ventana.title(titulo)
    if raiz is not None:
        ventana.transient(raiz)
    cuerpo = tk.Frame(ventana, padx=14, pady=12)
    cuerpo.pack(fill="both", expand=True)
    tk.Label(cuerpo, text="✖", fg="#c0392b",
             font=("Segoe UI", 20, "bold")).pack(side="left", anchor="n",
                                                padx=(0, 12))
    caja = tk.Text(cuerpo, wrap="word", width=78,
                   height=max(4, min(20, str(mensaje).count("\n") + 3)),
                   font=("Segoe UI", 9), relief="flat",
                   background=ventana.cget("background"))
    caja.insert("1.0", str(mensaje))
    caja.config(state="disabled")
    caja.pack(side="left", fill="both", expand=True)
    botones = tk.Frame(ventana, padx=14, pady=10)
    botones.pack(fill="x")

    def copiar():
        ventana.clipboard_clear()
        ventana.clipboard_append(f"{titulo}\n\n{mensaje}")

    tk.Button(botones, text="Aceptar", width=12,
              command=ventana.destroy).pack(side="right")
    tk.Button(botones, text="Copiar", width=12, bg="#1f5fa8", fg="white",
              font=("Segoe UI", 9, "bold"),
              command=copiar).pack(side="right", padx=8)
    ventana.bind("<Escape>", lambda e: ventana.destroy())
    ventana.grab_set()
    ventana.wait_window()


# ------------------------------------------------------------
# La ventana
# ------------------------------------------------------------

def main():
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    root = tk.Tk()
    root.title("Buscar rutas de PI AF (SoC)")
    root.geometry("1250x720")

    prefs = leer_preferencias()
    estado = {"af": None, "elementos": [], "nodos": {}, "elegido": None,
              "ocupado": False, "centrales": None, "central": None}
    var_servidor = tk.StringVar(value=prefs.get("servidor", SERVIDOR_AF))
    var_base = tk.StringVar(value=prefs.get("base", BASE_AF))
    var_texto = tk.StringVar()
    var_info = tk.StringVar()
    var_ruta = tk.StringVar()

    def avisar(mensaje):
        root.after(0, var_info.set, mensaje)

    # ---- arriba: conexion ----------------------------------------
    arriba = tk.LabelFrame(root, text="PI AF", padx=10, pady=6)
    arriba.pack(fill="x", padx=14, pady=(10, 4))
    tk.Label(arriba, text="Servidor:").pack(side="left")
    tk.Entry(arriba, textvariable=var_servidor, width=18).pack(side="left",
                                                              padx=4)
    tk.Label(arriba, text="Base:").pack(side="left", padx=(8, 0))
    tk.Entry(arriba, textvariable=var_base, width=8).pack(side="left", padx=4)
    boton_lista = tk.Button(arriba, text="Actualizar lista SAE*")
    boton_lista.pack(side="left", padx=8)
    boton_centrales = tk.Button(arriba, text="Centrales.xlsx...")
    boton_centrales.pack(side="left", padx=8)

    # ---- busqueda --------------------------------------------------
    medio = tk.LabelFrame(root, text="Nombre aproximado", padx=10, pady=6)
    medio.pack(fill="x", padx=14, pady=4)
    entrada = tk.Entry(medio, textvariable=var_texto,
                       font=("Consolas", 11, "bold"))
    entrada.pack(side="left", fill="x", expand=True)
    boton_base = tk.Button(medio, text="Buscar en toda la base")
    boton_base.pack(side="left", padx=8)
    tk.Label(medio, text="ej. manzano, arica 2, lile", fg="#444444",
             font=("Segoe UI", 8)).pack(side="left", padx=4)

    # ---- centro: centrales | pestañas | atributos -----------------
    centro = tk.PanedWindow(root, orient="horizontal", sashwidth=5)
    centro.pack(fill="both", expand=True, padx=14, pady=4)

    marco_centrales = tk.LabelFrame(centro, text="Centrales (SoC_PI)",
                                    padx=4, pady=4)
    lista_centrales = ttk.Treeview(marco_centrales, columns=("c",),
                                   displaycolumns=(), show="tree",
                                   selectmode="browse")
    lista_centrales.column("#0", width=230)
    lista_centrales.tag_configure("sin", foreground="#c0392b")
    lista_centrales.tag_configure("con", foreground="#2d7a2d")
    lista_centrales.pack(fill="both", expand=True)
    centro.add(marco_centrales, width=250)

    pestanas = ttk.Notebook(centro)
    centro.add(pestanas, width=650)

    # Buscar
    marco_tabla = tk.Frame(pestanas)
    columnas = (("nombre", "Elemento", 190), ("pcp", "Nombre PCP/PID", 200),
                ("soc", "SoC", 40), ("parecido", "Parecido", 65),
                ("ruta", "Ruta", 420))
    tabla = ttk.Treeview(marco_tabla, columns=[c for c, _, _ in columnas],
                         show="headings", selectmode="browse")
    for col, titulo, ancho in columnas:
        tabla.heading(col, text=titulo)
        tabla.column(col, width=ancho,
                     anchor="center" if col in ("soc", "parecido") else "w")
    sb = ttk.Scrollbar(marco_tabla, orient="vertical", command=tabla.yview)
    tabla.configure(yscrollcommand=sb.set)
    sb.pack(side="right", fill="y")
    tabla.pack(side="left", fill="both", expand=True)
    pestanas.add(marco_tabla, text="Buscar")

    # Navegar
    marco_arbol = tk.Frame(pestanas)
    arbol = ttk.Treeview(marco_arbol, show="tree", selectmode="browse")
    arbol.tag_configure("sae", font=("Segoe UI", 9, "bold"))
    sb2 = ttk.Scrollbar(marco_arbol, orient="vertical", command=arbol.yview)
    arbol.configure(yscrollcommand=sb2.set)
    sb2.pack(side="right", fill="y")
    arbol.pack(side="left", fill="both", expand=True)
    pestanas.add(marco_arbol, text="Navegar")

    marco_attr = tk.LabelFrame(centro, text="Atributos del elemento",
                               padx=4, pady=4)
    attrs = ttk.Treeview(marco_attr, columns=("valor",), show="tree headings")
    attrs.heading("#0", text="Atributo")
    attrs.heading("valor", text="Valor")
    attrs.column("#0", width=170)
    attrs.column("valor", width=170)
    attrs.tag_configure("clave", font=("Segoe UI", 9, "bold"))
    attrs.pack(fill="both", expand=True)
    centro.add(marco_attr, width=330)

    # ---- abajo: la ruta ---------------------------------------------
    abajo = tk.Frame(root)
    abajo.pack(fill="x", padx=14, pady=(4, 2))
    tk.Label(abajo, text="Ruta:").pack(side="left")
    tk.Entry(abajo, textvariable=var_ruta, state="readonly",
             font=("Consolas", 9)).pack(side="left", fill="x", expand=True,
                                        padx=6)
    boton_copiar = tk.Button(abajo, text="Copiar ruta", bg="#1f5fa8",
                             fg="white", font=("Segoe UI", 9, "bold"))
    boton_copiar.pack(side="left")
    boton_ir = tk.Button(abajo, text="Ver en el arbol")
    boton_ir.pack(side="left", padx=6)
    boton_guardar = tk.Button(abajo, text="Guardar en Centrales.xlsx",
                              bg="#2d7a2d", fg="white",
                              font=("Segoe UI", 9, "bold"), state="disabled")
    boton_guardar.pack(side="left")
    var_central = tk.StringVar(value="(elegi una central a la izquierda)")
    tk.Label(abajo, textvariable=var_central, fg="#2d7a2d",
             font=("Segoe UI", 8, "bold")).pack(side="left", padx=6)
    tk.Label(root, textvariable=var_info, anchor="w", fg="#444444").pack(
        fill="x", padx=14, pady=(0, 8))

    botones_af = (boton_lista, boton_base, boton_ir)

    # ---- AF en segundo plano --------------------------------------
    def con_af(trabajo, despues, texto):
        """Corre trabajo(af) en un hilo; conecta la primera vez."""

        estado["ocupado"] = True
        for boton in botones_af:
            boton.config(state="disabled")
        var_info.set(texto)
        root.config(cursor="watch")

        def hilo():
            try:
                if estado["af"] is None:
                    avisar(f"Conectando a {var_servidor.get()}...")
                    estado["af"] = AF(var_servidor.get().strip(),
                                      var_base.get().strip())
                    guardar_preferencias(servidor=var_servidor.get().strip(),
                                         base=var_base.get().strip())
                resultado = trabajo(estado["af"])
                root.after(0, fin, resultado, None)
            except Exception as error:  # noqa: BLE001 - se muestra
                root.after(0, fin, None, error)

        def fin(resultado, error):
            estado["ocupado"] = False
            root.config(cursor="")
            for boton in botones_af:
                boton.config(state="normal")
            if error is not None:
                var_info.set("No se pudo.")
                mostrar_error("PI AF", str(error))
                return
            despues(resultado)

        threading.Thread(target=hilo, daemon=True).start()

    # ---- buscar -----------------------------------------------------
    def llenar_tabla(resultados):
        tabla.delete(*tabla.get_children())
        for elem, nota in resultados:
            tabla.insert("", "end", values=(
                elem["nombre"], elem["pcp"], "sí" if elem["soc"] else "",
                "" if nota is None else f"{nota} %", elem["ruta"]))

    def mostrar(*_):
        elementos = estado["elementos"]
        if not elementos:
            return
        if not limpiar(var_texto.get()):
            tabla.delete(*tabla.get_children())
            var_info.set(f"{len(elementos):,} elementos SAE* en la lista. "
                         f"Escribe un nombre.")
            return
        resultados = buscar(elementos, var_texto.get())
        llenar_tabla(resultados)
        var_info.set(
            f"{len(resultados)} parecido(s) de {len(elementos):,}. Si no "
            f"esta, prueba 'Buscar en toda la base' o la pestaña Navegar."
            if resultados else
            "Ninguno se parece. Prueba con menos palabras, 'Buscar en "
            "toda la base' o la pestaña Navegar.")

    def actualizar_lista():
        def guardar(elementos):
            estado["elementos"] = elementos
            try:
                ruta_lista().write_text(json.dumps(
                    {"servidor": var_servidor.get(), "base": var_base.get(),
                     "fecha": date.today().isoformat(),
                     "elementos": elementos},
                    ensure_ascii=False, indent=1), encoding="utf-8")
            except OSError:
                pass
            mostrar()
            var_info.set(f"{len(elementos):,} elementos SAE* leidos de AF. "
                         + var_info.get())

        con_af(lambda af: af.lista_sae(avisar=avisar), guardar,
               "Leyendo los elementos SAE* (puede tardar un poco)...")

    def en_toda_la_base():
        texto = var_texto.get().strip()
        if not texto:
            var_info.set("Escribe parte del nombre del elemento.")
            return

        def mostrar_base(elementos):
            notas = dict((id(e), n) for e, n in buscar(elementos, texto,
                                                        limite=10 ** 6))
            elementos.sort(key=lambda e: -notas.get(id(e), 0))
            llenar_tabla([(e, notas.get(id(e))) for e in elementos])
            pestanas.select(0)
            var_info.set(f"{len(elementos)} elemento(s) con '{texto}' en el "
                         f"nombre (maximo {MAXIMO_BUSQUEDA_BASE}).")

        con_af(lambda af: af.en_toda_la_base(texto, avisar=avisar),
               mostrar_base, f"Buscando '*{texto}*' en toda la base...")

    # ---- elegir un elemento -----------------------------------------
    def refrescar_guardar(*_):
        listo = bool(estado["central"] and var_ruta.get())
        boton_guardar.config(state="normal" if listo else "disabled")

    var_ruta.trace_add("write", refrescar_guardar)

    def elegir(ruta, elem_af=None):
        var_ruta.set(ruta)
        estado["elegido"] = ruta
        attrs.delete(*attrs.get_children())
        if not ruta or estado["ocupado"]:
            return

        def leer(af):
            elem = elem_af
            if elem is None:
                camino = af.por_ruta(ruta)
                if not camino:
                    raise RuntimeError(f"No se encontro {ruta} en AF.")
                elem = camino[-1]
            return af.atributos(elem)

        def pintar(filas):
            if estado["elegido"] != ruta:
                return
            for nombre, valor in filas:
                attrs.insert("", "end", text=nombre, values=(valor,),
                             tags=("clave",) if nombre in (
                                 ATRIBUTO_NOMBRE, ATRIBUTO_SOC) else ())
            tiene = any(n == ATRIBUTO_SOC for n, _ in filas)
            var_info.set("Tiene 'Porc Carga Disponible': sirve para el SoC."
                         if tiene else
                         "Este elemento NO tiene 'Porc Carga Disponible'.")

        con_af(leer, pintar, "Leyendo atributos...")

    def al_elegir_en_tabla(*_):
        sel = tabla.selection()
        if sel:
            elegir(tabla.item(sel[0], "values")[4])

    def copiar(*_):
        if var_ruta.get():
            root.clipboard_clear()
            root.clipboard_append(var_ruta.get())
            var_info.set(f"Copiado (pegalo en SoC_PI): {var_ruta.get()}")

    # ---- navegar ------------------------------------------------------
    def insertar_nodo(padre, elem):
        ruta = str(elem.GetPath())
        nombre = str(elem.Name)
        iid = arbol.insert(padre, "end", text=nombre,
                           tags=("sae",) if limpiar(nombre).startswith("SAE")
                           else ())
        estado["nodos"][iid] = {"elem": elem, "ruta": ruta, "cargado": False}
        if elem.HasChildren:
            arbol.insert(iid, "end", text="...")
        return iid

    def cargar_hijos(iid):
        nodo = estado["nodos"].get(iid)
        if nodo is None or nodo["cargado"]:
            return
        nodo["cargado"] = True
        arbol.delete(*arbol.get_children(iid))
        root.config(cursor="watch")
        root.update_idletasks()
        try:
            hijos = sorted(AF.hijos(nodo["elem"]), key=lambda e: str(e.Name))
            for hijo in hijos:
                insertar_nodo(iid, hijo)
        finally:
            root.config(cursor="")

    def raices():
        if arbol.get_children(""):
            return

        def pintar(elems):
            for elem in sorted(elems, key=lambda e: str(e.Name)):
                insertar_nodo("", elem)

        con_af(lambda af: af.raices(), pintar, "Leyendo la raiz de la base...")

    def al_abrir(*_):
        iid = arbol.focus()
        if iid:
            cargar_hijos(iid)

    def al_elegir_en_arbol(*_):
        sel = arbol.selection()
        if sel and sel[0] in estado["nodos"]:
            nodo = estado["nodos"][sel[0]]
            elegir(nodo["ruta"], nodo["elem"])

    def ver_en_arbol():
        ruta = var_ruta.get()
        if not ruta:
            return

        def abrir(camino):
            if not camino:
                var_info.set(f"No se encontro {ruta} en AF.")
                return
            if not arbol.get_children(""):
                for elem in sorted(estado["af"].raices(),
                                   key=lambda e: str(e.Name)):
                    insertar_nodo("", elem)
            iid = ""
            for elem in camino:
                if iid:
                    cargar_hijos(iid)
                nombre = str(elem.Name)
                iid = next((h for h in arbol.get_children(iid)
                            if arbol.item(h, "text") == nombre), None)
                if iid is None:
                    return
                arbol.item(iid, open=True)
            cargar_hijos(iid)
            pestanas.select(1)
            arbol.selection_set(iid)
            arbol.see(iid)
            arbol.focus(iid)

        con_af(lambda af: af.por_ruta(ruta), abrir, "Abriendo el arbol...")

    def al_cambiar_pestana(*_):
        if pestanas.index("current") == 1:
            raices()

    # ---- Centrales.xlsx ------------------------------------------------
    def pintar_centrales(ruta, elegir_central=None):
        estado["centrales"] = ruta
        estado["central"] = None
        var_central.set("(elegi una central a la izquierda)")
        refrescar_guardar()
        lista_centrales.delete(*lista_centrales.get_children())
        try:
            centrales = centrales_del_diccionario(ruta)
        except Exception as error:  # noqa: BLE001 - se muestra
            mostrar_error("Centrales.xlsx", str(error))
            return
        if not centrales:
            var_info.set("El Diccionario no tiene la tabla con encabezado "
                         "Balance_BESS.")
            return
        sin = 0
        for central, ruta_soc in sorted(centrales,
                                        key=lambda c: (bool(c[1]), c[0])):
            sin += not ruta_soc
            iid = lista_centrales.insert(
                "", "end", text=("✓ " if ruta_soc else "✗ ") + central,
                values=(ruta_soc,), tags=("con" if ruta_soc else "sin",))
            if elegir_central == central:
                lista_centrales.selection_set(iid)
                lista_centrales.see(iid)
        var_info.set(f"{len(centrales)} centrales, {sin} sin ruta en SoC_PI "
                     f"(en rojo). Clic en una para ver su ruta o buscarla.")

    def elegir_centrales():
        ruta = filedialog.askopenfilename(
            title="Centrales.xlsx", filetypes=[("Excel", "*.xlsx")],
            initialfile=prefs.get("centrales", ""))
        if ruta:
            guardar_preferencias(centrales=ruta)
            pintar_centrales(ruta)

    def al_elegir_central(*_):
        sel = lista_centrales.selection()
        if not sel:
            return
        texto = lista_centrales.item(sel[0], "text")[2:]
        ruta_soc = (lista_centrales.item(sel[0], "values") or ("",))[0]
        if estado["central"] == texto:
            return
        estado["central"] = texto
        var_central.set(f"-> {texto}")
        pestanas.select(0)
        var_texto.set(texto)
        if ruta_soc:
            # La que ya tiene: se muestra con sus atributos.
            elegir(ruta_soc)
            var_info.set(f"{texto} ya tiene ruta en SoC_PI. Si encontras "
                         f"otra mejor, elegila y apreta Guardar.")
            return
        filas = tabla.get_children()
        if filas:
            # La posible ruta: el resultado mas parecido.
            tabla.selection_set(filas[0])
            tabla.see(filas[0])
            var_info.set(f"Posible ruta para {texto}: la primera de la "
                         f"lista. Revisala y apreta 'Guardar en "
                         f"Centrales.xlsx'.")
        else:
            var_ruta.set("")
            var_info.set(f"No hay parecidos para {texto} en la lista SAE*. "
                         f"Proba 'Buscar en toda la base' o Navegar.")

    def guardar_en_centrales():
        central, ruta = estado["central"], var_ruta.get().strip()
        archivo = estado["centrales"]
        if not (central and ruta and archivo):
            return
        actual = next(
            ((lista_centrales.item(i, "values") or ("",))[0]
             for i in lista_centrales.get_children()
             if lista_centrales.item(i, "text")[2:] == central), "")
        if actual == ruta:
            var_info.set(f"{central} ya tiene esa ruta.")
            return
        pregunta = f"Guardar en SoC_PI de {central}:\n\n{ruta}"
        if actual:
            pregunta += f"\n\nReemplaza a:\n{actual}"
        if not messagebox.askyesno("Guardar en Centrales.xlsx", pregunta):
            return
        try:
            guardar_ruta_en_centrales(archivo, central, ruta)
        except CentralesAbierto as error:
            mostrar_error("Centrales.xlsx abierto", str(error))
            return
        except Exception as error:  # noqa: BLE001 - se muestra
            mostrar_error("Centrales.xlsx", str(error))
            return
        pintar_centrales(archivo, elegir_central=central)
        estado["central"] = central
        var_central.set(f"-> {central}")
        refrescar_guardar()
        var_info.set(f"Guardado en Centrales.xlsx: {central} -> {ruta}")

    # ---- arranque ----------------------------------------------------
    try:
        guardada = json.loads(ruta_lista().read_text(encoding="utf-8"))
        estado["elementos"] = guardada["elementos"]
        var_info.set(f"{len(estado['elementos']):,} elementos SAE* (lista "
                     f"del {guardada.get('fecha', '?')}). Escribe un nombre.")
    except (OSError, ValueError, KeyError):
        var_info.set("Aprieta 'Actualizar lista SAE*' (solo la primera vez) "
                     "o usa la pestaña Navegar.")

    if prefs.get("centrales") and Path(prefs["centrales"]).is_file():
        pintar_centrales(prefs["centrales"])

    boton_lista.config(command=actualizar_lista)
    boton_base.config(command=en_toda_la_base)
    boton_centrales.config(command=elegir_centrales)
    boton_copiar.config(command=copiar)
    boton_ir.config(command=ver_en_arbol)
    boton_guardar.config(command=guardar_en_centrales)
    tabla.bind("<<TreeviewSelect>>", al_elegir_en_tabla)
    tabla.bind("<Double-1>", copiar)
    arbol.bind("<<TreeviewOpen>>", al_abrir)
    arbol.bind("<<TreeviewSelect>>", al_elegir_en_arbol)
    arbol.bind("<Double-1>", copiar)
    lista_centrales.bind("<<TreeviewSelect>>", al_elegir_central)
    pestanas.bind("<<NotebookTabChanged>>", al_cambiar_pestana)
    def al_enter(*_):
        filas = tabla.get_children()
        if not filas:
            en_toda_la_base()
        else:
            tabla.selection_set(filas[0])
            tabla.focus(filas[0])

    entrada.bind("<Return>", al_enter)
    var_texto.trace_add("write", mostrar)
    entrada.focus_set()
    root.mainloop()


if __name__ == "__main__":
    main()
