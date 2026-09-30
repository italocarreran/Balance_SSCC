# -*- coding: utf-8 -*-
"""
Balance BESS / SSCC.

Ventana unica: se elige la carpeta base del caso, se ingresa el
periodo (AAMM) y debajo se dibuja el diagrama de la estructura del
caso con el estado de cada entrada (OK/FALTA/PENDIENTE).

Cada periodo trabaja en SU carpeta: dos meses no comparten una. Al
cambiar el AAMM, si ese periodo ya tuvo carpeta se vuelve a ella sola;
si no, se pregunta cual es (examinar / crear / cancelar), en vez de
seguir sobre la del mes anterior. Tambien esta el boton "Crear carpeta
del caso", abajo de todo; y a un caso al que le falte una subcarpeta
se la completa desde ahi mismo.

Todo lo que el programa puede hacer sale de un boton en la fila que
corresponde. Ademas, abajo de todo hay un boton "Ejecutar todo": abre
una ventana con el PLAN de la corrida (que falta, que esta al dia, que
se puede rehacer y que bloquea la corrida) y lo ejecuta respetando las
dependencias, en paralelo donde se puede. El grafo y el plan viven en
Script/nucleo/orquestador.py; esta ventana solo los dibuja.

    <CARPETA_BASE>/
        Medidas/
            Medidas_SAE.xlsx               [Actualizar]
            <algo>SOC<algo>AAMM<algo>.xlsx
        Auxiliares/
            Centrales.xlsx
                hoja 'Resumen BESS'
                hoja 'Diccionario'
            <algo>Homologacion<algo>.xlsx
                hoja 'homol'
                hoja 'Gen real'
        Ofertas/
            <algo>OfertasSSCC<algo>.xlsx (o .xlsm/.xlsb/.xls)
        Cmg/
            cmg<AAMM>_def_15minutal.csv    [Traer]
            cmg.xlsx                       [Generar]
        FD y FMA/
            SSCC_Desempeño_<algo>.xlsx     [Traer]
            fma_cpf_<AAMM>.xlsx            [Generar]
            fma_csf_<AAMM>.xlsx            [Generar]
            fma_cft_<AAMM>.xlsx            [Generar]
        Subastas/
            DB subastas/                   [Traer]
        Balance_BESS.xlsx                  [Actualizar]
            hoja 'Resumen'                 [Asignar]
            hoja 'PRORRATA_RETIROS'        [Traer]
            hoja 'COMPENSACION_CENTRAL'    [Resumir]
            hoja 'Calculo RE545'           [Calcular]
            hoja 'Calculo E Costos'        [Calcular]
            hoja 'Subastas'                [Actualizar]
            hoja 'FD'                      [Actualizar]
            hoja 'CMg'                     [Actualizar]
            hoja 'Ofertas SSCC'            [Actualizar]
            hoja 'Medidores'               [Actualizar]
        Control_corrida.xlsx

La salida es UNA sola planilla (antes eran dos,
Consolidado_entradas.xlsx y Pagos_BESS.xlsx) y sus hojas quedan
ordenadas de FIN A INICIO: arriba el Resumen -lo primero que se mira-
y abajo las entradas de las que sale todo. Las hojas de control de la
corrida (Alertas, Ejecucion, Log) ya no estan ahi adentro: viven en
Control_corrida.xlsx, al lado.

La planilla se desglosa por hoja igual que Centrales.xlsx: cada hoja
se actualiza sola, y lo que no se toca se conserva tal cual estaba en
el archivo. Si el archivo todavia no existe, se crea al actualizar la
primera hoja.

Cada archivo y cada carpeta del diagrama es un LINK a su ruta: el
click abre la carpeta en el explorador (la que contiene al archivo, si
la fila es un archivo -- nunca se abre el archivo). El diagrama no
lleva columna de detalle (se saco a pedido del usuario): lo unico que
queda a la derecha del boton es, en las filas que traen algo de
afuera del caso, de donde viene:

    SSCC_Desempeño_*      Origen: DCO
    cmg<AAMM>_..._.csv    Origen: CMg Reales
    DB subastas/          Origen: progdiar_adjudicaSEN
    fma_cpf/csf/cft       Origen inputs: ... (el FMA no se trae hecho,
                          se arma, pero sus insumos si vienen de afuera)

y ese nombre tambien es un link, a la carpeta exacta de origen del
periodo (ver Script/nucleo/origenes.py).

Medidas/_trabajo/ (los lotes que baja la API, la marca de
reanudacion) NO aparece en el diagrama a pedido del usuario: no es una
entrada ni una salida del caso, son andamios del proceso.

El calculo vive en Script/ (ver Script/__init__.py). Las dos claves de
las APIs del Coordinador que usa Medidas NO estan en el codigo: salen
de la seccion "claves_api" de config.json (ver Script/config.py y
config.ejemplo.json). La ubicacion de este .py no influye en nada salvo
en donde se guarda config.json.
"""

import os
import re
import socket
import subprocess
import sys
import threading
import time
import traceback
from pathlib import Path

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from Script import config, nucleo


COLOR_OK = "#1a6b1a"
COLOR_FALTA = "#b00020"
COLOR_PENDIENTE = "#a06000"
COLOR_NEUTRO = "#555555"
COLOR_BASE_OK = "#1a4fb0"
COLOR_LINK = "#1a4fb0"

SIMBOLO = {
    "ok": "OK",
    "falta": "FALTA",
    "pendiente": "PENDIENTE",
}

COLOR_ESTADO = {
    "ok": COLOR_OK,
    "falta": COLOR_FALTA,
    "pendiente": COLOR_PENDIENTE,
}

# Anchos de las columnas del diagrama, TODOS en pixeles y cada celda
# en un Frame de ancho fijo (pack_propagate(False)).
#
# Antes estaban en "caracteres" (width=N de un Label) y las columnas
# salian torcidas: una fila en negrita mide mas que la misma fila en
# redonda aunque las dos digan 52 caracteres, asi que el estado y el
# boton de las filas de nivel 0 quedaban corridos a la derecha
# respecto del resto. En pixeles no hay forma de que se corran: la
# celda mide lo mismo diga lo que diga adentro.
ANCHO_ESTRUCTURA = 420     # pixeles
ANCHO_ESTADO = 86          # pixeles
ANCHO_ACCION = 104         # pixeles
ALTO_ACCION = 26           # pixeles (alto de una fila CON boton)
ALTO_FILA = 18             # pixeles (alto de una fila sin boton)

# El boton ocupa su celda entera: todos los botones del diagrama miden
# exactamente lo mismo y quedan en una columna recta (pedido del
# usuario). Por eso los textos son de una palabra ("Traer",
# "Resumir", "Generar", "Actualizar", "Calcular").
RELLENO_ACCION = 4         # pixeles de aire alrededor del boton

# Rueda del mouse. La ventana entera se desplaza de a PIXELES, no de a
# "unidades" del Canvas: una unidad de Canvas sin yscrollincrement es
# un decimo del alto visible, y ese salto es lo que se sentia pegado.
# El registro, en cambio, se desplaza de a lineas, que es lo natural en
# un widget de texto.
PIXELES_POR_MUESCA = 45    # cuanto baja la ventana por cada muesca
LINEAS_POR_MUESCA = 3      # cuantas lineas baja el registro por muesca
# Windows/macOS mandan <MouseWheel>; X11 manda los botones 4 y 5.
EVENTOS_RUEDA = ("<MouseWheel>", "<Button-4>", "<Button-5>")


# ============================================================
# CONFIG POR PC/USUARIO
# ============================================================

def get_usuario():
    usuario = (
        os.environ.get("USERNAME")
        or os.environ.get("USER")
        or "desconocido"
    )
    return f"{socket.gethostname()}_{usuario}"


def leer_config():
    """Lo que dejo guardado ESTE PC/usuario (carpeta base, AAMM)."""

    return config.seccion(get_usuario())


def guardar_config(data):
    """
    Guarda en la seccion de este PC/usuario, sin tocar el resto del
    archivo -- ni las secciones de otros, ni "claves_api".
    """

    config.actualizar_seccion(get_usuario(), data)


# La carpeta que uso este PC/usuario para cada periodo:
# {"2607": "D:/...Balance BESS 2607", "2608": ...}. Es lo que permite
# cumplir la regla del usuario -"dos meses no pueden tener la misma
# carpeta"- incluso cuando el nombre de la carpeta no dice el AAMM:
# una vez que una carpeta quedo asociada a un periodo, cambiar de mes
# ya no la puede reusar en silencio.
CLAVE_CARPETAS_PERIODO = "carpetas_por_periodo"


def carpetas_por_periodo():
    """{aamm: carpeta} de este PC/usuario. {} si todavia no hay ninguna."""

    datos = leer_config().get(CLAVE_CARPETAS_PERIODO)

    if not isinstance(datos, dict):
        return {}

    return {
        str(periodo): str(ruta)
        for periodo, ruta in datos.items()
        if isinstance(ruta, str) and ruta
    }


def carpeta_recordada(aamm):
    """La carpeta que se uso para ese periodo, si todavia existe."""

    ruta = carpetas_por_periodo().get(str(aamm), "")

    try:
        return ruta if ruta and Path(ruta).is_dir() else ""
    except OSError:
        return ""


def recordar_carpeta(aamm, ruta):
    """
    Deja anotado que 'ruta' es la carpeta del periodo 'aamm'. Si esa
    carpeta estaba anotada como la de OTRO periodo, se le saca: una
    carpeta pertenece a un solo periodo.
    """

    if not aamm or not ruta:
        return

    ruta = str(ruta)

    carpetas = {
        periodo: otra
        for periodo, otra in carpetas_por_periodo().items()
        if periodo == str(aamm) or otra != ruta
    }
    carpetas[str(aamm)] = ruta

    guardar_config({CLAVE_CARPETAS_PERIODO: carpetas})


def carpeta_a_abrir(ruta, es_archivo=False):
    """
    Que carpeta se abre al hacer click en una fila del diagrama: la
    propia si la fila es una carpeta, la que CONTIENE al archivo si es
    un archivo -- nunca se abre el archivo, para no arrancar Excel sin
    que se lo pidan.

    Si todavia no existe (una carpeta que falta, un archivo que aun no
    se genero, el mes que el DCO no publico), se sube hasta el primer
    ancestro que si exista: asi el click siempre lleva a algun lado.
    Devuelve None si no existe ni la raiz (unidad desconectada).
    """

    carpeta = Path(ruta)

    if es_archivo:
        carpeta = carpeta.parent

    while True:

        try:
            if carpeta.is_dir():
                return carpeta
        except OSError:
            return None

        padre = carpeta.parent

        if padre == carpeta:
            return None

        carpeta = padre


def abrir_en_explorador(ruta, es_archivo=False):
    """
    Abre en el explorador la carpeta de esa ruta. Devuelve la carpeta
    que abrio, o None si no habia ninguna que abrir.
    """

    if not ruta:
        return None

    carpeta = carpeta_a_abrir(ruta, es_archivo)

    if carpeta is None:
        return None

    if sys.platform == "win32":
        subprocess.Popen(["explorer", str(carpeta)])
    elif sys.platform == "darwin":
        subprocess.Popen(["open", str(carpeta)])
    else:
        subprocess.Popen(["xdg-open", str(carpeta)])

    return carpeta


def formato_tiempo(segundos):
    minutos, seg = divmod(int(segundos), 60)
    horas, minutos = divmod(minutos, 60)
    return f"{horas:02d}:{minutos:02d}:{seg:02d}"


# ============================================================
# DIAGRAMA DE CARPETAS (arbol de texto tipo consola)
#
# revisar_estructura() devuelve una lista plana de filas con su
# NIVEL (0 = raiz del caso, 1 = adentro de una carpeta/archivo, 2 =
# un nivel mas). Aca solo se traduce esa lista a prefijos de consola
# (├── / └── / │): nucleo.py sabe de estructura, no de dibujo.
# ============================================================

def _es_ultimo_en_su_nivel(niveles, i):
    """
    True si, mirando hacia adelante desde i, se sube de nivel antes
    de encontrar otra fila con el MISMO nivel (o se llega al final de
    la lista): o sea, si i es la ultima de su grupo.
    """
    nivel = niveles[i]
    for j in range(i + 1, len(niveles)):
        if niveles[j] < nivel:
            return True
        if niveles[j] == nivel:
            return False
    return True


def _prefijos_arbol(niveles):
    """
    Prefijos tipo consola (├── / └── / │) para una lista plana de
    niveles (0 = raiz), calculando el relleno de cada ancestro segun
    si ESE ancestro es o no el ultimo de su propio grupo.
    """
    n = len(niveles)
    ultimos = [_es_ultimo_en_su_nivel(niveles, i) for i in range(n)]
    prefijos = []

    for i, nivel in enumerate(niveles):

        if nivel == 0:
            prefijos.append("")
            continue

        relleno = ""
        for ancestro_nivel in range(1, nivel):
            indice_ancestro = next(
                (k for k in range(i - 1, -1, -1) if niveles[k] == ancestro_nivel),
                None,
            )
            relleno += (
                "    "
                if (indice_ancestro is None or ultimos[indice_ancestro])
                else "│   "
            )

        prefijos.append(relleno + ("└── " if ultimos[i] else "├── "))

    return prefijos


# ============================================================
# VENTANA
# ============================================================

def main():

    cfg = leer_config()

    root = tk.Tk()
    root.title("Balance BESS / SSCC")
    root.geometry("1180x820")

    var_base = tk.StringVar(value=cfg.get("carpeta_base", ""))
    var_aamm = tk.StringVar(value=cfg.get("aamm", ""))
    var_estado = tk.StringVar(value="Listo")
    var_tiempo = tk.StringVar(value="00:00:00")

    # Los botones viven DENTRO de las filas del arbol, que se repinta
    # entero en cada revisar(): se guardan por id de fila para poder
    # deshabilitarlos mientras corre algo, y se renuevan en cada
    # pintado.
    botones_arbol = {}
    corriendo = {"activo": False}

    # --------------------------------------------------------
    # BOTONES FIJOS ABAJO (primero, para que no los tape nada)
    # --------------------------------------------------------

    frame_botones = tk.Frame(root)
    frame_botones.pack(side="bottom", fill="x", pady=10)

    # --------------------------------------------------------
    # CANVAS CON SCROLL
    # --------------------------------------------------------

    # yscrollincrement=1: una 'unidad' de desplazamiento del Canvas pasa
    # a ser UN PIXEL. Sin esto, una unidad es un decimo del alto
    # visible y la rueda mueve la ventana a los saltos.
    canvas = tk.Canvas(
        root, borderwidth=0, highlightthickness=0, yscrollincrement=1,
    )
    scroll = tk.Scrollbar(root, orient="vertical", command=canvas.yview)
    canvas.configure(yscrollcommand=scroll.set)
    scroll.pack(side="right", fill="y")
    canvas.pack(side="left", fill="both", expand=True)

    contenedor = tk.Frame(canvas)
    ventana_canvas = canvas.create_window((0, 0), window=contenedor, anchor="nw")

    def ajustar(event=None):
        canvas.configure(scrollregion=canvas.bbox("all"))
        canvas.itemconfig(ventana_canvas, width=canvas.winfo_width())

    contenedor.bind("<Configure>", ajustar)
    canvas.bind("<Configure>", ajustar)

    # --------------------------------------------------------
    # RUEDA DEL MOUSE
    #
    # Dos problemas que tenia el binding de antes
    # (bind_all("<MouseWheel>") -> canvas.yview_scroll(int(-delta/120))):
    #
    # 1. Era global: la rueda movia la ventana entera aunque el puntero
    #    estuviera sobre el registro. El registro no se podia recorrer
    #    sin recorrer toda la ventana.
    # 2. int(-delta/120) trunca hacia cero. Un touchpad manda deltas
    #    chicos (30, 40...) y todos daban 0: la ventana no se movia
    #    hasta que el gesto era grande, y ahi saltaba de golpe. Eso es
    #    lo que se sentia "pegado". Ahora el resto de cada evento se
    #    acumula y nada se pierde.
    #
    # El widget que se desplaza es el que esta DEBAJO DEL PUNTERO: si
    # es el registro (o algo adentro de el), se desplaza el registro y
    # nada mas; si no, la ventana.
    desplazables = {}   # widget que recibe la rueda -> como desplazarlo
    resto_rueda = {"ventana": 0.0, "registro": 0.0}

    def registrar_desplazable(widget, desplazar):
        """Un widget que se queda con la rueda cuando el puntero esta encima."""

        desplazables[str(widget)] = desplazar
        # Ademas del registro en el diccionario, la rueda se ata al
        # widget mismo: las ataduras de widget corren ANTES que las de
        # clase, asi que el "break" de rueda() evita que la atadura de
        # clase de Text (que tambien desplaza) lo mueva una segunda vez.
        for nombre in EVENTOS_RUEDA:
            widget.bind(nombre, rueda)

    def _acumular(clave, cantidad):
        """Parte entera a desplazar, guardando el resto para el proximo evento.

        Un touchpad manda deltas chicos: sin acumular el resto, cada
        evento se redondea a 0 y no se mueve nada.
        """

        resto = resto_rueda[clave]
        # Al cambiar de sentido el resto pendiente ya no sirve: seguirlo
        # sumando frenaria medio pixel el primer evento del otro lado.
        if resto and (resto > 0) != (cantidad > 0):
            resto = 0.0
        total = cantidad + resto
        entero = int(total)
        resto_rueda[clave] = total - entero
        return entero

    def _desplazar_ventana(muescas):
        # scrollregion mas chica que lo visible = no hay nada que
        # desplazar; sin esto la ventana "tiembla" con la rueda.
        region = canvas.bbox("all")
        if not region or region[3] - region[1] <= canvas.winfo_height():
            return
        pixeles = _acumular("ventana", muescas * PIXELES_POR_MUESCA)
        if pixeles:
            canvas.yview_scroll(pixeles, "units")

    def _muescas(event):
        """Muescas de rueda de este evento (negativo = hacia arriba)."""

        # X11 no manda delta: usa los botones 4 (arriba) y 5 (abajo).
        if getattr(event, "num", None) in (4, 5):
            return -1.0 if event.num == 4 else 1.0
        delta = getattr(event, "delta", 0)
        if not delta:
            return 0.0
        # En macOS el delta ya viene en muescas. En Windows viene en
        # 120avos de muesca, y un touchpad de precision manda valores
        # MENORES a 120 (40, 60...) que son fracciones de muesca: hay
        # que dividir igual y acumular el resto, no tratarlos como
        # muescas enteras (serian 40 muescas de un saque).
        if sys.platform == "darwin":
            return -float(delta)
        return -delta / 120.0

    def rueda(event):
        muescas = _muescas(event)
        if not muescas:
            return "break"
        widget = event.widget
        try:
            debajo = root.winfo_containing(event.x_root, event.y_root)
        except tk.TclError:
            debajo = None
        widget = debajo if debajo is not None else widget
        # El puntero puede estar sobre un hijo (la barra del registro,
        # por ejemplo): se sube por los padres hasta encontrar a alguien
        # que se haya anotado para recibir la rueda.
        while widget is not None:
            desplazar = desplazables.get(str(widget))
            if desplazar is not None:
                desplazar(muescas)
                return "break"
            widget = getattr(widget, "master", None)
        _desplazar_ventana(muescas)
        return "break"

    for nombre in EVENTOS_RUEDA:
        root.bind_all(nombre, rueda)
        canvas.bind(nombre, rueda)

    # --------------------------------------------------------
    # SELECTOR DE CARPETA BASE
    # --------------------------------------------------------

    frame_carpeta = tk.LabelFrame(
        contenedor, text="Carpeta base del caso", padx=10, pady=8
    )
    frame_carpeta.pack(fill="x", padx=20, pady=(14, 6))

    lbl_base = tk.Label(
        frame_carpeta,
        textvariable=var_base,
        wraplength=940,
        justify="center",
        cursor="hand2",
        font=("Segoe UI", 9),
    )
    lbl_base.pack()

    lbl_base.bind(
        "<Button-1>",
        lambda e: abrir_en_explorador(var_base.get()) if var_base.get() else None,
    )

    # --------------------------------------------------------
    # PERIODO DEL CASO (AAMM)
    #
    # Se ingresa aca arriba y ya no aparece como fila del diagrama:
    # no es parte de la estructura de carpetas. De el dependen dos
    # filas (el SoC dentro de Medidas/ y el CSV dentro de Cmg/), que
    # lo dicen en su propio detalle cuando falta.
    # --------------------------------------------------------

    frame_periodo = tk.LabelFrame(
        contenedor, text="Periodo del caso (AAMM)", padx=10, pady=8
    )
    frame_periodo.pack(fill="x", padx=20, pady=6)

    validador_aamm = (
        root.register(lambda s: s == "" or (s.isdigit() and len(s) <= 4)),
        "%P",
    )

    entry_aamm = tk.Entry(
        frame_periodo,
        textvariable=var_aamm,
        width=10,
        justify="center",
        font=("Segoe UI", 10, "bold"),
        validate="key",
        validatecommand=validador_aamm,
    )
    entry_aamm.pack(side="left")

    tk.Label(
        frame_periodo,
        text=(
            "4 digitos: año+mes simplificado. Ej. 2607 para julio de "
            "2026. Con el se buscan el archivo de SoC dentro de "
            "Medidas/ y el CSV 15-minutal dentro de Cmg/."
        ),
        fg=COLOR_NEUTRO,
        font=("Segoe UI", 8),
        wraplength=820,
        justify="left",
    ).pack(side="left", padx=(10, 0))

    # --------------------------------------------------------
    # DIAGRAMA DE LA ESTRUCTURA DEL CASO
    # --------------------------------------------------------

    frame_arbol = tk.LabelFrame(
        contenedor, text="Estructura del caso", padx=8, pady=8
    )
    frame_arbol.pack(fill="both", expand=True, padx=20, pady=6)

    enc_arbol = tk.Frame(frame_arbol)
    enc_arbol.pack(fill="x")

    # Las mismas celdas de ancho fijo que las filas, para que el
    # titulo de cada columna caiga justo arriba de su columna.
    for texto, ancho in (
        ("Estructura", ANCHO_ESTRUCTURA),
        ("Estado", ANCHO_ESTADO),
        ("Accion", ANCHO_ACCION),
    ):
        celda = tk.Frame(enc_arbol, width=ancho, height=16)
        celda.pack(side="left")
        celda.pack_propagate(False)
        tk.Label(
            celda, text=texto, font=("Segoe UI", 8, "bold"), anchor="w",
        ).pack(side="left")

    tk.Label(
        enc_arbol, text="Origen", font=("Segoe UI", 8, "bold"), anchor="w",
    ).pack(side="left")

    filas_arbol = tk.Frame(frame_arbol)
    filas_arbol.pack(fill="x")

    def abrir_ruta_de_fila(ruta, es_carpeta):
        """
        Click en el nombre de una fila: abre SU carpeta (la propia si
        es una carpeta, la que la contiene si es un archivo). Nunca
        abre el archivo.
        """

        abierta = abrir_en_explorador(ruta, es_archivo=not es_carpeta)

        if abierta is None:
            log(f"No se pudo abrir {ruta}")

    def abrir_origen(id_origen):
        """
        Click en el link de "Origen: ..." / "Origen inputs: ...": abre
        la carpeta de afuera del caso de la que sale esa entrada.

        Se resuelve en un hilo aparte porque mirar el arbol del DCO (o
        cualquiera de las otras unidades de red) puede tardar segundos
        cuando la unidad no esta conectada, y no vale la pena congelar
        la ventana por un click.
        """

        aamm = var_aamm.get().strip()

        def trabajo():

            try:
                ruta = nucleo.ruta_origen(id_origen, aamm)
            except Exception as error:
                root.after(0, log, f"No se pudo resolver el origen: {error}")
                return

            root.after(0, log, f"Origen: {ruta}")

            if abrir_en_explorador(ruta, es_archivo=False) is None:
                root.after(
                    0, log,
                    f"No se pudo abrir {ruta} (revisa que la unidad de "
                    f"red este conectada).",
                )

        threading.Thread(target=trabajo, daemon=True).start()

    def _celda(parent, ancho, alto):
        """Una celda de ancho fijo: lo de adentro no mueve la columna."""

        celda = tk.Frame(parent, width=ancho, height=alto)
        celda.pack(side="left")
        celda.pack_propagate(False)
        return celda

    def _fila_arbol(parent, prefijo, texto, estado=None, negrita=False,
                    boton=None, id_fila=None, ruta="", es_carpeta=False,
                    origen=None):
        """
        Una fila del diagrama. Columnas, en orden: estructura, estado,
        accion (el boton, si la fila tiene uno) y origen.

        Cada una vive en una celda de ancho fijo EN PIXELES, asi que
        el estado y los botones quedan en una columna recta aunque la
        fila este en negrita o el nombre sea largo.

        El nombre de la fila (no el prefijo del arbol) es un link a su
        ruta cuando la fila tiene una. La ultima columna ya no trae el
        detalle (el usuario lo pidio sacar): queda solo el
        "Origen: <algo>" -tambien link- de las filas que traen algo de
        afuera del caso.
        """

        fila = tk.Frame(parent)
        fila.pack(fill="x", pady=1)

        estilo = "bold" if negrita else "normal"

        # Una fila sin boton no necesita el alto de un boton: asi el
        # diagrama no se estira al doble de largo.
        alto = ALTO_ACCION if boton is not None else ALTO_FILA

        celda_estructura = _celda(fila, ANCHO_ESTRUCTURA, alto)

        # El prefijo del arbol (├── / └── / │) va en su propia etiqueta
        # para que el subrayado del link tape solo el nombre.
        if prefijo:
            tk.Label(
                celda_estructura,
                text=prefijo,
                font=("Consolas", 9, estilo),
                anchor="w",
            ).pack(side="left")

        etiqueta = tk.Label(
            celda_estructura,
            text=texto,
            font=("Consolas", 9, f"{estilo} underline" if ruta else estilo),
            anchor="w",
            fg=COLOR_LINK if ruta else "black",
            cursor="hand2" if ruta else "",
        )
        etiqueta.pack(side="left")

        if ruta:
            etiqueta.bind(
                "<Button-1>",
                lambda e, r=ruta, c=es_carpeta: abrir_ruta_de_fila(r, c),
            )

        celda_estado = _celda(fila, ANCHO_ESTADO, alto)

        tk.Label(
            celda_estado,
            text=SIMBOLO.get(estado, estado) if estado is not None else "",
            anchor="w",
            fg=COLOR_ESTADO.get(estado, COLOR_NEUTRO),
            font=("Segoe UI", 9, "bold"),
        ).pack(side="left")

        celda_accion = _celda(fila, ANCHO_ACCION, alto)

        if boton is not None:
            texto_boton, comando = boton
            widget = tk.Button(
                celda_accion, text=texto_boton, font=("Segoe UI", 8, "bold"),
                bg="#fdf0d5", command=comando,
            )
            # Ocupa la celda entera: todos los botones salen del mismo
            # tamaño, uno debajo del otro y sin escalones.
            widget.pack(
                fill="both", expand=True,
                padx=(0, RELLENO_ACCION * 2), pady=RELLENO_ACCION // 2,
            )
            if id_fila is not None:
                botones_arbol[id_fila] = widget

        celda_origen = tk.Frame(fila)
        celda_origen.pack(side="left", fill="x", expand=True)

        if origen:

            tk.Label(
                celda_origen,
                text=f"{origen['titulo']}: ",
                anchor="w",
                fg=COLOR_NEUTRO,
                font=("Segoe UI", 8),
            ).pack(side="left")

            link = tk.Label(
                celda_origen,
                text=origen["etiqueta"],
                anchor="w",
                fg=COLOR_LINK,
                cursor="hand2",
                font=("Segoe UI", 8, "underline"),
            )
            link.pack(side="left")
            link.bind(
                "<Button-1>",
                lambda e, i=origen["id"]: abrir_origen(i),
            )

        return fila

    def _boton_de_fila(id_fila):
        """
        Que boton lleva cada fila del diagrama. El id lo pone
        nucleo.revisar_estructura(); la ventana decide que accion le
        cuelga, para que nucleo.py no sepa nada de botones.
        """

        if id_fila == "medidas_sae":
            return ("Actualizar", actualizar_medidas_sae)

        if id_fila == "cmg_csv":
            return ("Traer", traer_cmg_15min)

        if id_fila == "cmg_xlsx":
            return ("Generar", generar_cmg)

        # Un boton por cosa que produce: el FD lo trae la fila de su
        # archivo, y las tres salidas de FMA las arma un solo boton,
        # colgado de la primera de las tres.
        if id_fila == "sscc":
            return ("Traer", traer_fd)

        if id_fila.startswith("fma_"):
            tipo = id_fila.split("_", 1)[1]
            return ("Generar", lambda t=tipo: generar_fma(t))

        if id_fila == "db_subastas":
            return ("Traer", traer_subastas)

        # Los botones dicen SOLO el verbo (pedido del usuario): de que
        # hoja se trata ya lo dice la fila en la que esta el boton.
        if id_fila == "pagos:compensacion_central":
            return ("Resumir",
                    lambda: actualizar_pagos({"compensacion_central"}))

        if id_fila == "pagos:prorrata_retiros":
            return ("Traer", lambda: actualizar_pagos({"prorrata_retiros"}))

        if id_fila == "pagos:resumen":
            return ("Asignar", lambda: actualizar_pagos({"resumen"}))

        # Las dos mitades de la planilla, de un solo boton: primero las
        # entradas y despues el calculo (ese es el orden en que se
        # necesitan, aunque en el libro queden al reves).
        if id_fila == "salida":
            return ("Actualizar", actualizar_planilla_entera)

        if id_fila.startswith("consolidado:"):
            seccion = id_fila.split(":", 1)[1]
            return ("Actualizar", lambda s=seccion: actualizar_consolidado({s}))

        if id_fila.startswith("pagos:"):
            seccion = id_fila.split(":", 1)[1]
            return ("Calcular", lambda s=seccion: actualizar_pagos({s}))

        return None

    def pintar_arbol(filas):

        for hijo in filas_arbol.winfo_children():
            hijo.destroy()

        botones_arbol.clear()

        if not filas:
            tk.Label(
                filas_arbol,
                text="Selecciona una carpeta base.",
                fg=COLOR_NEUTRO,
                anchor="w",
            ).pack(fill="x")
            return

        prefijos = _prefijos_arbol([fila["nivel"] for fila in filas])

        for fila, prefijo in zip(filas, prefijos):
            _fila_arbol(
                filas_arbol,
                prefijo,
                fila["etiqueta"],
                estado=fila["estado"],
                negrita=(fila["nivel"] == 0),
                boton=_boton_de_fila(fila["id"]),
                id_fila=fila["id"],
                ruta=fila.get("ruta", ""),
                es_carpeta=fila.get("es_carpeta", False),
                origen=fila.get("origen"),
            )

        if corriendo["activo"]:
            habilitar_botones(False)

    def habilitar_botones(activos):
        for widget in botones_arbol.values():
            try:
                widget.config(state="normal" if activos else "disabled")
            except tk.TclError:
                pass

    # --------------------------------------------------------
    # PROGRESO Y LOG
    # --------------------------------------------------------

    frame_estado = tk.Frame(contenedor)
    frame_estado.pack(fill="x", padx=20, pady=(10, 2))

    tk.Label(
        frame_estado, textvariable=var_estado, anchor="w", font=("Segoe UI", 9),
    ).pack(side="left")

    tk.Label(
        frame_estado, textvariable=var_tiempo,
        font=("Consolas", 10, "bold"), fg="#2d7a2d",
    ).pack(side="right")

    barra = ttk.Progressbar(contenedor, mode="determinate", maximum=100)
    barra.pack(fill="x", padx=20, pady=(0, 8))

    frame_log = tk.LabelFrame(contenedor, text="Registro", padx=6, pady=6)
    frame_log.pack(fill="both", expand=True, padx=20, pady=6)

    txt_log = tk.Text(frame_log, height=14, font=("Consolas", 9), wrap="word")
    scroll_log = tk.Scrollbar(frame_log, command=txt_log.yview)
    txt_log.configure(yscrollcommand=scroll_log.set)
    scroll_log.pack(side="right", fill="y")
    txt_log.pack(side="left", fill="both", expand=True)

    # Con el puntero sobre el registro, la rueda mueve el registro y
    # NADA MAS: no arrastra la ventana. Tampoco cuando el registro ya
    # esta en un extremo -- encadenar ahi es justo lo que hacia que
    # recorrer el registro terminara moviendo toda la ventana.
    def _desplazar_registro(muescas):
        lineas = _acumular("registro", muescas * LINEAS_POR_MUESCA)
        if lineas:
            txt_log.yview_scroll(lineas, "units")

    for widget in (frame_log, txt_log, scroll_log):
        registrar_desplazable(widget, _desplazar_registro)

    def log(mensaje):
        txt_log.insert("end", str(mensaje) + "\n")
        txt_log.see("end")
        root.update_idletasks()

    # --------------------------------------------------------
    # TIMER
    # --------------------------------------------------------

    timer = {"corriendo": False, "inicio": 0.0}

    def tick():
        if timer["corriendo"]:
            var_tiempo.set(formato_tiempo(time.time() - timer["inicio"]))
            root.after(500, tick)

    # --------------------------------------------------------
    # REVISION DE LA CARPETA
    # --------------------------------------------------------

    def revisar(*_):

        ruta = var_base.get()

        if not ruta or not Path(ruta).is_dir():
            lbl_base.config(fg=COLOR_FALTA)
            pintar_arbol([])
            btn_abrir_salida.config(state="disabled")
            return

        lbl_base.config(fg=COLOR_BASE_OK)

        try:
            _, filas = nucleo.revisar_estructura(ruta, var_aamm.get().strip())
        except Exception as error:
            # El motivo va al registro: el diagrama ya no tiene
            # columna de detalle donde ponerlo.
            log(f"Error al revisar la carpeta: {error}")
            pintar_arbol(
                [
                    {
                        "id": "error",
                        "etiqueta": "Error al revisar (ver el registro)",
                        "nivel": 0,
                        "estado": "falta",
                    }
                ]
            )
            btn_abrir_salida.config(state="disabled")
            return

        pintar_arbol(filas)

        faltan = [f for f in filas if f["estado"] == "falta"]

        if faltan:
            var_estado.set(f"Faltan {len(faltan)} entradas requeridas.")
        else:
            var_estado.set("Entradas completas.")

        btn_abrir_salida.config(state="normal")

    def faltan_subcarpetas(base):
        """Las subcarpetas del caso que todavia no estan en 'base'."""

        base = Path(base)

        return [
            sub for sub in nucleo.SUBCARPETAS_CASO
            if not (base / sub).is_dir()
        ]

    def usar_carpeta(ruta, completar=True, parent=None):
        """
        Pasa a trabajar sobre esa carpeta: la recuerda como la del
        periodo que hay escrito arriba, la guarda en config.json y
        repinta el diagrama.

        completar=True ofrece ademas crear las subcarpetas que le
        falten (es lo que ya hacia "Examinar"): una carpeta recien
        elegida no tiene por que estar completa.
        """

        ruta = str(ruta)

        var_base.set(ruta)
        guardar_config({"carpeta_base": ruta})

        aamm = var_aamm.get().strip()

        try:
            nucleo.validar_aamm(aamm)
        except nucleo.ErrorEntrada:
            aamm = ""

        if aamm:
            recordar_carpeta(aamm, ruta)
            log(f"Carpeta del periodo {aamm}: {ruta}")
        else:
            log(f"Carpeta base: {ruta}")

        revisar()

        if not completar:
            return

        # Una carpeta recien elegida (o un caso al que le falta una
        # subcarpeta) se completa aca mismo, sin obligar a armarlas a
        # mano en el explorador.
        faltan = faltan_subcarpetas(ruta)

        if faltan and messagebox.askyesno(
            "Faltan carpetas del caso",
            f"A {Path(ruta).name} le faltan {len(faltan)} carpeta(s) del "
            f"caso:\n\n"
            + "\n".join(f"  {sub}/" for sub in faltan)
            + "\n\n¿Las creo ahora? (vacias; no se toca nada de lo que "
              "ya hay)",
            parent=parent or root,
        ):
            crear_carpetas_en(ruta)

    def examinar(parent=None, titulo="Selecciona la carpeta base del caso"):
        """El dialogo de carpeta. Devuelve True si se eligio una."""

        inicial = var_base.get()

        if not inicial or not Path(inicial).is_dir():
            inicial = ""

        ruta = filedialog.askdirectory(
            title=titulo, initialdir=inicial, parent=parent or root,
        )

        if not ruta:
            return False

        usar_carpeta(ruta, parent=parent)

        return True

    def seleccionar():
        examinar()

    tk.Button(frame_carpeta, text="Examinar", command=seleccionar).pack(pady=(6, 0))

    # El ultimo AAMM que se proceso, para no volver a preguntar lo
    # mismo cada vez que el campo pierde el foco sin haber cambiado.
    # Tambien es a donde se vuelve si el usuario cancela: un periodo
    # sin carpeta propia no se queda escrito arriba.
    ultimo_aamm = {"valor": var_aamm.get().strip()}

    # Mientras se pregunta por la carpeta de un periodo, el campo del
    # AAMM pierde el foco -y <FocusOut> vuelve a llamar a
    # aamm_cambiado()-. Sin esta marca se abriria una segunda ventana
    # encima de la primera.
    preguntando = {"activo": False}

    def volver_al_periodo_anterior():
        """
        Deshace el cambio de periodo. Se usa cuando el usuario cancela
        la eleccion de carpeta: quedarse con el mes nuevo sobre la
        carpeta del mes viejo es justo lo que no puede pasar.
        """

        var_aamm.set(ultimo_aamm["valor"])
        guardar_config({"aamm": ultimo_aamm["valor"]})
        revisar()

    def ventana_carpeta_del_periodo(aamm, situacion):
        """
        El periodo que se acaba de escribir no tiene carpeta propia:
        se pregunta cual es, en vez de seguir trabajando sobre la del
        mes anterior.

        Pedido del usuario: *"dos meses no pueden tener la misma
        carpeta; si cambio el año y el mes no tiene carpeta debe
        pedirme examinar y seleccionar o crear"*.

        situacion es lo que devolvio
        nucleo.carpeta_corresponde_al_periodo():

          - CARPETA_DE_OTRO_PERIODO: la carpeta abierta es la de otro
            mes (lo dice su nombre, o que ya quedo anotada como la de
            ese otro mes). Solo se puede examinar o crear.
          - CARPETA_SIN_PERIODO: no hay con que saberlo (el nombre no
            trae ningun AAMM y nunca se anoto). Se agrega un tercer
            boton para decir "esta misma es la de este mes", que deja
            la carpeta anotada y no vuelve a preguntar.
        """

        base = var_base.get()

        top = tk.Toplevel(root)
        top.title(f"Carpeta del periodo {aamm}")
        top.transient(root)
        top.resizable(False, False)

        if situacion == nucleo.CARPETA_DE_OTRO_PERIODO:
            explicacion = (
                f"La carpeta que esta abierta no es la del periodo "
                f"{aamm}:\n\n    {base or '(ninguna)'}\n\n"
                f"Cada periodo trabaja en su propia carpeta. Elegi la "
                f"de {aamm}, o crea una nueva."
            )
        else:
            explicacion = (
                f"No se puede saber de que periodo es la carpeta que "
                f"esta abierta:\n\n    {base}\n\n"
                f"Su nombre no trae ningun AAMM. Si ESA es la carpeta "
                f"de {aamm}, decilo con el tercer boton y no se "
                f"pregunta mas; si no, elegi la de {aamm} o crea una "
                f"nueva."
            )

        tk.Label(
            top, text=explicacion, font=("Segoe UI", 9), justify="left",
            anchor="w", wraplength=560,
        ).pack(fill="x", padx=16, pady=(14, 10))

        pie = tk.Frame(top)
        pie.pack(fill="x", padx=16, pady=(0, 14))

        resuelto = {"si": False}

        def cerrar(ok):
            resuelto["si"] = ok
            top.destroy()

        def con_examinar():
            if examinar(
                parent=top,
                titulo=f"Selecciona la carpeta del periodo {aamm}",
            ):
                cerrar(True)

        def con_crear():
            # Se esconde en vez de cerrarse: hay que esperar a ver si
            # el caso se crea de verdad o se cancela.
            top.grab_release()
            top.withdraw()
            cerrar(ventana_nuevo_caso(aamm))

        def con_esta():
            usar_carpeta(base, parent=top)
            cerrar(True)

        tk.Button(
            pie, text="Examinar...", font=("Segoe UI", 9, "bold"),
            bg="#fdf0d5", command=con_examinar, width=16,
        ).pack(side="left")

        tk.Button(
            pie, text="Crear la carpeta", font=("Segoe UI", 9, "bold"),
            bg="#fdf0d5", command=con_crear, width=16,
        ).pack(side="left", padx=8)

        if situacion == nucleo.CARPETA_SIN_PERIODO and base:
            tk.Button(
                pie, text=f"Esta es la de {aamm}", command=con_esta,
                width=18,
            ).pack(side="left", padx=8)

        tk.Button(
            pie, text="Cancelar", command=lambda: cerrar(False), width=12,
        ).pack(side="right")

        top.protocol("WM_DELETE_WINDOW", lambda: cerrar(False))

        top.grab_set()
        root.wait_window(top)

        # Sin carpeta para el periodo nuevo se vuelve al anterior: la
        # ventana nunca queda con un mes apuntando a la carpeta de otro.
        if not resuelto["si"]:
            log(
                f"Periodo {aamm}: no se eligio carpeta, se vuelve a "
                f"{ultimo_aamm['valor'] or '(sin periodo)'}."
            )
            volver_al_periodo_anterior()

        return resuelto["si"]

    def aamm_cambiado(*_):

        if preguntando["activo"]:
            return

        aamm = var_aamm.get().strip()

        if aamm == ultimo_aamm["valor"]:
            return

        guardar_config({"aamm": aamm})

        try:
            nucleo.validar_aamm(aamm)
        except nucleo.ErrorEntrada:
            # Todavia no son 4 digitos (lo esta tipeando): se repinta
            # y no se pregunta nada.
            revisar()
            return

        if corriendo["activo"]:
            revisar()
            return

        anterior = ultimo_aamm["valor"]
        ultimo_aamm["valor"] = aamm

        # 1. Si este periodo ya tuvo carpeta, se vuelve a ella sola:
        #    cambiar de mes es cambiar de carpeta.
        recordada = carpeta_recordada(aamm)

        if recordada and recordada != var_base.get():
            log(f"Periodo {aamm}: se abre su carpeta, {recordada}")
            usar_carpeta(recordada, completar=False)
            return

        # 2. Si no, ¿la carpeta abierta es la de este periodo?
        situacion = nucleo.carpeta_corresponde_al_periodo(
            var_base.get(), aamm, carpetas_por_periodo(),
        )

        if situacion != nucleo.CARPETA_DEL_PERIODO:
            # Es la de otro mes, o no se sabe: se pregunta. Es el caso
            # que antes se quedaba callado sobre la carpeta del mes
            # anterior.
            ultimo_aamm["valor"] = anterior
            revisar()

            preguntando["activo"] = True
            try:
                elegida = ventana_carpeta_del_periodo(aamm, situacion)
            finally:
                preguntando["activo"] = False

            if elegida:
                ultimo_aamm["valor"] = aamm
            return

        # 3. Es la de este periodo: queda anotada (asi otro mes no la
        #    puede reusar) y solo se ofrece completarle lo que falte.
        recordar_carpeta(aamm, var_base.get())
        revisar()

        faltan = faltan_subcarpetas(var_base.get()) if var_base.get() else []

        if faltan and messagebox.askyesno(
            "Faltan carpetas del caso",
            f"A la carpeta del caso le faltan {len(faltan)} carpeta(s):\n\n"
            + "\n".join(f"  {sub}/" for sub in faltan)
            + "\n\n¿Las creo ahora?",
        ):
            crear_carpetas_en(var_base.get())

    entry_aamm.bind("<FocusOut>", aamm_cambiado)
    entry_aamm.bind("<Return>", aamm_cambiado)

    # --------------------------------------------------------
    # LANZADOR GENERICO (hilo aparte; log, barra y timer de la
    # ventana principal)
    # --------------------------------------------------------

    def caso_listo():
        """Valida lo minimo comun a todos los botones del arbol."""

        ruta = var_base.get()

        if not ruta or not Path(ruta).is_dir():
            messagebox.showwarning(
                "Falta la carpeta base",
                "Elegi primero la carpeta base del caso.",
            )
            return None

        return ruta

    def lanzar(funcion, kwargs, etiqueta_salida):
        """
        Corre funcion(**kwargs, registrar=..., progreso=...) en un
        hilo aparte. registrar/progreso escriben en el log y la barra
        de la ventana (via root.after, para no tocar tkinter desde el
        hilo). Mientras corre, todos los botones del arbol quedan
        deshabilitados.
        """

        if corriendo["activo"]:
            return

        corriendo["activo"] = True
        habilitar_botones(False)
        barra["value"] = 0
        txt_log.delete("1.0", "end")

        timer["corriendo"] = True
        timer["inicio"] = time.time()
        tick()

        resultado = {"ok": False}

        def trabajo():
            try:
                var_estado.set("Procesando...")
                funcion(
                    **kwargs,
                    registrar=lambda m: root.after(0, log, m),
                    progreso=lambda v: root.after(
                        0, barra.configure, {"value": v}
                    ),
                )
                resultado["ok"] = True

            except nucleo.ErrorEntrada as error:
                root.after(0, log, f"\nERROR DE ENTRADA:\n{error}")
                resultado["error"] = str(error)

            except Exception as error:
                root.after(
                    0, log,
                    f"\nERROR INESPERADO:\n{error}\n\n{traceback.format_exc()}",
                )
                resultado["error"] = str(error)

            finally:
                root.after(0, terminar, resultado)

        def terminar(res):
            timer["corriendo"] = False
            corriendo["activo"] = False

            # revisar() repinta el arbol entero (y con el, todos los
            # botones): no hace falta rehabilitarlos a mano.
            revisar()

            if res["ok"]:
                var_estado.set("Terminado correctamente.")
                barra["value"] = 100
                messagebox.showinfo(
                    "Listo", f"Se genero/actualizo:\n{etiqueta_salida}"
                )
            else:
                var_estado.set("Termino con errores.")
                messagebox.showerror(
                    "Error", res.get("error", "Revisa el registro.")
                )

        threading.Thread(target=trabajo, daemon=True).start()

    # --------------------------------------------------------
    # ACCIONES DE LAS FILAS
    # --------------------------------------------------------

    def traer_cmg_15min():
        """Baja el CSV 15-minutal del periodo a <CARPETA_BASE>/Cmg/."""

        ruta = caso_listo()
        if ruta is None:
            return

        aamm = var_aamm.get().strip()

        try:
            nucleo.validar_aamm(aamm)
        except nucleo.ErrorEntrada as error:
            messagebox.showwarning("Falta el periodo", str(error))
            return

        lanzar(
            nucleo.traer_csv_cmg,
            dict(carpeta_base=ruta, aamm=aamm),
            nucleo.extrae_cmg.nombre_csv_15min(aamm),
        )

    def traer_fd():
        """
        Baja el FD del periodo (SSCC_Disponibilidad_CSF_* /
        SSCC_Desempeño_*) del arbol de indicadores del DCO a
        <CARPETA_BASE>/FD y FMA/.
        """

        ruta = caso_listo()
        if ruta is None:
            return

        aamm = var_aamm.get().strip()

        try:
            nucleo.validar_aamm(aamm)
        except nucleo.ErrorEntrada as error:
            messagebox.showwarning("Falta el periodo", str(error))
            return

        lanzar(
            nucleo.traer_fd,
            dict(carpeta_base=ruta, aamm=aamm),
            f"{nucleo.CARPETA_FD_FMA}/",
        )

    def generar_fma(tipo):
        """
        Arma una de las tres salidas de FMA del periodo en
        <CARPETA_BASE>/FD y FMA/. Cada una tiene su propio boton.
        """

        ruta = caso_listo()
        if ruta is None:
            return

        aamm = var_aamm.get().strip()

        try:
            nucleo.validar_aamm(aamm)
        except nucleo.ErrorEntrada as error:
            messagebox.showwarning("Falta el periodo", str(error))
            return

        lanzar(
            nucleo.generar_fma,
            dict(carpeta_base=ruta, aamm=aamm, tipos={tipo}),
            f"fma_{tipo}_{aamm}",
        )

    def traer_subastas():
        """
        Copia los Access de subastas del periodo
        (OfertasSSCCAdj*.accdb) de la unidad de red a
        <CARPETA_BASE>/Subastas/DB subastas/.
        """

        ruta = caso_listo()
        if ruta is None:
            return

        aamm = var_aamm.get().strip()

        try:
            nucleo.validar_aamm(aamm)
        except nucleo.ErrorEntrada as error:
            messagebox.showwarning("Falta el periodo", str(error))
            return

        lanzar(
            nucleo.traer_subastas,
            dict(carpeta_base=ruta, aamm=aamm),
            f"{nucleo.CARPETA_DB_SUBASTAS}/",
        )

    def generar_cmg():
        """Arma cmg.xlsx con el CSV que ya esta en Cmg/."""

        ruta = caso_listo()
        if ruta is None:
            return

        aamm = var_aamm.get().strip()

        try:
            nucleo.validar_aamm(aamm)
        except nucleo.ErrorEntrada as error:
            messagebox.showwarning("Falta el periodo", str(error))
            return

        lanzar(
            nucleo.generar_cmg,
            dict(carpeta_base=ruta, aamm=aamm),
            nucleo.ARCHIVO_CMG,
        )

    def actualizar_medidas_sae():
        """
        Corre los cuatro pasos de Medidas_SAE.xlsx de un viaje
        (homologacion -> descarga -> claves -> API de operacion real).
        Es lo mas lento del programa: baja el mes completo de las dos
        APIs del Coordinador. Si se corta, la corrida siguiente retoma
        donde quedo.
        """

        ruta = caso_listo()
        if ruta is None:
            return

        aamm = var_aamm.get().strip()

        try:
            nucleo.validar_aamm(aamm)
        except nucleo.ErrorEntrada as error:
            messagebox.showwarning("Falta el periodo", str(error))
            return

        if not messagebox.askyesno(
            f"Generar {nucleo.ARCHIVO_MEDIDAS_SAE}",
            f"Se van a bajar las medidas del periodo {aamm} de las dos "
            f"APIs del Coordinador, punto de medida por punto de "
            f"medida.\n\nEs el proceso mas lento del programa (puede "
            f"tardar bastante). Si se corta, la proxima vez retoma "
            f"donde quedo.\n\n¿Seguir?",
        ):
            return

        lanzar(
            nucleo.generar_medidas_sae,
            dict(carpeta_base=ruta, aamm=aamm),
            nucleo.ARCHIVO_MEDIDAS_SAE,
        )

    def actualizar_consolidado(secciones):
        """
        secciones: set con UN id de SECCIONES_CONSOLIDADO (el boton
        de esa hoja) o None para todas. Lo que no entra se conserva
        tal cual esta hoy en el archivo -- las hojas de calculo
        incluidas: las dos mitades comparten planilla.
        """

        ruta = caso_listo()
        if ruta is None:
            return

        if secciones is None:
            secciones = {s[0] for s in nucleo.SECCIONES_CONSOLIDADO}

        lanzar(
            nucleo.generar_consolidado,
            dict(
                carpeta_base=ruta,
                aamm=var_aamm.get().strip(),
                secciones_activas=secciones,
            ),
            nucleo.ARCHIVO_SALIDA,
        )

    def actualizar_pagos(secciones):
        """Idem para las hojas de calculo / SECCIONES_PAGOS."""

        ruta = caso_listo()
        if ruta is None:
            return

        if secciones is None:
            secciones = {s[0] for s in nucleo.SECCIONES_PAGOS}

        lanzar(
            nucleo.generar_pagos_bess,
            dict(carpeta_base=ruta, secciones_activas=secciones,
                 aamm=var_aamm.get().strip()),
            nucleo.ARCHIVO_SALIDA,
        )

    def actualizar_planilla_entera():
        """
        El boton de la fila del archivo: las dos mitades de la
        planilla de una sola vez, en el orden en que se necesitan
        -primero las entradas, despues el calculo que las lee-, en
        una sola corrida con un solo registro.

        No es lo mismo que "Ejecutar todo": aca no se trae ni se
        genera ninguna ENTRADA (ni medidas, ni CMg, ni FD, ni
        subastas), solo se rehacen las hojas de la planilla con lo que
        ya hay en la carpeta del caso.
        """

        ruta = caso_listo()
        if ruta is None:
            return

        aamm = var_aamm.get().strip()

        def correr(registrar, progreso):

            nucleo.generar_consolidado(
                carpeta_base=ruta,
                aamm=aamm,
                secciones_activas={s[0] for s in nucleo.SECCIONES_CONSOLIDADO},
                registrar=registrar,
                progreso=lambda v: progreso(v / 2),
            )
            nucleo.generar_pagos_bess(
                carpeta_base=ruta,
                aamm=aamm,
                secciones_activas={s[0] for s in nucleo.SECCIONES_PAGOS},
                registrar=registrar,
                progreso=lambda v: progreso(50 + v / 2),
            )

        lanzar(correr, {}, nucleo.ARCHIVO_SALIDA)

    # --------------------------------------------------------
    # CASO NUEVO: CREAR LA CARPETA DEL PERIODO CON SUS SUBCARPETAS
    # --------------------------------------------------------

    def crear_carpetas_en(base, usar=True):
        """
        Crea (o completa) la estructura del caso en 'base'. Si
        usar=True, esa carpeta pasa a ser la carpeta base de la
        ventana.
        """

        try:
            creadas = nucleo.crear_estructura_caso(base, registrar=log)
        except nucleo.ErrorEntrada as error:
            messagebox.showerror("No se pudo crear", str(error))
            return False

        if creadas:
            log(f"Caso preparado en {base} ({len(creadas)} carpeta(s) nuevas).")
        else:
            log(f"{base} ya tenia todas sus carpetas.")

        if usar:
            # usar_carpeta() ademas la deja ANOTADA como la carpeta de
            # este periodo: sin eso, el mes siguiente podria volver a
            # caer en esta misma.
            usar_carpeta(base, completar=False)
        else:
            revisar()

        return True

    def ventana_nuevo_caso(aamm_objetivo=None):
        """
        Crea la carpeta de un periodo que todavia no existe, con todas
        sus subcarpetas adentro (pedido del usuario: "cuando yo elija
        un mes que no exista, me permita elegir y crear la carpeta con
        las carpetas dentro").

        No impone ninguna convencion de nombre: propone el nombre de
        la carpeta que se estaba usando con el AAMM cambiado ("Balance
        BESS 2607" -> "Balance BESS 2608") y deja cambiar tanto el
        nombre como donde crearla.
        """

        if corriendo["activo"]:
            return False

        aamm = (aamm_objetivo or var_aamm.get()).strip()

        base_actual = Path(var_base.get()) if var_base.get() else None

        # Por defecto, al lado de la carpeta del periodo anterior.
        if base_actual and base_actual.parent.is_dir():
            padre_inicial = str(base_actual.parent)
        else:
            padre_inicial = str(Path.home())

        try:
            nombre_inicial = nucleo.nombre_caso_sugerido(aamm, base_actual)
        except nucleo.ErrorEntrada:
            nombre_inicial = base_actual.name if base_actual else "Caso nuevo"

        top = tk.Toplevel(root)
        top.title("Crear la carpeta del caso")
        top.geometry("760x460")
        top.transient(root)

        # Quien la abre espera el resultado (ver el final de la
        # funcion): si se cancela, el periodo nuevo se queda sin
        # carpeta y quien llamo tiene que hacer algo al respecto.
        resultado = {"ok": False}

        var_padre = tk.StringVar(value=padre_inicial)
        var_nombre = tk.StringVar(value=nombre_inicial)

        tk.Label(
            top,
            text=(
                f"Se va a crear la carpeta del periodo {aamm or '(sin AAMM)'} "
                f"con todas sus subcarpetas vacias adentro.\n"
                f"No se copia ni se mueve ningun archivo."
            ),
            font=("Segoe UI", 9), anchor="w", justify="left",
        ).pack(fill="x", padx=14, pady=(12, 8))

        frame_padre = tk.LabelFrame(
            top, text="Donde crearla", padx=10, pady=8
        )
        frame_padre.pack(fill="x", padx=14, pady=4)

        tk.Entry(
            frame_padre, textvariable=var_padre, font=("Segoe UI", 9),
        ).pack(side="left", fill="x", expand=True)

        def elegir_padre():
            elegida = filedialog.askdirectory(
                title="Carpeta donde crear el caso",
                initialdir=var_padre.get() or str(Path.home()),
                parent=top,
            )
            if elegida:
                var_padre.set(elegida)

        tk.Button(frame_padre, text="Examinar", command=elegir_padre).pack(
            side="left", padx=(8, 0)
        )

        frame_nombre = tk.LabelFrame(
            top, text="Nombre de la carpeta", padx=10, pady=8
        )
        frame_nombre.pack(fill="x", padx=14, pady=4)

        tk.Entry(
            frame_nombre, textvariable=var_nombre, font=("Segoe UI", 9),
        ).pack(fill="x")

        frame_lista = tk.LabelFrame(
            top, text="Subcarpetas que se crean", padx=10, pady=8
        )
        frame_lista.pack(fill="both", expand=True, padx=14, pady=4)

        tk.Label(
            frame_lista,
            text="\n".join(f"{sub}/" for sub in nucleo.SUBCARPETAS_CASO),
            font=("Consolas", 9), anchor="w", justify="left", fg=COLOR_NEUTRO,
        ).pack(fill="both", expand=True)

        frame_pie = tk.Frame(top)
        frame_pie.pack(fill="x", padx=14, pady=(4, 12))

        def crear():

            padre = var_padre.get().strip()
            nombre = var_nombre.get().strip()

            if not padre or not Path(padre).is_dir():
                messagebox.showwarning(
                    "Falta donde crearla",
                    "Elegi una carpeta que exista para crear el caso adentro.",
                    parent=top,
                )
                return

            if not nombre:
                messagebox.showwarning(
                    "Falta el nombre",
                    "Escribi como se va a llamar la carpeta del caso.",
                    parent=top,
                )
                return

            destino = Path(padre) / nombre

            if destino.is_dir() and not messagebox.askyesno(
                "Ya existe",
                f"{destino} ya existe.\n\nSe le van a crear las "
                f"subcarpetas que le falten (no se toca nada de lo que "
                f"ya tiene). ¿Seguir?",
                parent=top,
            ):
                return

            # El periodo se fija ANTES de crear: crear_carpetas_en()
            # anota la carpeta como la de "el periodo que hay escrito
            # arriba", y ese tiene que ser ya el nuevo.
            if aamm:
                var_aamm.set(aamm)
                ultimo_aamm["valor"] = aamm
                guardar_config({"aamm": aamm})

            if crear_carpetas_en(destino):
                resultado["ok"] = True
                top.destroy()

        tk.Button(
            frame_pie, text="Crear y usar", font=("Segoe UI", 9, "bold"),
            bg="#fdf0d5", command=crear,
        ).pack(side="right", padx=8)

        tk.Button(frame_pie, text="Cancelar", command=top.destroy).pack(
            side="right", padx=8
        )

        top.grab_set()
        root.wait_window(top)

        return resultado["ok"]

    # --------------------------------------------------------
    # EJECUTAR TODO (una ventana con el plan de la corrida)
    # --------------------------------------------------------

    def ventana_ejecutar_todo():
        """
        Abre la ventana del plan: una fila por tarea, tildada si hay
        que hacerla, con lo que falta para poder ejecutar. El boton
        "Ejecutar" queda bloqueado mientras falte una entrada o una
        dependencia (pedido del usuario).

        El plan lo arma nucleo.planificar(): esta ventana solo lo
        dibuja y devuelve la seleccion.
        """

        ruta = caso_listo()
        if ruta is None:
            return

        if corriendo["activo"]:
            return

        aamm = var_aamm.get().strip()

        # None = "lo que falte" (el arranque normal); despues pasa a
        # ser el set que el usuario va tildando.
        seleccion = {"ids": None}

        top = tk.Toplevel(root)
        top.title("Ejecutar todo")
        top.geometry("900x640")
        top.transient(root)

        tk.Label(
            top,
            text=(
                "Se hace solo lo que falta. Tilda tambien lo que quieras "
                "rehacer: lo que sale de ahi se tilda solo."
            ),
            font=("Segoe UI", 9), anchor="w", justify="left",
        ).pack(fill="x", padx=14, pady=(12, 4))

        frame_pasos = tk.LabelFrame(top, text="Pasos", padx=8, pady=6)
        frame_pasos.pack(fill="both", expand=True, padx=14, pady=4)

        lienzo = tk.Canvas(frame_pasos, borderwidth=0, highlightthickness=0)
        barra_pasos = tk.Scrollbar(
            frame_pasos, orient="vertical", command=lienzo.yview
        )
        lienzo.configure(yscrollcommand=barra_pasos.set)
        barra_pasos.pack(side="right", fill="y")
        lienzo.pack(side="left", fill="both", expand=True)

        lista = tk.Frame(lienzo)
        ventana_lista = lienzo.create_window((0, 0), window=lista, anchor="nw")

        def ajustar_lista(event=None):
            lienzo.configure(scrollregion=lienzo.bbox("all"))
            lienzo.itemconfig(ventana_lista, width=lienzo.winfo_width())

        lista.bind("<Configure>", ajustar_lista)
        lienzo.bind("<Configure>", ajustar_lista)

        frame_avisos = tk.LabelFrame(
            top, text="Para poder ejecutar", padx=8, pady=6
        )
        frame_avisos.pack(fill="x", padx=14, pady=4)

        txt_avisos = tk.Text(
            frame_avisos, height=6, font=("Segoe UI", 8), wrap="word"
        )
        txt_avisos.pack(fill="x")

        frame_pie = tk.Frame(top)
        frame_pie.pack(fill="x", padx=14, pady=(4, 12))

        btn_ejecutar = tk.Button(
            frame_pie, text="Ejecutar", font=("Segoe UI", 9, "bold"),
            bg="#fdf0d5",
        )

        estado_plan = {"plan": None}

        def repintar():

            plan = nucleo.planificar(ruta, aamm, seleccion["ids"])
            estado_plan["plan"] = plan

            for hijo in lista.winfo_children():
                hijo.destroy()

            for tarea in plan["tareas"]:

                fila = tk.Frame(lista)
                fila.pack(fill="x", pady=1)

                var = tk.BooleanVar(value=tarea["seleccionada"])

                def alternar(id_tarea=tarea["id"], var=var):
                    actual = {
                        t["id"] for t in estado_plan["plan"]["tareas"]
                        if t["seleccionada"]
                    }
                    if var.get():
                        # Tildar algo arrastra lo que sale de ahi: si
                        # no, queda mezclado con la corrida anterior.
                        actual |= nucleo.propagar_seleccion({id_tarea})
                    else:
                        actual.discard(id_tarea)
                    seleccion["ids"] = actual
                    repintar()

                tk.Checkbutton(
                    fila, variable=var, command=alternar,
                    state="disabled" if corriendo["activo"] else "normal",
                ).pack(side="left")

                tk.Label(
                    fila, text=tarea["etiqueta"], width=44, anchor="w",
                    font=("Consolas", 9),
                ).pack(side="left")

                color = {
                    "al dia": COLOR_OK,
                    "se genera": COLOR_BASE_OK,
                    "se rehace": COLOR_PENDIENTE,
                    "bloqueada": COLOR_FALTA,
                    "sin tildar": COLOR_NEUTRO,
                }[tarea["estado"]]

                tk.Label(
                    fila, text=tarea["estado"], width=12, anchor="w",
                    fg=color, font=("Segoe UI", 9, "bold"),
                ).pack(side="left")

                motivo = ", ".join(
                    tarea["faltan_requisitos"] + tarea["faltan_dependencias"]
                )

                tk.Label(
                    fila,
                    text=(f"falta: {motivo}" if motivo else ""),
                    anchor="w", fg=COLOR_NEUTRO, font=("Segoe UI", 8),
                ).pack(side="left", fill="x", expand=True)

            txt_avisos.delete("1.0", "end")

            if plan["bloqueos"]:
                txt_avisos.insert(
                    "end", "\n".join(f"- {b}" for b in plan["bloqueos"])
                )
            else:
                pasos = " -> ".join(
                    nucleo.GRUPO_POR_ID[g].etiqueta for g in plan["orden"]
                )
                txt_avisos.insert(
                    "end", f"Todo listo. Orden de la corrida: {pasos}\n"
                )

            for aviso in plan["advertencias"]:
                txt_avisos.insert("end", f"\n[AVISO] {aviso}")

            btn_ejecutar.config(
                state="normal" if plan["puede_ejecutar"] else "disabled"
            )

            ajustar_lista()

        def solo_lo_que_falta():
            seleccion["ids"] = None
            repintar()

        def tildar_todo():
            seleccion["ids"] = {t.id for t in nucleo.TAREAS}
            repintar()

        def ejecutar():
            plan = estado_plan["plan"]
            if plan is None or not plan["puede_ejecutar"]:
                return
            ids = {
                t["id"] for t in plan["tareas"] if t["seleccionada"]
            }
            top.destroy()
            lanzar(
                nucleo.ejecutar_plan,
                dict(carpeta_base=ruta, aamm=aamm, seleccion=ids),
                nucleo.ARCHIVO_SALIDA,
            )

        btn_ejecutar.config(command=ejecutar)

        tk.Button(
            frame_pie, text="Solo lo que falta", command=solo_lo_que_falta,
        ).pack(side="left", padx=(0, 8))

        tk.Button(
            frame_pie, text="Rehacer todo", command=tildar_todo,
        ).pack(side="left", padx=8)

        tk.Button(frame_pie, text="Cerrar", command=top.destroy).pack(
            side="right", padx=8
        )

        btn_ejecutar.pack(side="right", padx=8)

        repintar()

    # --------------------------------------------------------
    # BOTONES FIJOS ABAJO
    # --------------------------------------------------------

    def abrir_salida():
        abrir_en_explorador(var_base.get())

    tk.Button(
        frame_botones,
        text="Crear carpeta del caso",
        command=lambda: ventana_nuevo_caso(var_aamm.get().strip()),
    ).pack(side="left", padx=10, expand=True)

    tk.Button(
        frame_botones,
        text="Ejecutar todo",
        font=("Segoe UI", 9, "bold"),
        bg="#fdf0d5",
        command=ventana_ejecutar_todo,
    ).pack(side="left", padx=10, expand=True)

    btn_abrir_salida = tk.Button(
        frame_botones,
        text="Abrir carpeta del caso",
        state="disabled",
        command=abrir_salida,
    )
    btn_abrir_salida.pack(side="left", padx=10, expand=True)

    tk.Button(frame_botones, text="Salir", command=root.destroy).pack(
        side="left", padx=10, expand=True
    )

    # --------------------------------------------------------
    # ARRANQUE
    # --------------------------------------------------------

    pintar_arbol([])
    log(
        "Selecciona la carpeta base del caso e ingresa el periodo "
        "(AAMM). Cada fila del diagrama de abajo trae su propio boton: "
        "'Traer' y 'Generar' en Cmg/, 'Traer' y un 'Generar' por cada "
        "FMA en 'FD y FMA/', 'Traer' en Subastas/DB subastas/, y en "
        f"{nucleo.ARCHIVO_SALIDA} un boton por hoja ('Actualizar' en "
        "las de entrada, 'Calcular' / 'Resumir' / 'Traer' / 'Asignar' "
        "en las de calculo). Cada archivo y cada carpeta del diagrama "
        "es un link a su ruta, y las filas que traen algo de afuera "
        "dicen 'Origen: ...' con un link a la carpeta de origen."
    )

    if var_aamm.get():
        log(f"Periodo recordado: {var_aamm.get()}")

    if var_base.get():
        log(f"Carpeta recordada: {var_base.get()}")
        revisar()

    # El par (carpeta, periodo) con el que arranca la ventana es el que
    # quedo de la ultima vez: se anota como la carpeta de ESE periodo.
    # Sin esto, volver a un mes viejo pediria de nuevo su carpeta, que
    # es justo lo que se acaba de elegir.
    if var_base.get() and var_aamm.get():
        try:
            nucleo.validar_aamm(var_aamm.get())
        except nucleo.ErrorEntrada:
            pass
        else:
            if nucleo.carpeta_corresponde_al_periodo(
                var_base.get(), var_aamm.get(), carpetas_por_periodo(),
            ) != nucleo.CARPETA_DE_OTRO_PERIODO:
                recordar_carpeta(var_aamm.get(), var_base.get())

    root.mainloop()


if __name__ == "__main__":
    main()
