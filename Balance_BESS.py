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
            (fma_cft_<AAMM>.xlsx: oculto, FMA_CON_CTF = False)
        Subastas/
            DB subastas/                   [Traer]
        11_PAGOS_BESS_AAMM.xlsx            [Actualizar]
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
        SAE_AAMM.csv                   [Generar]
        Log/
            Control_corrida.xlsx
            Registro_ultima_corrida.txt

La salida es UNA sola planilla (antes eran dos,
Consolidado_entradas.xlsx y Pagos_BESS.xlsx) y sus hojas quedan
ordenadas de FIN A INICIO: arriba el Resumen -lo primero que se mira-
y abajo las entradas de las que sale todo. Las hojas de control de la
corrida (Alertas, Ejecucion, Log) ya no estan ahi adentro: viven en
Log/Control_corrida.xlsx. La planilla se llama 11_PAGOS_BESS_AAMM.xlsx
(con el periodo) y el ultimo boton deja SAE_AAMM.csv, el CSV de pagos.

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
de __config__/tks.json, y lo que recuerda la ventana de
__config__/config.json ("modo_bess"); __config__ es la carpeta hermana
de la de este programa (ver Script/config.py y config.ejemplo.json).
"""

import os
import queue
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

from Script import config
from Script.arbol import _prefijos_arbol
from Script.dialogos import mostrar_error

# Script.nucleo (pandas, openpyxl y ~60 modulos: medio segundo o mas, en
# Windows varias veces eso) NO se importa aca arriba: se carga en main()
# con la ventana ya a la vista y respondiendo (cartel "Cargando..."). El
# resto del archivo escribe nucleo.algo dentro de funciones que corren
# despues de cargarlo. Ver PLAN_OPTIMIZACION_VENTANAS_BESS.md, Fase 2.
nucleo = None


def _cargar_nucleo():
    """Importa Script.nucleo (una vez) y lo deja en el nombre `nucleo`."""

    global nucleo
    if nucleo is None:
        from Script import nucleo as _nucleo
        nucleo = _nucleo
    return nucleo


def __getattr__(nombre):
    """Lo que antes se copiaba de nucleo al importar (PEP 562)."""

    if nombre == "ARCHIVO_REGISTRO":
        return _cargar_nucleo().ARCHIVO_REGISTRO
    raise AttributeError(f"module {__name__!r} has no attribute {nombre!r}")


# El hilo de verdad, aunque una prueba parchee threading.Thread para
# correr los procesos en el mismo hilo.
_HiloDeCarga = threading.Thread

TEXTO_CARGANDO = "Cargando\u2026"


def _ventana_viva(root):
    """False si la ventana ya se cerro (un update() despues no siempre falla)."""

    try:
        return bool(root.winfo_exists())
    except tk.TclError:
        return False


def _importar_con_la_ventana_viva(root):
    """
    Carga nucleo en un hilo mientras la ventana sigue respondiendo (se
    puede mover, no queda "No responde"). Devuelve False si la cerraron
    mientras cargaba o si la carga fallo (en ese caso ya se mostro el
    error y se cerro la ventana).
    """

    if nucleo is not None:
        return True

    resultado = {}

    def cargar():
        try:
            _cargar_nucleo()
        except BaseException as error:  # se muestra en el hilo principal
            resultado["error"] = error

    hilo = _HiloDeCarga(target=cargar, name="cargar-nucleo", daemon=True)
    hilo.start()

    while hilo.is_alive():
        try:
            root.update()
        except tk.TclError:
            return False  # cerraron la ventana mientras cargaba
        if not _ventana_viva(root):
            return False
        time.sleep(0.01)

    if not _ventana_viva(root):
        return False

    if "error" in resultado:
        error = resultado["error"]
        try:
            mostrar_error(
                "No se pudo cargar el programa",
                "".join(traceback.format_exception(error)),
                parent=root,
            )
            root.destroy()
        except tk.TclError:
            pass
        return False

    return True


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

# El simbolo que la fila de Medidas_SAE.xlsx muestra cuando el archivo
# se genero con informacion incompleta (ver nucleo.aviso_puntos_fallidos).
# Va sin el selector de emoji (U+FE0F): Tk en Windows lo dibuja como un
# cuadrito.
SIMBOLO_AVISO = "\u26a0"
COLOR_AVISO = "#d97706"

# Colores del Registro (pedido del usuario): verde lo que termino bien,
# amarillo los avisos, rojo los errores. El amarillo es oscuro (ambar)
# porque un amarillo puro no se lee sobre el fondo blanco.
COLOR_REGISTRO = {
    "ok": "#1a7f1a",
    "aviso": "#b8860b",
    "error": "#c00020",
}

# Donde queda el registro de la ultima corrida, en la carpeta Log/ del
# caso (para poder mandarlo sin copiar a mano): nucleo.ARCHIVO_REGISTRO,
# que tambien se lee como Balance_BESS.ARCHIVO_REGISTRO (ver __getattr__).

_PATRON_ERROR = re.compile(
    r"\bERROR\b|Traceback|\bNO se corre\b|: error\b|termino con problemas",
    re.IGNORECASE,
)
_PATRON_AVISO = re.compile(
    r"\[AVISO\]|\bAVISO\b|\bOJO\b|advertencia|desactualizad|omitido|"
    r"queda fuera|no se encontro|sin datos|incomplet|cambio de hora",
    re.IGNORECASE,
)
_PATRON_OK = re.compile(
    r"^(\[[^\]]+\]\s*)?(Listo|Terminado|Hecho|OK\b)|: ok\b|correctamente",
    re.IGNORECASE,
)


def nivel_del_mensaje(texto):
    """'error' / 'aviso' / 'ok' / None: que color lleva en el Registro."""

    texto = str(texto).strip()
    if not texto:
        return None
    if _PATRON_ERROR.search(texto):
        return "error"
    if _PATRON_AVISO.search(texto):
        return "aviso"
    if _PATRON_OK.search(texto):
        return "ok"
    return None


def formatear_reloj(segundos):
    """83.4 -> '01:23'; 3725 -> '1:02:05'."""

    entero = int(max(0, segundos))
    horas, resto = divmod(entero, 3600)
    minutos, seg = divmod(resto, 60)
    if horas:
        return f"{horas}:{minutos:02d}:{seg:02d}"
    return f"{minutos:02d}:{seg:02d}"

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

# Cuantos mensajes del hilo de trabajo se escriben en el registro por
# pasada (ver vaciar_cola_log): una rafaga grande no congela la ventana.
MENSAJES_POR_TANDA = 400


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
    """
    Lo que dejo guardado ESTE PC/usuario (carpeta base, AAMM), en
    __config__/config.json -> "modo_bess" -> <equipo>_<usuario> (ver
    Script/config.py). La primera vez se copia del config.json de antes.
    """

    return config.seccion_usuario(
        config.MODO_BESS, get_usuario(), clave_antigua=get_usuario()
    )


def guardar_config(data):
    """
    Guarda en el bloque de este PC/usuario dentro de "modo_bess", sin
    tocar el resto del archivo (es compartido con los otros programas).
    """

    config.actualizar_seccion_usuario(config.MODO_BESS, get_usuario(), data)


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


def recordar_carpeta(aamm, ruta, ademas=None):
    """
    Deja anotado que 'ruta' es la carpeta del periodo 'aamm'. Si esa
    carpeta estaba anotada como la de OTRO periodo, se le saca: una
    carpeta pertenece a un solo periodo.

    Si ya estaba anotada asi, no se escribe nada (config.json es
    compartido y puede estar en red). 'ademas' son otras claves a
    guardar en la MISMA escritura (usar_carpeta guarda tambien la
    carpeta base).
    """

    datos = dict(ademas or {})

    if aamm and ruta:
        ruta = str(ruta)
        actuales = carpetas_por_periodo()
        carpetas = {
            periodo: otra
            for periodo, otra in actuales.items()
            if periodo == str(aamm) or otra != ruta
        }
        carpetas[str(aamm)] = ruta
        if carpetas != actuales:
            datos[CLAVE_CARPETAS_PERIODO] = carpetas

    if datos:
        guardar_config(datos)


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


def abrir_archivo(ruta):
    """
    Abre ESE archivo con su programa (Excel para un .xlsx). Es la
    excepcion a "nunca se abre el archivo": solo se usa cuando el
    usuario lo pide con un boton (el "Ver detalle" del aviso de
    Medidas_SAE.xlsx). Devuelve False si el archivo no esta.
    """

    ruta = Path(ruta)

    if not ruta.is_file():
        return False

    if sys.platform == "win32":
        os.startfile(str(ruta))
    elif sys.platform == "darwin":
        subprocess.Popen(["open", str(ruta)])
    else:
        subprocess.Popen(["xdg-open", str(ruta)])

    return True


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
# (├── / └── / │): nucleo.py sabe de estructura, no de dibujo. Los
# prefijos viven en Script/arbol.py (los comparte Balance_CRA.py).
# ============================================================

# ============================================================
# VENTANA
# ============================================================

def main():

    # La ventana primero, con un cartel: lo pesado (nucleo) se importa
    # con ella ya a la vista.
    root = tk.Tk()
    root.title("Balance BESS / SSCC")
    root.geometry("1180x820")
    cartel = tk.Label(
        root, text=TEXTO_CARGANDO, font=("Segoe UI", 11), fg=COLOR_NEUTRO,
    )
    cartel.pack(expand=True)
    try:
        root.update()                      # la ventana ya se ve
    except tk.TclError:
        return
    if not _ventana_viva(root):
        return

    if not _importar_con_la_ventana_viva(root):
        return                             # la cerraron mientras cargaba
    cartel.destroy()

    # Los tokens de las APIs pasan solos a __config__/tks.json la primera
    # vez (antes estaban en "claves_api" del config.json local). Va
    # despues de mostrar la ventana: __config__ puede estar en red.
    config.migrar_tokens()
    cfg = leer_config()

    var_base = tk.StringVar(value=cfg.get("carpeta_base", ""))
    var_aamm = tk.StringVar(value=cfg.get("aamm", ""))
    var_estado = tk.StringVar(value="Listo")
    var_tiempo = tk.StringVar(value="00:00:00")

    # Los botones viven DENTRO de las filas del arbol: se guardan por id
    # de fila para poder deshabilitarlos mientras corre algo, y se
    # renuevan cada vez que el arbol se rehace entero.
    botones_arbol = {}
    corriendo = {"activo": False}
    # Las filas ya dibujadas (por id) y lo ultimo que se pinto: con eso
    # pintar_arbol cambia solo lo que cambio (ver pintar_arbol).
    filas_vivas = {}
    arbol_pintado = {"forma": None, "contenido": {}}

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

    # Dos ataduras separadas: el canvas que cambia de tamaño solo ajusta
    # el ancho de lo de adentro (y solo si cambio), y lo de adentro que
    # cambia solo ajusta la region desplazable. Antes las dos hacian las
    # dos cosas y una disparaba a la otra.
    ancho_pintado = {}

    def ajustar_ancho(lienzo, ventana_interna, event):
        if ancho_pintado.get(str(lienzo)) != event.width:
            ancho_pintado[str(lienzo)] = event.width
            lienzo.itemconfig(ventana_interna, width=event.width)

    def ajustar(event=None):
        canvas.configure(scrollregion=canvas.bbox("all"))

    contenedor.bind("<Configure>", ajustar)
    canvas.bind(
        "<Configure>", lambda e: ajustar_ancho(canvas, ventana_canvas, e)
    )

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

        resto = resto_rueda.setdefault(clave, 0.0)
        # Al cambiar de sentido el resto pendiente ya no sirve: seguirlo
        # sumando frenaria medio pixel el primer evento del otro lado.
        if resto and (resto > 0) != (cantidad > 0):
            resto = 0.0
        total = cantidad + resto
        entero = int(total)
        resto_rueda[clave] = total - entero
        return entero

    # Un desplazamiento por cuadro, no por evento: un touchpad de
    # precision manda decenas de eventos por segundo y cada yview_scroll
    # redibuja el arbol. Los pixeles se suman y se aplican de una vez
    # cuando Tk queda libre (after_idle: en el mismo update()).
    desplazamiento_pendiente = {}

    def _desplazar_por_cuadro(lienzo, pixeles):
        clave = str(lienzo)
        if clave in desplazamiento_pendiente:
            desplazamiento_pendiente[clave] += pixeles
            return
        desplazamiento_pendiente[clave] = pixeles

        def aplicar():
            total = desplazamiento_pendiente.pop(clave, 0)
            if total:
                try:
                    lienzo.yview_scroll(total, "units")
                except tk.TclError:
                    pass  # se cerro la ventana del lienzo

        root.after_idle(aplicar)

    def _desplazar_ventana(muescas):
        # scrollregion mas chica que lo visible = no hay nada que
        # desplazar; sin esto la ventana "tiembla" con la rueda.
        region = canvas.bbox("all")
        if not region or region[3] - region[1] <= canvas.winfo_height():
            return
        pixeles = _acumular("ventana", muescas * PIXELES_POR_MUESCA)
        if pixeles:
            _desplazar_por_cuadro(canvas, pixeles)

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
        inicio = widget
        while widget is not None:
            desplazar = desplazables.get(str(widget))
            if desplazar is not None:
                desplazar(muescas)
                return "break"
            widget = getattr(widget, "master", None)
        # Sobre otra ventana (un dialogo como "Ejecutar todo") la rueda
        # no mueve la principal que esta detras (pedido del usuario: "al
        # mover la ruedita se mueve la ventana principal").
        try:
            otra_ventana = (
                inicio is not None and inicio.winfo_toplevel() is not root
            )
        except (tk.TclError, AttributeError):
            otra_ventana = False
        if otra_ventana:
            return "break"
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
    frame_carpeta.pack(fill="x", padx=14, pady=(10, 4))

    # Diseño de esta parte tomado del Revisor_Relq_15min (repo abby),
    # pedido del usuario: ruta alineada a la izquierda y botones debajo,
    # marcos mas juntos.
    lbl_base = tk.Label(
        frame_carpeta,
        textvariable=var_base,
        wraplength=980,
        justify="left",
        anchor="w",
        cursor="hand2",
        font=("Segoe UI", 9),
    )
    lbl_base.pack(anchor="w")

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
        contenedor, text="Periodo del caso (AAMM)", padx=10, pady=6
    )
    frame_periodo.pack(fill="x", padx=14, pady=4)

    fila_aamm = tk.Frame(frame_periodo)
    fila_aamm.pack(fill="x")

    validador_aamm = (
        root.register(lambda s: s == "" or (s.isdigit() and len(s) <= 4)),
        "%P",
    )

    entry_aamm = tk.Entry(
        fila_aamm,
        textvariable=var_aamm,
        width=8,
        justify="center",
        font=("Consolas", 11, "bold"),
        validate="key",
        validatecommand=validador_aamm,
    )
    entry_aamm.pack(side="left")

    # Como en el Revisor_Relq_15min: un boton explicito para aplicar el
    # mes (ademas de Enter y de salir del campo).
    boton_ir_mes = tk.Button(
        fila_aamm, text="Ir a este mes", bg="#1f5fa8", fg="white",
        font=("Segoe UI", 9, "bold"),
    )
    boton_ir_mes.pack(side="left", padx=8)

    tk.Label(
        fila_aamm,
        text=(
            "4 digitos: año+mes simplificado. Ej. 2607 para julio de "
            "2026. Con el se buscan el archivo de SoC dentro de "
            "Medidas/ y el CSV 15-minutal dentro de Cmg/."
        ),
        fg="#444444",
        font=("Segoe UI", 8),
        wraplength=760,
        justify="left",
        anchor="w",
    ).pack(side="left", padx=6)

    # Cambio de hora del mes: se detecta solo con el AAMM (ver
    # nucleo.cambio_hora) y queda a la vista mientras se trabaja.
    var_cambio_hora = tk.StringVar(value="")
    label_cambio_hora = tk.Label(
        frame_periodo,
        textvariable=var_cambio_hora,
        font=("Segoe UI", 9, "bold"),
        fg=COLOR_NEUTRO,
        wraplength=900,
        justify="left",
        anchor="w",
    )
    label_cambio_hora.pack(fill="x", pady=(6, 0))

    def mostrar_cambio_hora(*_):
        aamm = var_aamm.get().strip()
        try:
            nucleo.validar_aamm(aamm)
            texto = nucleo.describir_periodo(aamm)
            hay = bool(nucleo.cambios_del_periodo(aamm))
        except (nucleo.ErrorEntrada, ValueError):
            texto, hay = "", False
        var_cambio_hora.set(texto)
        label_cambio_hora.config(fg=COLOR_AVISO if hay else COLOR_NEUTRO)
        return texto, hay

    var_aamm.trace_add("write", mostrar_cambio_hora)
    mostrar_cambio_hora()

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

    def presentar(top, modal=False, ancho=None, alto=None):
        """
        Muestra un Toplevel ya construido (se crea con withdraw()),
        centrado sobre la ventana principal: asi aparece de una vez, en
        vez de verse armarse. Una ventana oculta todavia no tiene tamaño
        real: se usa el pedido, o el que se le fijo (ancho, alto).
        """

        top.update_idletasks()
        ancho = ancho or top.winfo_reqwidth()
        alto = alto or top.winfo_reqheight()
        x = root.winfo_rootx() + (root.winfo_width() - ancho) // 2
        y = root.winfo_rooty() + (root.winfo_height() - alto) // 3
        top.geometry(f"+{max(x, 0)}+{max(y, 0)}")
        top.deiconify()
        if modal:
            # En este orden: un grab_set sobre una ventana oculta falla
            # ("grab failed: window not viewable").
            top.wait_visibility()
            top.grab_set()

    # "Ejecutar todo" y "Generar antecedentes SAE" son de una sola
    # instancia: un segundo click trae al frente la que ya esta abierta.
    dialogos_abiertos = {}

    def traer_al_frente(clave):
        """True si ya habia una ventana abierta de esas (y quedo al frente)."""

        top = dialogos_abiertos.get(clave)
        try:
            if top is not None and top.winfo_exists():
                top.deiconify()
                top.lift()
                top.focus_set()
                return True
        except tk.TclError:
            pass
        dialogos_abiertos.pop(clave, None)
        return False

    def ventana_aviso(aviso):
        """
        El cuadro que abre el simbolo de aviso de una fila: el mensaje,
        un boton "Ver detalle" que abre el archivo con el detalle
        (Puntos_fallidos.xlsx) y "Cerrar".
        """

        ventana = tk.Toplevel(root)
        ventana.withdraw()
        ventana.title(aviso.get("titulo", "Aviso"))
        ventana.transient(root)
        ventana.resizable(False, False)

        cuerpo = tk.Frame(ventana, padx=18, pady=14)
        cuerpo.pack(fill="both", expand=True)

        tk.Label(
            cuerpo, text=SIMBOLO_AVISO, fg=COLOR_AVISO,
            font=("Segoe UI", 22, "bold"),
        ).pack(side="left", anchor="n", padx=(0, 14))

        tk.Label(
            cuerpo, text=aviso.get("mensaje", ""), justify="left",
            anchor="w", font=("Segoe UI", 9), wraplength=620,
        ).pack(side="left", fill="both", expand=True)

        botones = tk.Frame(ventana, pady=10)
        botones.pack(fill="x")

        def ver_detalle():
            ruta = aviso.get("ruta", "")
            if not abrir_archivo(ruta):
                messagebox.showwarning(
                    "No esta el archivo",
                    f"No se encontro el archivo con el detalle:\n{ruta}",
                    parent=ventana,
                )
                return
            ventana.destroy()

        tk.Button(
            botones, text="Cerrar", width=12, command=ventana.destroy,
        ).pack(side="right", padx=(6, 18))
        tk.Button(
            botones, text="Ver detalle", width=12,
            font=("Segoe UI", 9, "bold"), bg="#fdf0d5",
            command=ver_detalle,
        ).pack(side="right", padx=6)

        ventana.bind("<Escape>", lambda e: ventana.destroy())
        presentar(ventana, modal=True)
        ventana.focus_set()
        return ventana

    def _fila_arbol(parent, prefijo, texto, estado=None, negrita=False,
                    boton=None, id_fila=None, ruta="", es_carpeta=False,
                    origen=None, aviso=None):
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

        lbl_estado = tk.Label(
            celda_estado,
            text=SIMBOLO.get(estado, estado) if estado is not None else "",
            anchor="w",
            fg=COLOR_ESTADO.get(estado, COLOR_NEUTRO),
            font=("Segoe UI", 9, "bold"),
        )
        lbl_estado.pack(side="left")

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

        # El aviso va primero, pegado al boton: es lo que hay que ver.
        simbolo = _simbolo_aviso(celda_origen, aviso) if aviso else None

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

        if id_fila is not None:
            # Lo que puede cambiar sin rehacer la fila (ver pintar_arbol).
            filas_vivas[id_fila] = {
                "etiqueta": etiqueta, "estado": lbl_estado,
                "celda_origen": celda_origen, "simbolo": simbolo,
                "estilo": estilo,
            }

        return fila

    def _simbolo_aviso(celda_origen, aviso):
        """El simbolo de aviso de una fila, primero en su celda de origen."""

        simbolo = tk.Label(
            celda_origen, text=SIMBOLO_AVISO, fg=COLOR_AVISO,
            cursor="hand2", font=("Segoe UI", 12, "bold"),
        )
        hermanos = celda_origen.pack_slaves()
        if hermanos:
            simbolo.pack(side="left", padx=(0, 8), before=hermanos[0])
        else:
            simbolo.pack(side="left", padx=(0, 8))
        simbolo.bind("<Button-1>", lambda e, a=aviso: ventana_aviso(a))
        return simbolo

    def _actualizar_fila(id_fila, fila):
        """Cambia en el lugar lo que cambio de una fila ya dibujada."""

        piezas = filas_vivas[id_fila]
        ruta = fila.get("ruta", "")
        es_carpeta = fila.get("es_carpeta", False)
        estilo = piezas["estilo"]

        etiqueta = piezas["etiqueta"]
        etiqueta.config(
            text=fila["etiqueta"],
            font=("Consolas", 9, f"{estilo} underline" if ruta else estilo),
            fg=COLOR_LINK if ruta else "black",
            cursor="hand2" if ruta else "",
        )
        if ruta:
            etiqueta.bind(
                "<Button-1>",
                lambda e, r=ruta, c=es_carpeta: abrir_ruta_de_fila(r, c),
            )
        else:
            etiqueta.unbind("<Button-1>")

        estado = fila["estado"]
        piezas["estado"].config(
            text=SIMBOLO.get(estado, estado) if estado is not None else "",
            fg=COLOR_ESTADO.get(estado, COLOR_NEUTRO),
        )

        if piezas["simbolo"] is not None:
            piezas["simbolo"].destroy()
            piezas["simbolo"] = None
        if fila.get("aviso"):
            piezas["simbolo"] = _simbolo_aviso(
                piezas["celda_origen"], fila["aviso"]
            )

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

        if id_fila == "subastas_xlsx":
            return ("Consolidar", consolidar_subastas)

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

        if id_fila == "sae_csv":
            return ("Generar", generar_csv_sae)

        # Lo que recibe cada central y por que, en la hoja "Reporte" del
        # control (pedido del usuario: un boton "Generar reporte" al
        # lado de Control_corrida).
        if id_fila == "control":
            return ("Generar reporte", generar_reporte)

        if id_fila.startswith("consolidado:"):
            seccion = id_fila.split(":", 1)[1]
            return ("Actualizar", lambda s=seccion: actualizar_consolidado({s}))

        if id_fila.startswith("pagos:"):
            seccion = id_fila.split(":", 1)[1]
            return ("Calcular", lambda s=seccion: actualizar_pagos({s}))

        return None

    def _contenido(fila):
        """Lo de una fila que se puede cambiar sin rehacerla."""

        return (
            fila["etiqueta"], fila["estado"], str(fila.get("ruta", "")),
            fila.get("es_carpeta", False), fila.get("aviso"),
        )

    def pintar_arbol(filas):
        """
        Dibuja el arbol. Si tiene la misma FORMA que el que ya esta (las
        mismas filas, en el mismo orden y con el mismo origen), solo se
        cambia lo que cambio de cada una: sin destruir y recrear los
        ~400 widgets en cada "Actualizar" (en Windows eso parpadea). Si
        la forma cambio, se rehace entero.
        """

        prefijos = _prefijos_arbol([fila["nivel"] for fila in filas]) if filas else []
        forma = tuple(
            (
                fila["id"], fila["nivel"], prefijo,
                tuple(sorted((fila.get("origen") or {}).items())),
            )
            for fila, prefijo in zip(filas, prefijos)
        ) if filas and all(fila["id"] != "error" for fila in filas) else None

        if forma is not None and forma == arbol_pintado["forma"]:
            for fila in filas:
                contenido = _contenido(fila)
                if contenido != arbol_pintado["contenido"].get(fila["id"]):
                    _actualizar_fila(fila["id"], fila)
                    arbol_pintado["contenido"][fila["id"]] = contenido
            habilitar_botones(not corriendo["activo"])
            return

        for hijo in filas_arbol.winfo_children():
            hijo.destroy()

        botones_arbol.clear()
        filas_vivas.clear()
        arbol_pintado["forma"] = forma
        arbol_pintado["contenido"] = {}

        if not filas:
            tk.Label(
                filas_arbol,
                text="Selecciona una carpeta base.",
                fg=COLOR_NEUTRO,
                anchor="w",
            ).pack(fill="x")
            return

        for fila, prefijo in zip(filas, prefijos):
            arbol_pintado["contenido"][fila["id"]] = _contenido(fila)
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
                aviso=fila.get("aviso"),
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

    # Sin barra de progreso (pedido del usuario: "saca la barra de
    # progreso, deja solo el tiempo"): el tiempo va arriba a la derecha y
    # cada paso dice el suyo en el registro.

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

    for nivel, color in COLOR_REGISTRO.items():
        txt_log.tag_configure(nivel, foreground=color)
    txt_log.tag_configure("hora", foreground=COLOR_NEUTRO)
    # Cada paso (boton) como encabezado destacado, y su cierre con el
    # tiempo (pedido del usuario).
    txt_log.tag_configure(
        "encabezado", background="#244062", foreground="white",
        font=("Segoe UI", 9, "bold"), spacing1=6, spacing3=2,
    )
    txt_log.tag_configure(
        "encabezado_fin", foreground="#375623",
        font=("Segoe UI", 9, "bold"), spacing3=4,
    )

    # Inicio del proceso en curso: cada linea lleva la hora y el tiempo
    # transcurrido desde ahi ("control de tiempo", pedido del usuario).
    inicio_log = {"t": None}

    def _insertar_log(mensaje, ahora=None):
        texto = str(mensaje)
        nivel = nivel_del_mensaje(texto)
        ahora = time.time() if ahora is None else ahora
        for linea in texto.split("\n"):
            if linea.startswith(nucleo.ENCABEZADO_REGISTRO):
                txt_log.insert("end", f" {linea} \n", ("encabezado",))
                continue
            if linea.startswith(nucleo.FIN_REGISTRO):
                txt_log.insert("end", f"{linea}\n", ("encabezado_fin",))
                continue
            if linea.strip():
                marca = time.strftime("%H:%M:%S", time.localtime(ahora))
                if inicio_log["t"] is not None:
                    marca += f" +{formatear_reloj(ahora - inicio_log['t'])}"
                txt_log.insert("end", f"[{marca}] ", ("hora",))
            txt_log.insert("end", linea + "\n", (nivel,) if nivel else ())

    def log(mensaje):
        # Sin update_idletasks: forzar el redibujo en cada linea, durante
        # una corrida con cientos de mensajes, es lo que hacia que el
        # panel "Estructura del caso" se viera con filas repetidas al
        # desplazarlo (pedido del usuario).
        _insertar_log(mensaje)
        txt_log.see("end")

    # Los mensajes del hilo de trabajo llegan a esta cola y se escriben de
    # a tandas cada 100 ms (antes, un root.after por mensaje). Cada tanda
    # tiene tope: una rafaga de miles de lineas se escribe en varias
    # pasadas seguidas (cada 10 ms) en vez de congelar la ventana.
    cola_log = queue.Queue()
    sondeo_log = {"programado": None}

    def vaciar_cola_log(todo=False):
        """
        Escribe en el registro lo que llego por la cola. todo=True la
        vacia entera (al terminar un proceso, antes de la linea de
        cierre y de guardar el registro); si no, como mucho
        MENSAJES_POR_TANDA por pasada.
        """

        if sondeo_log["programado"] is not None:
            try:
                root.after_cancel(sondeo_log["programado"])
            except tk.TclError:
                pass
            sondeo_log["programado"] = None

        # Solo se baja al final si ya se estaba al final: asi se puede
        # leer hacia arriba durante una corrida.
        al_final = txt_log.yview()[1] >= 0.999
        escritos = 0
        while todo or escritos < MENSAJES_POR_TANDA:
            try:
                ahora, mensaje = cola_log.get_nowait()
            except queue.Empty:
                break
            _insertar_log(mensaje, ahora)
            escritos += 1
        if escritos and al_final:
            txt_log.see("end")
        if not cola_log.empty():
            sondeo_log["programado"] = root.after(10, vaciar_cola_log)
        elif corriendo["activo"]:
            sondeo_log["programado"] = root.after(100, vaciar_cola_log)

    def registrar_desde_hilo(mensaje):
        cola_log.put((time.time(), mensaje))

    def guardar_registro():
        """Deja el registro en <caso>/Log/Registro_ultima_corrida.txt."""

        base = var_base.get()
        if not base or not os.path.isdir(base):
            return None
        carpeta_log = os.path.join(base, nucleo.CARPETA_LOG)
        ruta = os.path.join(carpeta_log, nucleo.ARCHIVO_REGISTRO)
        try:
            os.makedirs(carpeta_log, exist_ok=True)
            with open(ruta, "w", encoding="utf-8") as archivo:
                archivo.write(txt_log.get("1.0", "end"))
        except OSError:
            return None
        return ruta

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

        aamm = var_aamm.get().strip()

        try:
            nucleo.validar_aamm(aamm)
        except nucleo.ErrorEntrada:
            aamm = ""

        # Una sola escritura de config.json: la carpeta base y, si hay
        # periodo, la carpeta de ese periodo.
        recordar_carpeta(aamm, ruta, ademas={"carpeta_base": ruta})
        if aamm:
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

    def actualizar_vista():
        """
        Vuelve a mirar las carpetas del caso y repinta el diagrama: un
        archivo pegado a mano en el explorador no aparece hasta que
        algo llame a revisar(). Antes habia que cambiar de mes y
        volver. Mientras corre un proceso no hace nada: repintar el
        arbol recrea sus botones, que en ese momento estan
        deshabilitados a proposito.
        """

        if corriendo["activo"]:
            return

        revisar()
        log("Carpetas revisadas de nuevo.")

    frame_botones_carpeta = tk.Frame(frame_carpeta)
    frame_botones_carpeta.pack(pady=(6, 0))
    tk.Button(
        frame_botones_carpeta, text="Examinar", command=seleccionar
    ).pack(side="left", padx=4)
    tk.Button(
        frame_botones_carpeta, text="Actualizar", command=actualizar_vista
    ).pack(side="left", padx=4)

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
        top.withdraw()
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

        presentar(top, modal=True)
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

        texto_cambio, hay_cambio = mostrar_cambio_hora()
        if hay_cambio:
            log(f"Periodo {aamm}: {texto_cambio}")

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
    entry_aamm.bind("<Return>", lambda e: root.focus_set())
    boton_ir_mes.config(command=lambda: (root.focus_set(), aamm_cambiado()))

    # --------------------------------------------------------
    # CURSOR DE TEXTO
    #
    # Pedido del usuario: "el cursor de texto se queda parpadeando
    # siempre en el cuadro de AAMM... que se quite cuando se hace clic
    # en otra parte". En Tk un clic en un boton, etiqueta o fondo NO
    # le quita el foco al campo; aca se lo quita cualquier clic fuera
    # de un campo de texto (el foco pasa a la ventana, el cursor deja
    # de parpadear y <FocusOut> aplica el AAMM escrito, como antes).
    # --------------------------------------------------------

    CAMPOS_DE_TEXTO = (tk.Entry, tk.Text, tk.Spinbox, ttk.Entry)

    def soltar_foco(evento):
        widget = evento.widget
        if isinstance(widget, str) or isinstance(widget, CAMPOS_DE_TEXTO):
            return
        try:
            enfocado = widget.focus_get()
        except (KeyError, tk.TclError):
            return
        if isinstance(enfocado, (tk.Entry, tk.Spinbox, ttk.Entry)):
            widget.winfo_toplevel().focus_set()

    root.bind_all("<Button-1>", soltar_foco, add="+")

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

    TITULO_DE_FUNCION = {
        "traer_csv_cmg": "Traer cmg_15min",
        "traer_fd": "Traer FD (SSCC_Desempeño)",
        "generar_fma": "Generar FMA",
        "traer_subastas": "Traer subastas",
        "consolidar_subastas": "Consolidar subastas",
        "generar_cmg": "Generar cmg.xlsx",
        "generar_medidas_sae": "Actualizar Medidas_SAE",
        "generar_consolidado": "Planilla 11_PAGOS_BESS (entradas)",
        "generar_pagos_bess": "Planilla 11_PAGOS_BESS (calculo)",
        "ejecutar_plan": "Ejecutar todo",
        "generar_csv_sae": "Generar SAE_AAMM.csv",
        "generar_reporte_corrida": "Generar reporte de la corrida",
        "generar_antecedentes": "Generar antecedentes SAE",
        "correr": "Planilla 11_PAGOS_BESS (entradas y calculo)",
    }

    def titulo_de(funcion, kwargs):
        titulo = TITULO_DE_FUNCION.get(
            getattr(funcion, "__name__", ""), getattr(funcion, "__name__", "proceso")
        )
        secciones = kwargs.get("secciones_activas") or kwargs.get("tipos")
        if secciones:
            titulo += ": " + ", ".join(sorted(str(s) for s in secciones))
        return titulo

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
        txt_log.delete("1.0", "end")

        titulo = titulo_de(funcion, kwargs)
        timer["corriendo"] = True
        timer["inicio"] = time.time()
        inicio_log["t"] = timer["inicio"]
        log(f"{nucleo.ENCABEZADO_REGISTRO}{titulo}")
        log(f"Inicio: {time.strftime('%d-%m-%Y %H:%M:%S')}")
        var_estado.set("Procesando...")
        tick()
        vaciar_cola_log()

        resultado = {"ok": False}

        def trabajo():
            try:
                funcion(
                    **kwargs,
                    registrar=registrar_desde_hilo,
                    progreso=lambda _v: None,
                )
                resultado["ok"] = True

            except nucleo.ErrorEntrada as error:
                registrar_desde_hilo(f"\nERROR DE ENTRADA:\n{error}")
                resultado["error"] = str(error)

            except Exception as error:
                registrar_desde_hilo(
                    f"\nERROR INESPERADO:\n{error}\n\n{traceback.format_exc()}",
                )
                resultado["error"] = str(error)

            finally:
                root.after(0, terminar, resultado)

        def terminar(res):
            timer["corriendo"] = False
            corriendo["activo"] = False
            vaciar_cola_log(todo=True)

            duracion = time.time() - timer["inicio"]
            log("")
            log(
                f"{nucleo.FIN_REGISTRO}{titulo}: "
                f"{'listo' if res['ok'] else 'termino con problemas'} en "
                f"{nucleo.formatear_duracion(duracion)}"
            )
            guardado = guardar_registro()
            if guardado:
                log(f"Registro guardado en {guardado}")
            inicio_log["t"] = None

            # El arbol ya no se rehace entero en cada revisar() (reusa
            # las filas): los botones se rehabilitan aca.
            habilitar_botones(True)
            revisar()

            if res["ok"]:
                var_estado.set("Terminado correctamente.")
                messagebox.showinfo(
                    "Listo", f"Se genero/actualizo:\n{etiqueta_salida}"
                )
            else:
                var_estado.set("Termino con errores.")
                mostrar_error(
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

    def consolidar_subastas():
        """
        Junta los Access de Subastas/DB subastas/ en
        Subastas/subastas_AAMM.xlsx. La hoja 'Subastas' del libro se
        lleva despues de ahi, con su propio boton.
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
            nucleo.consolidar_subastas,
            dict(carpeta_base=ruta, aamm=aamm),
            f"subastas_{aamm}.xlsx",
        )

    def nombre_planilla():
        """11_PAGOS_BESS_AAMM.xlsx con el periodo escrito arriba."""

        aamm = var_aamm.get().strip()
        try:
            return nucleo.nombre_salida(nucleo.validar_aamm(aamm))
        except nucleo.ErrorEntrada:
            return nucleo.ARCHIVO_SALIDA

    def generar_csv_sae():
        """
        El ultimo boton: SAE_AAMM.csv, el CSV de pagos que se sube, desde
        Calculo E Costos y Calculo RE545 (lo que hacia la macro
        Generar_CSV_SAE).
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
            nucleo.generar_csv_sae,
            dict(carpeta_base=ruta, aamm=aamm),
            nucleo.nombre_csv_sae(aamm),
        )

    def generar_reporte():
        """
        Hoja "Reporte" de Log/Control_corrida.xlsx: lo que recibe cada
        central y, si no recibe nada (o solo por un lado), por que
        (ver nucleo/reporte.py).
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
            nucleo.generar_reporte_corrida,
            dict(carpeta_base=ruta, aamm=aamm),
            nucleo.ARCHIVO_CONTROL,
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

        # Sin confirmacion previa (pedido del usuario): el avance y lo
        # que tarda ya se ven en el registro y la barra.
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
            nombre_planilla(),
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
            nombre_planilla(),
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

        lanzar(correr, {}, nombre_planilla())

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
            mostrar_error("No se pudo crear", str(error))
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
        top.withdraw()
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

        presentar(top, modal=True, ancho=760, alto=460)
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

        if traer_al_frente("ejecutar_todo"):
            return

        ruta = caso_listo()
        if ruta is None:
            return

        if corriendo["activo"]:
            return

        aamm = var_aamm.get().strip()

        # None = "lo que falte" (el arranque normal); despues pasa a
        # ser el set que el usuario va tildando.
        seleccion = {"ids": None}

        # La lista se arma UNA vez y despues solo se actualizan sus
        # valores (pedido del usuario: "cuando selecciono una casilla hace
        # un parpadeo"). Antes cada tilde borraba y volvia a crear todas
        # las filas -y volvia a revisar la carpeta entera-, y la lista
        # quedaba en blanco un instante. La estructura de la carpeta no
        # cambia mientras la ventana esta abierta: se revisa al abrirla,
        # ANTES de crear la ventana (si falla no queda una vacia abierta).
        try:
            _, filas_plan = nucleo.revisar_estructura(ruta, aamm)
        except Exception as error:  # noqa: BLE001 - se muestra
            mostrar_error("No se pudo revisar la carpeta", str(error))
            return

        top = tk.Toplevel(root)
        top.withdraw()
        top.title("Ejecutar todo")
        top.geometry("900x640")
        top.transient(root)
        dialogos_abiertos["ejecutar_todo"] = top

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

        lienzo = tk.Canvas(
            frame_pasos, borderwidth=0, highlightthickness=0,
            yscrollincrement=1,
        )
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

        lista.bind("<Configure>", ajustar_lista)
        lienzo.bind(
            "<Configure>", lambda e: ajustar_ancho(lienzo, ventana_lista, e)
        )

        frame_avisos = tk.LabelFrame(
            top, text="Para poder ejecutar", padx=8, pady=6
        )
        frame_avisos.pack(fill="x", padx=14, pady=4)

        txt_avisos = tk.Text(
            frame_avisos, height=6, font=("Segoe UI", 8), wrap="word"
        )
        txt_avisos.pack(fill="x")

        # La rueda sobre la lista desplaza la lista, y sobre los avisos
        # los avisos: nunca la ventana principal que esta detras.
        def desplazar_lista(muescas):
            region = lienzo.bbox("all")
            if not region or region[3] - region[1] <= lienzo.winfo_height():
                return
            pixeles = _acumular("ejecutar_todo", muescas * PIXELES_POR_MUESCA)
            if pixeles:
                _desplazar_por_cuadro(lienzo, pixeles)

        def desplazar_avisos(muescas):
            lineas = _acumular("ejecutar_todo_avisos", muescas * LINEAS_POR_MUESCA)
            if lineas:
                txt_avisos.yview_scroll(lineas, "units")

        anotados = (frame_pasos, lienzo, lista, barra_pasos)
        for widget in anotados:
            registrar_desplazable(widget, desplazar_lista)
        for widget in (frame_avisos, txt_avisos):
            registrar_desplazable(widget, desplazar_avisos)

        def olvidar_desplazables(event):
            if event.widget is top:
                for widget in anotados + (frame_avisos, txt_avisos):
                    desplazables.pop(str(widget), None)

        top.bind("<Destroy>", olvidar_desplazables, add="+")

        frame_pie = tk.Frame(top)
        frame_pie.pack(fill="x", padx=14, pady=(4, 12))

        btn_ejecutar = tk.Button(
            frame_pie, text="Ejecutar", font=("Segoe UI", 9, "bold"),
            bg="#fdf0d5",
        )

        estado_plan = {"plan": None}

        vista = {
            "armada": False,
            "seleccionadas": set(),
            "tareas": {},     # id -> (var, cuadro, etiqueta de estado, motivo)
            "libro": None,    # (var, cuadro, contador, ids)
        }

        COLOR_ESTADO_PLAN = {
            "al dia": COLOR_OK,
            "se genera": COLOR_BASE_OK,
            "se rehace": COLOR_PENDIENTE,
            "bloqueada": COLOR_FALTA,
            "sin tildar": COLOR_NEUTRO,
        }

        def cambiar(ids, tildar):
            actual = set(vista["seleccionadas"])
            if tildar:
                # Tildar algo arrastra lo que sale de ahi: si no, queda
                # mezclado con la corrida anterior.
                actual |= nucleo.propagar_seleccion(set(ids))
            else:
                actual -= set(ids)
            seleccion["ids"] = actual
            repintar()

        def armar(plan):
            # La lista se arma con el mismo orden y la misma sangria que
            # el arbol de la ventana principal (pedido del usuario): cada
            # cosa con boton bajo su carpeta, y las hojas del libro
            # adentro de Balance_BESS.xlsx, que tiene su propio cuadro
            # para tildarlas o destildarlas todas juntas.
            tarea_de_fila = {}
            for tarea in plan["tareas"]:
                fila_estado = nucleo.TAREA_POR_ID[tarea["id"]].fila_estado
                tarea_de_fila.setdefault(fila_estado, tarea)

            def pintar_tarea(tarea, sangria):

                fila = tk.Frame(lista)
                fila.pack(fill="x", pady=1)

                var = tk.BooleanVar(value=tarea["seleccionada"])
                cuadro = tk.Checkbutton(
                    fila, variable=var,
                    command=lambda i=tarea["id"], v=var: cambiar({i}, v.get()),
                )
                cuadro.pack(side="left", padx=(sangria, 0))

                tk.Label(
                    fila, text=tarea["etiqueta"], width=44 - sangria // 8,
                    anchor="w", font=("Consolas", 9),
                ).pack(side="left")

                estado = tk.Label(
                    fila, width=12, anchor="w", font=("Segoe UI", 9, "bold"),
                )
                estado.pack(side="left")

                motivo = tk.Label(
                    fila, anchor="w", fg=COLOR_NEUTRO, font=("Segoe UI", 8),
                )
                motivo.pack(side="left", fill="x", expand=True)

                vista["tareas"][tarea["id"]] = (var, cuadro, estado, motivo)

            def pintar_titulo(texto, sangria):
                tk.Label(
                    lista, text=texto, anchor="w",
                    font=("Segoe UI", 9, "bold"), fg=COLOR_NEUTRO,
                ).pack(fill="x", padx=(sangria, 0), pady=(6, 0))

            def pintar_libro(etiqueta, hojas):
                fila = tk.Frame(lista)
                fila.pack(fill="x", pady=(6, 1))
                ids = {t["id"] for t in hojas}
                var = tk.BooleanVar()
                cuadro = tk.Checkbutton(
                    fila, variable=var,
                    command=lambda v=var: cambiar(ids, v.get()),
                )
                cuadro.pack(side="left")
                tk.Label(
                    fila, text=etiqueta, anchor="w",
                    font=("Segoe UI", 9, "bold"),
                ).pack(side="left")
                contador = tk.Label(
                    fila, anchor="w", fg=COLOR_NEUTRO, font=("Segoe UI", 8),
                )
                contador.pack(side="left")
                vista["libro"] = (var, cuadro, contador, ids)

            pintadas = set()
            titulo_pendiente = None

            for fila_arbol in filas_plan:
                id_fila = fila_arbol["id"]

                if id_fila == "salida":
                    hojas = [
                        tarea_de_fila[f["id"]] for f in filas_plan
                        if f["id"] in tarea_de_fila
                        and f["id"].split(":", 1)[0] in ("consolidado", "pagos")
                    ]
                    if hojas:
                        pintar_libro(fila_arbol["etiqueta"], hojas)
                    titulo_pendiente = None
                    continue

                if fila_arbol["nivel"] == 0:
                    # Una carpeta: se escribe solo si tiene algo con
                    # boton adentro.
                    titulo_pendiente = fila_arbol["etiqueta"]
                    continue

                tarea = tarea_de_fila.get(id_fila)
                if tarea is None or tarea["id"] in pintadas:
                    continue

                if titulo_pendiente:
                    pintar_titulo(titulo_pendiente, 0)
                    titulo_pendiente = None

                pintar_tarea(tarea, 20 * fila_arbol["nivel"])
                pintadas.add(tarea["id"])

            # Lo que no tenga fila en el arbol (no deberia pasar): al
            # final, para que nunca quede una tarea sin mostrar.
            for tarea in plan["tareas"]:
                if tarea["id"] not in pintadas:
                    pintar_tarea(tarea, 0)

            vista["armada"] = True

        def repintar():

            plan = nucleo.planificar(
                ruta, aamm, seleccion["ids"], filas=filas_plan
            )
            estado_plan["plan"] = plan

            if not vista["armada"]:
                armar(plan)

            seleccionadas = {
                t["id"] for t in plan["tareas"] if t["seleccionada"]
            }
            vista["seleccionadas"] = seleccionadas
            deshabilitado = "disabled" if corriendo["activo"] else "normal"

            for tarea in plan["tareas"]:
                widgets = vista["tareas"].get(tarea["id"])
                if widgets is None:
                    continue
                var, cuadro, estado, motivo = widgets
                var.set(tarea["seleccionada"])
                cuadro.config(state=deshabilitado)
                estado.config(
                    text=tarea["estado"], fg=COLOR_ESTADO_PLAN[tarea["estado"]]
                )
                faltan = ", ".join(
                    tarea["faltan_requisitos"] + tarea["faltan_dependencias"]
                )
                motivo.config(text=f"falta: {faltan}" if faltan else "")

            if vista["libro"] is not None:
                var, cuadro, contador, ids = vista["libro"]
                var.set(ids <= seleccionadas)
                cuadro.config(state=deshabilitado)
                contador.config(
                    text=(
                        f"  {len(ids & seleccionadas)} de {len(ids)} "
                        "hoja(s) tildada(s)"
                    )
                )

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
                nombre_planilla(),
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
        presentar(top, ancho=900, alto=640)

    # --------------------------------------------------------
    # ANTECEDENTES SAE
    #
    # Pedido del usuario (2026-11-10): un boton que cree, al lado del
    # 11_, la carpeta "03 Antecedentes SAE" con el codigo para que los
    # coordinados recalculen las hojas (y el SAE_AAMM.csv). Por defecto
    # van esos dos; "Agregar otros archivos" despliega los archivos del
    # caso para sumar los que se quieran.
    # --------------------------------------------------------

    def ventana_antecedentes():
        if traer_al_frente("antecedentes"):
            return
        ruta = caso_listo()
        if ruta is None:
            return
        aamm = var_aamm.get().strip()
        try:
            nucleo.validar_aamm(aamm)
        except nucleo.ErrorEntrada as error:
            messagebox.showwarning("Falta el periodo", str(error))
            return
        try:
            rutas, filas = nucleo.revisar_estructura(ruta, aamm)
        except Exception as error:  # noqa: BLE001 - se muestra
            mostrar_error("No se pudo revisar la carpeta", str(error))
            return
        if not rutas["salida"].is_file():
            messagebox.showwarning(
                "Falta la planilla",
                f"No esta {rutas['salida'].name}: primero calcula la planilla.",
            )
            return

        top = tk.Toplevel(root)
        top.withdraw()
        top.title("Generar antecedentes SAE")
        top.transient(root)
        top.geometry("720x260")
        dialogos_abiertos["antecedentes"] = top

        tk.Label(
            top,
            text=(
                f"Se crea la carpeta '{nucleo.CARPETA_ANTECEDENTES}' al lado de "
                f"{rutas['salida'].name} con lo tildado."
            ),
            anchor="w", justify="left", wraplength=680,
            font=("Segoe UI", 9),
        ).pack(fill="x", padx=14, pady=(12, 6))

        marco = tk.LabelFrame(top, text="Por defecto", padx=10, pady=6)
        marco.pack(fill="x", padx=14, pady=4)
        var_codigo = tk.BooleanVar(value=True)
        tk.Checkbutton(
            marco, variable=var_codigo, anchor="w",
            text=f"Código de recálculo ({nucleo.nombre_recalculador(aamm)})",
        ).pack(fill="x")
        csv = rutas["csv_sae"]
        var_csv = tk.BooleanVar(value=csv.is_file())
        tk.Checkbutton(
            marco, variable=var_csv, anchor="w",
            text=csv.name + ("" if csv.is_file() else "   (todavía no generado)"),
            state="normal" if csv.is_file() else "disabled",
        ).pack(fill="x")

        # Los archivos del caso, para sumar los que se quieran.
        excluir = {Path(rutas["salida"]), Path(csv)}
        archivos = [
            fila for fila in filas
            if fila.get("ruta") and not fila.get("es_carpeta")
            and Path(fila["ruta"]).is_file()
            and Path(fila["ruta"]) not in excluir
        ]
        otros = tk.LabelFrame(top, text="Otros archivos del caso", padx=10, pady=6)
        vars_otros = []
        for fila in archivos:
            var = tk.BooleanVar(value=False)
            vars_otros.append((var, fila["ruta"]))
            tk.Checkbutton(
                otros, variable=var, anchor="w", font=("Consolas", 9),
                text="    " * int(fila.get("nivel", 0)) + Path(fila["ruta"]).name,
            ).pack(fill="x")
        if not archivos:
            tk.Label(otros, text="No hay otros archivos en la carpeta del caso.",
                     fg=COLOR_NEUTRO).pack(anchor="w")

        abajo = tk.Frame(top)
        abajo.pack(side="bottom", fill="x", padx=14, pady=10)
        desplegado = {"si": False}

        def desplegar():
            desplegado["si"] = not desplegado["si"]
            if desplegado["si"]:
                otros.pack(fill="both", expand=True, padx=14, pady=4, before=abajo)
                boton_otros.config(text="Agregar otros archivos ▴")
                top.geometry(f"720x{min(260 + 24 * max(len(archivos), 1), 760)}")
            else:
                otros.pack_forget()
                boton_otros.config(text="Agregar otros archivos ▾")
                top.geometry("720x260")

        boton_otros = tk.Button(abajo, text="Agregar otros archivos ▾", command=desplegar)
        boton_otros.pack(side="left")

        def generar():
            seleccion = [r for var, r in vars_otros if var.get()]
            if not (var_codigo.get() or var_csv.get() or seleccion):
                messagebox.showwarning("Nada seleccionado", "Tilda al menos un archivo.", parent=top)
                return
            top.destroy()
            lanzar(
                nucleo.generar_antecedentes,
                dict(carpeta_base=ruta, aamm=aamm, incluir_codigo=var_codigo.get(),
                     incluir_csv=var_csv.get(), otros=seleccion),
                nucleo.CARPETA_ANTECEDENTES,
            )

        tk.Button(abajo, text="Cancelar", width=12, command=top.destroy).pack(side="right")
        tk.Button(
            abajo, text="Generar", width=12, bg="#1f5fa8", fg="white",
            font=("Segoe UI", 9, "bold"), command=generar,
        ).pack(side="right", padx=8)

        presentar(top, ancho=720, alto=260)

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

    tk.Button(
        frame_botones,
        text="Generar antecedentes SAE",
        command=ventana_antecedentes,
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
