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

La carpeta que se elige es la del MES, la misma que usa Balance_CRA.py:
lo propio del BESS va en <mes>/BESS/ y lo comun con el CRA en la raiz
(ver Script/nucleo/carpetas.py). Un caso de antes, con todo suelto en
la carpeta, se sigue usando tal cual, sin mover nada.

    <CARPETA DEL MES>/
        BESS/
            Medidas/
                Medidas_SAE.xlsx           [Actualizar]
                SOC_AAMM.xlsx              [Traer]
            Auxiliares/
                Centrales.xlsx
                    hoja 'Resumen BESS'
                    hoja 'Diccionario'
                <algo>Homologacion<algo>.xlsx
                    hoja 'homol'
                    hoja 'Gen real'
            Ofertas/
                <algo>OfertasSSCC<algo>.xlsx (o .xlsm/.xlsb/.xls)
            Subastas/
                DB subastas/               [Traer]
            11_PAGOS_BESS_AAMM.xlsx        [Actualizar]
                hoja 'Resumen' ... hoja 'Medidores'  [uno por hoja]
            SAE_AAMM.csv                   [Generar]
            Log/
                Control_corrida.xlsx
                Registro_ultima_corrida.txt
        CRA/                               (Balance_CRA.py)
        Cmg/                               (comun con el CRA)
            cmg<AAMM>_def_15minutal.csv    [Traer]
            cmg.xlsx                       [Generar]
        FD y FMA/                          (comun con el CRA)
            SSCC_Desempeño_<algo>.xlsx     [Traer]
            fma_cpf_<AAMM>.xlsx            [Generar]
            fma_csf_<AAMM>.xlsx            [Generar]
            (fma_cft_<AAMM>.xlsx: oculto, FMA_CON_CTF = False)
        Prorrata retiros/                  (comun con el CRA)

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
import socket
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


# El aspecto (colores, anchos, rueda, registro, abrir carpetas) es comun
# con Balance_CRA.py: vive en Script/ventana_estilo.py.
from Script.ventana_estilo import (  # noqa: E402,F401
    COLOR_OK,
    COLOR_FALTA,
    COLOR_PENDIENTE,
    COLOR_NEUTRO,
    COLOR_BASE_OK,
    COLOR_LINK,
    SIMBOLO,
    SIMBOLO_AVISO,
    COLOR_AVISO,
    COLOR_REGISTRO,
    nivel_del_mensaje,
    formatear_reloj,
    COLOR_ESTADO,
    ANCHO_ESTRUCTURA,
    ANCHO_ESTADO,
    ANCHO_ACCION,
    ALTO_ACCION,
    ALTO_FILA,
    RELLENO_ACCION,
    PIXELES_POR_MUESCA,
    LINEAS_POR_MUESCA,
    EVENTOS_RUEDA,
    MENSAJES_POR_TANDA,
    carpeta_a_abrir,
    abrir_en_explorador,
    abrir_archivo,
    formato_tiempo,
)


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


# ============================================================
# EL BALANCE DEL CASO: AÑO_MES_PROCESO
#
# Pedido del usuario (2026-11-18): en vez de escribir el AAMM, tres
# casillas -año, mes escrito y proceso- que se leen como
# "2026_Marzo_Definitivo". Cada balance (mes + proceso) tiene su propia
# carpeta: el Preliminar y el Definitivo del mismo mes son dos casos.
# Por dentro el calculo sigue usando el AAMM ("2603"), que sale del año
# y el mes.
# ============================================================

ANO_INICIAL = 2025

MESES = (
    "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio",
    "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre",
)

# Los que estan siempre; los demas (R01P, R01D, ...) los agrega el
# usuario con el "+" de la lista y quedan en config.json.
PROCESOS_FIJOS = ("Preliminar", "Definitivo")
OPCION_AGREGAR_PROCESO = "+ Agregar otro..."
CLAVE_PROCESOS = "procesos"


def anos_disponibles(ano_actual=None):
    """['2025', '2026', ...] hasta el año de hoy."""

    actual = int(ano_actual or time.localtime().tm_year)
    return [str(ano) for ano in range(ANO_INICIAL, max(actual, ANO_INICIAL) + 1)]


def aamm_de(ano, mes):
    """('2026', 'Marzo') -> '2603'. '' si falta algo o no es valido."""

    try:
        numero = MESES.index(str(mes)) + 1
        ano = int(ano)
    except (ValueError, TypeError):
        return ""
    return f"{ano % 100:02d}{numero:02d}"


def ano_mes_de(aamm):
    """'2603' -> ('2026', 'Marzo'). ('', '') si no es un AAMM."""

    aamm = str(aamm or "").strip()
    if len(aamm) != 4 or not aamm.isdigit() or not 1 <= int(aamm[2:]) <= 12:
        return "", ""
    return str(2000 + int(aamm[:2])), MESES[int(aamm[2:]) - 1]


def nombre_balance(ano, mes, proceso):
    """'2026_Marzo_Definitivo' (lo que falte no se escribe)."""

    return "_".join(str(p) for p in (ano, mes, proceso) if p)


def clave_balance(aamm, proceso):
    """La clave de carpetas_por_periodo: '2603_Definitivo'."""

    return f"{aamm}_{proceso}" if proceso else str(aamm)


def limpiar_proceso(texto):
    """
    El nombre de un proceso agregado a mano: sin espacios ni
    caracteres que no van en un nombre de carpeta. '' si no queda nada.
    """

    texto = "".join(
        c for c in str(texto or "").strip()
        if c.isalnum() or c in "-"
    )
    return texto


def procesos_guardados():
    """Preliminar, Definitivo y los que agrego este usuario."""

    extra = leer_config().get(CLAVE_PROCESOS)
    procesos = list(PROCESOS_FIJOS)
    if isinstance(extra, list):
        for proceso in extra:
            proceso = limpiar_proceso(proceso)
            if proceso and proceso.lower() not in (p.lower() for p in procesos):
                procesos.append(proceso)
    return procesos


def proceso_en_el_nombre(ruta, procesos=None):
    """
    El proceso que dice el nombre de la carpeta ('... 2603 Definitivo' ->
    'Definitivo'), o '' si no dice ninguno (o dice mas de uno).
    """

    nombre = Path(str(ruta or "")).name.lower()
    hallados = [
        proceso for proceso in (procesos or PROCESOS_FIJOS)
        if proceso.lower() in nombre
    ]
    return hallados[0] if len(hallados) == 1 else ""


def carpeta_del_balance(aamm, proceso):
    """
    La carpeta de ese balance, si ya tiene una y todavia existe; '' si
    no. Una carpeta anotada solo con el AAMM (de antes de que hubiera
    proceso) sirve si su nombre dice ese proceso.
    """

    carpetas = carpetas_por_periodo()
    ruta = carpetas.get(clave_balance(aamm, proceso), "")
    if not ruta and proceso:
        vieja = carpetas.get(str(aamm), "")
        if vieja and proceso.lower() in Path(vieja).name.lower():
            ruta = vieja
    try:
        return ruta if ruta and Path(ruta).is_dir() else ""
    except OSError:
        return ""


def balance_de_la_carpeta(ruta):
    """La clave del balance que ya usa esa carpeta ('2602_Definitivo'), o ''."""

    if not ruta:
        return ""
    for clave, otra in carpetas_por_periodo().items():
        try:
            if Path(otra).resolve() == Path(ruta).resolve():
                return clave
        except OSError:
            if str(otra) == str(ruta):
                return clave
    return ""


def nombre_de_clave(clave):
    """'2602_Definitivo' -> '2026_Febrero_Definitivo'; '2602' -> '2026_Febrero'."""

    aamm, _, proceso = str(clave).partition("_")
    ano, mes = ano_mes_de(aamm)
    return nombre_balance(ano, mes, proceso) if ano else str(clave)

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
    # El proceso del caso abierto (Preliminar, Definitivo, R01P...). Un
    # config de antes no lo tiene: se toma del nombre de la carpeta.
    var_proceso = tk.StringVar(
        value=limpiar_proceso(cfg.get("proceso", ""))
        or proceso_en_el_nombre(cfg.get("carpeta_base", ""), procesos_guardados())
    )
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
    # EL BALANCE DEL CASO: AÑO _ MES _ PROCESO
    #
    # Pedido del usuario (2026-11-18): tres casillas que se leen
    # "2026_Marzo_Definitivo", la ruta del balance al lado y dos botones:
    # "Ir a este balance" (azul si el balance ya tiene carpeta, bloqueado
    # si no) y "Examinar" (destacado cuando no la tiene). Elegir en las
    # casillas NO cambia el caso abierto: recien lo cambian los botones.
    # var_aamm / var_proceso son los del caso ABIERTO.
    # --------------------------------------------------------

    frame_periodo = tk.LabelFrame(
        contenedor, text="Balance del caso", padx=10, pady=6
    )
    frame_periodo.pack(fill="x", padx=14, pady=(10, 4))

    fila_balance = tk.Frame(frame_periodo)
    fila_balance.pack(fill="x")

    ano_abierto, mes_abierto = ano_mes_de(var_aamm.get())
    hoy = time.localtime()
    var_sel_ano = tk.StringVar(value=ano_abierto or str(max(hoy.tm_year, ANO_INICIAL)))
    var_sel_mes = tk.StringVar(value=mes_abierto or MESES[hoy.tm_mon - 1])
    var_sel_proceso = tk.StringVar(
        value=var_proceso.get() or PROCESOS_FIJOS[-1]
    )

    def _casilla(valores, variable, ancho):
        caja = ttk.Combobox(
            fila_balance, textvariable=variable, values=valores,
            state="readonly", width=ancho, font=("Consolas", 11, "bold"),
            justify="center",
        )
        caja.pack(side="left")
        return caja

    def _guion():
        tk.Label(
            fila_balance, text="_", font=("Consolas", 11, "bold"),
        ).pack(side="left", padx=2)

    _casilla(anos_disponibles(), var_sel_ano, 6)
    _guion()
    _casilla(list(MESES), var_sel_mes, 11)
    _guion()
    cb_proceso = _casilla(
        procesos_guardados() + [OPCION_AGREGAR_PROCESO], var_sel_proceso, 16,
    )

    # Nombre del balance y su ruta (la ruta es un link a la carpeta).
    fila_ruta = tk.Frame(frame_periodo)
    fila_ruta.pack(fill="x", pady=(6, 0))
    var_nombre_balance = tk.StringVar(value="")
    var_ruta_balance = tk.StringVar(value="")
    tk.Label(
        fila_ruta, textvariable=var_nombre_balance,
        font=("Segoe UI", 10, "bold"), anchor="w",
    ).pack(side="left")
    lbl_ruta_balance = tk.Label(
        fila_ruta, textvariable=var_ruta_balance, font=("Segoe UI", 9),
        anchor="w", justify="left", wraplength=820, cursor="hand2",
    )
    lbl_ruta_balance.pack(side="left", padx=(4, 0))

    fila_botones_balance = tk.Frame(frame_periodo)
    fila_botones_balance.pack(anchor="w", pady=(6, 0))
    btn_ir_balance = tk.Button(
        fila_botones_balance, text="Ir a este balance", width=18,
        font=("Segoe UI", 9, "bold"),
    )
    btn_ir_balance.pack(side="left")
    btn_examinar_balance = tk.Button(
        fila_botones_balance, text="Examinar", width=14,
    )
    btn_examinar_balance.pack(side="left", padx=8)

    # Cambio de hora del mes elegido: se detecta solo con el AAMM (ver
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

    def seleccion():
        """(aamm, proceso, nombre) de lo elegido en las casillas."""

        proceso = var_sel_proceso.get()
        if proceso == OPCION_AGREGAR_PROCESO:
            proceso = ""
        return (
            aamm_de(var_sel_ano.get(), var_sel_mes.get()),
            proceso,
            nombre_balance(var_sel_ano.get(), var_sel_mes.get(), proceso),
        )

    def mostrar_cambio_hora(*_):
        aamm = seleccion()[0]
        try:
            nucleo.validar_aamm(aamm)
            texto = nucleo.describir_periodo(aamm)
            hay = bool(nucleo.cambios_del_periodo(aamm))
        except (nucleo.ErrorEntrada, ValueError):
            texto, hay = "", False
        var_cambio_hora.set(texto)
        label_cambio_hora.config(fg=COLOR_AVISO if hay else COLOR_NEUTRO)
        return texto, hay

    def refrescar_balance(*_):
        """La ruta y los dos botones segun lo elegido en las casillas."""

        aamm, proceso, nombre = seleccion()
        ruta = carpeta_del_balance(aamm, proceso) if aamm and proceso else ""
        abierto = bool(ruta) and ruta == var_base.get() and aamm == var_aamm.get() \
            and proceso == var_proceso.get()

        var_nombre_balance.set(nombre)
        if ruta:
            var_ruta_balance.set(f"- {ruta}" + ("   (abierto)" if abierto else ""))
            lbl_ruta_balance.config(fg=COLOR_BASE_OK if abierto else COLOR_LINK)
            btn_ir_balance.config(state="normal", bg="#1f5fa8", fg="white")
            btn_examinar_balance.config(
                bg=COLOR_FONDO_BOTON, fg="black", font=("Segoe UI", 9),
            )
        else:
            var_ruta_balance.set("- (sin carpeta: elegila con Examinar)")
            lbl_ruta_balance.config(fg=COLOR_FALTA)
            btn_ir_balance.config(
                state="disabled", bg=COLOR_FONDO_BOTON, fg="black",
            )
            btn_examinar_balance.config(
                bg=COLOR_AVISO, fg="white", font=("Segoe UI", 9, "bold"),
            )
        mostrar_cambio_hora()

    COLOR_FONDO_BOTON = btn_examinar_balance.cget("bg")

    proceso_anterior = {"valor": var_sel_proceso.get()}

    def proceso_elegido(*_):
        """El '+' de la lista: pide el nombre del proceso nuevo y lo guarda."""

        if var_sel_proceso.get() != OPCION_AGREGAR_PROCESO:
            proceso_anterior["valor"] = var_sel_proceso.get()
            return
        from tkinter import simpledialog

        texto = simpledialog.askstring(
            "Agregar proceso",
            "Nombre del proceso (por ejemplo R01P o R01D):",
            parent=root,
        )
        nuevo = limpiar_proceso(texto)
        if not nuevo:
            var_sel_proceso.set(proceso_anterior["valor"])
            return
        procesos = procesos_guardados()
        existente = next((p for p in procesos if p.lower() == nuevo.lower()), None)
        if existente is None:
            extra = [p for p in procesos if p not in PROCESOS_FIJOS] + [nuevo]
            guardar_config({CLAVE_PROCESOS: extra})
            procesos.append(nuevo)
            log(f"Proceso agregado a la lista: {nuevo}")
        cb_proceso.config(values=procesos + [OPCION_AGREGAR_PROCESO])
        var_sel_proceso.set(existente or nuevo)

    cb_proceso.bind("<<ComboboxSelected>>", proceso_elegido)
    for variable in (var_sel_ano, var_sel_mes, var_sel_proceso):
        variable.trace_add("write", refrescar_balance)
    lbl_ruta_balance.bind(
        "<Button-1>",
        lambda e: abrir_en_explorador(carpeta_del_balance(*seleccion()[:2]))
        if carpeta_del_balance(*seleccion()[:2]) else None,
    )
    refrescar_balance()

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

        if id_fila == "soc":
            return ("Traer", traer_soc)

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
    # EJECUTAR TODO Y ANTECEDENTES: debajo del arbol, justo arriba del
    # registro (pedido del usuario). Los comandos se ponen al final,
    # cuando ya estan definidos.
    # --------------------------------------------------------

    frame_acciones = tk.Frame(contenedor)
    frame_acciones.pack(fill="x", padx=20, pady=(8, 0))

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
        """Deja el registro en <mes>/BESS/Log/Registro_ultima_corrida.txt."""

        base = var_base.get()
        if not base or not os.path.isdir(base):
            return None
        # <mes>/BESS/Log (o <caso>/Log en un caso de antes de BESS/).
        rutas_caso = nucleo.resolver_rutas(base, var_aamm.get().strip())
        carpeta_log = str(rutas_caso["log_dir"])
        ruta = str(rutas_caso["registro"])
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
            pintar_arbol([])
            return

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
            return

        pintar_arbol(filas)

        faltan = [f for f in filas if f["estado"] == "falta"]

        if faltan:
            var_estado.set(f"Faltan {len(faltan)} entradas requeridas.")
        else:
            var_estado.set("Entradas completas.")

    def faltan_subcarpetas(base):
        """
        Las subcarpetas del caso que todavia no estan en 'base' (las de
        un caso de antes de BESS/ si es de antes: ver
        nucleo.subcarpetas_del_caso).
        """

        return nucleo.faltan_subcarpetas(base)

    def usar_carpeta(ruta, completar=True, parent=None):
        """
        Pasa a trabajar sobre esa carpeta como la del balance elegido en
        las casillas (año, mes y proceso): la anota como la de ese
        balance, la guarda en config.json y repinta el diagrama.

        completar=True ofrece ademas crear las subcarpetas que le
        falten: una carpeta recien elegida no tiene por que estar
        completa.
        """

        ruta = str(ruta)
        aamm, proceso, nombre = seleccion()

        var_base.set(ruta)
        var_aamm.set(aamm)
        var_proceso.set(proceso)

        # Una sola escritura de config.json: el balance abierto y la
        # carpeta de ese balance (una carpeta es de un solo balance).
        recordar_carpeta(
            clave_balance(aamm, proceso) if aamm else "", ruta,
            ademas={"carpeta_base": ruta, "aamm": aamm, "proceso": proceso},
        )
        log(f"Balance {nombre}: {ruta}")

        texto_cambio, hay_cambio = mostrar_cambio_hora()
        if hay_cambio:
            log(f"Periodo {aamm}: {texto_cambio}")

        revisar()
        refrescar_balance()

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

    def ir_a_balance():
        """Abre la carpeta que ya tiene el balance elegido."""

        if corriendo["activo"]:
            return
        aamm, proceso, nombre = seleccion()
        ruta = carpeta_del_balance(aamm, proceso)
        if not ruta:
            refrescar_balance()
            return
        usar_carpeta(ruta, completar=False)

    def examinar_balance():
        """
        Elige (o crea, desde el mismo dialogo) la carpeta del balance de
        las casillas. Si esa carpeta ya era la de otro balance, se
        pregunta antes: una carpeta es de un solo balance.
        """

        if corriendo["activo"]:
            return
        aamm, proceso, nombre = seleccion()
        if not aamm or not proceso:
            messagebox.showwarning(
                "Falta el balance", "Elegi el año, el mes y el proceso.",
            )
            return

        inicial = carpeta_del_balance(aamm, proceso) or var_base.get()
        if inicial and Path(inicial).is_dir() and inicial == var_base.get():
            inicial = str(Path(inicial).parent)
        if not inicial or not Path(inicial).is_dir():
            inicial = ""

        ruta = filedialog.askdirectory(
            title=f"Selecciona la carpeta del balance {nombre}",
            initialdir=inicial, parent=root,
        )
        if not ruta:
            return

        duena = balance_de_la_carpeta(ruta)
        if duena and duena != clave_balance(aamm, proceso) and not messagebox.askyesno(
            "Carpeta de otro balance",
            f"{ruta}\n\nya es la carpeta de {nombre_de_clave(duena)}.\n\n"
            f"¿Pasa a ser la de {nombre}? (cada balance tiene su propia "
            f"carpeta: {nombre_de_clave(duena)} queda sin carpeta)",
        ):
            return

        usar_carpeta(ruta)

    def actualizar_vista():
        """
        Vuelve a mirar las carpetas del caso y repinta el diagrama: un
        archivo pegado a mano en el explorador no aparece hasta que
        algo llame a revisar(). Mientras corre un proceso no hace nada:
        repintar el arbol recrea sus botones, que en ese momento estan
        deshabilitados a proposito.
        """

        if corriendo["activo"]:
            return

        revisar()
        refrescar_balance()
        log("Carpetas revisadas de nuevo.")

    btn_ir_balance.config(command=ir_a_balance)
    btn_examinar_balance.config(command=examinar_balance)

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
        "generar_soc": "Traer SoC desde PI",
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

    def traer_soc():
        """
        Trae desde PI AF (SRV-SCADA-AF2) el SoC de todas las centrales
        de 'Resumen BESS' y deja Medidas/SOC_AAMM.xlsx en tres
        columnas: central | fecha | soc. Las rutas salen de la columna
        SoC_PI del Diccionario (ver nucleo.generar_soc).
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
            nucleo.generar_soc,
            dict(carpeta_base=ruta, aamm=aamm),
            nucleo.nombre_archivo_soc(aamm),
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
    # BOTONES: Ejecutar todo y antecedentes arriba del registro; abajo
    # solo Actualizar estado y Salir.
    # --------------------------------------------------------

    tk.Button(
        frame_acciones,
        text="Ejecutar todo",
        font=("Segoe UI", 9, "bold"),
        bg="#fdf0d5",
        command=ventana_ejecutar_todo,
    ).pack(side="left", padx=(0, 10))

    tk.Button(
        frame_acciones,
        text="Generar antecedentes SAE",
        command=ventana_antecedentes,
    ).pack(side="left", padx=10)

    tk.Button(
        frame_botones, text="Actualizar estado", command=actualizar_vista,
    ).pack(side="left", padx=10, expand=True)

    tk.Button(frame_botones, text="Salir", command=root.destroy).pack(
        side="left", padx=10, expand=True
    )

    # --------------------------------------------------------
    # ARRANQUE
    # --------------------------------------------------------

    pintar_arbol([])
    log(
        "Elige el balance arriba (año _ mes _ proceso) y entra con 'Ir a "
        "este balance' o, si todavia no tiene carpeta, con 'Examinar'. "
        "Cada fila del diagrama de abajo trae su propio boton: "
        "'Traer' y 'Generar' en Cmg/, 'Traer' y un 'Generar' por cada "
        "FMA en 'FD y FMA/', 'Traer' en Subastas/DB subastas/, y en "
        f"{nucleo.ARCHIVO_SALIDA} un boton por hoja ('Actualizar' en "
        "las de entrada, 'Calcular' / 'Resumir' / 'Traer' / 'Asignar' "
        "en las de calculo). Cada archivo y cada carpeta del diagrama "
        "es un link a su ruta, y las filas que traen algo de afuera "
        "dicen 'Origen: ...' con un link a la carpeta de origen."
    )

    if var_base.get():
        ano, mes = ano_mes_de(var_aamm.get())
        log(
            f"Balance abierto: {nombre_balance(ano, mes, var_proceso.get()) or '(sin periodo)'}"
            f" - {var_base.get()}"
        )
        revisar()

    # El caso con que arranca la ventana es el que quedo de la ultima
    # vez: se anota como la carpeta de ESE balance (si se sabe su
    # proceso), para que "Ir a este balance" lo encuentre.
    if var_base.get() and var_proceso.get() and ano_mes_de(var_aamm.get())[0] \
            and Path(var_base.get()).is_dir() \
            and not carpeta_del_balance(var_aamm.get(), var_proceso.get()):
        recordar_carpeta(
            clave_balance(var_aamm.get(), var_proceso.get()), var_base.get(),
        )
    refrescar_balance()

    root.mainloop()


if __name__ == "__main__":
    main()
