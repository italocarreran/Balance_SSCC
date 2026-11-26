# -*- coding: utf-8 -*-
"""
Balance CRA (remuneracion CRA).

Ventana unica, con el MISMO formato que Balance_BESS.py (pedido del
usuario 2026-11-16 y 2026-11-26, "copia el formato de los bess [...] o
la ventana en general"): arriba el "Balance del caso" -año _ mes _
proceso, la ruta del balance, "Ir a este balance", "Examinar", "Bloquear
caso" y el cambio de hora-, al medio el diagrama del caso (Estructura | Estado | Accion | Detalle, con
un boton en cada fila que se puede generar y cada nombre como link a su
carpeta), el registro con colores, hora y tiempo, y abajo "Actualizar
estado" y "Salir". Colores,
anchos, rueda del mouse y registro salen de Script/ventana_estilo.py,
el mismo modulo que usa Balance_BESS.py.

La carpeta que se elige es la del MES, la misma que usa Balance_BESS.py.
Lo propio del CRA va en <mes>/CRA/ y lo comun con el BESS en la raiz del
mes (ver Script/nucleo/carpetas.py):

    <CARPETA DEL MES>/
        CRA/
            Auxiliares/
                centrales_cra.xlsx         (configuraciones del CRA)
            Energia/
                Formato_Solicitud_SSAA_SSCC_Hidro_*.xlsx
            FP/
                fp_<AAMM>_1_<N>.xlsx       [Generar]
            CO/
                cvar_cra_<AAMM>_1_<N>.xlsx [Generar]
            SC y CO/
                Reporte_CRA*.csv
                Cálculo_SobrecostosSSCC_*.xlsm (hoja SOBRECOSTOS)
            Log/
                Registro_ultima_corrida.txt
            5_REMUNERACION_CRA_<AAMM>.xlsx [Actualizar]
                hoja 'PRORRATA_RETIROS'    [Actualizar]  (solo bloque 1)
                hoja 'FD_CTF' / 'FD_CSF' / 'FD_CPF'  [Actualizar]
                hoja 'SC y CO' / 'CO' / 'FP' / 'ENERGIA'  [Actualizar]
        BESS/                              (Balance_BESS.py)
        Cmg/                               (comun con el BESS; para el TC)
        FD y FMA/                          (comun con el BESS)
            SSCC_Desempeño_*.xlsx          [Traer]
        Prorrata retiros/                  (comun con el BESS)
            Prorrata_Retiros_<AAMM>_pre/_def.xlsx

Un caso de antes (todo suelto en la carpeta, con su propia FD/) se
sigue leyendo tal cual.

Cada balance (mes + proceso) tiene SU carpeta, como en el BESS (las
funciones de año/mes/proceso son las de Script/balance_caso.py). "Ir a
este balance" busca la que anoto el CRA y, si no, la que anoto el BESS
para ese mismo balance (es la misma carpeta del mes). "Bloquear caso"
bloquea solo el CRA (archivo en <mes>/CRA/; ver nucleo/bloqueo.py).

Por ahora solo estan las hojas de entrada con su carga definida en
docs/Trazabilidad_CRA_Periodo_Generico_v5_Auditoria_Formulas.md; el
resto se muestra como PENDIENTE. Todo el calculo vive en Script/cra:
esta ventana solo lo llama.
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
from Script.balance_caso import (
    ANO_INICIAL, CLAVE_PROCESOS, MESES, OPCION_AGREGAR_PROCESO,
    PROCESOS_FIJOS, aamm_de, ano_mes_de, anos_disponibles, clave_balance,
    limpiar_proceso, nombre_balance, nombre_de_clave, proceso_en_el_nombre,
)
from Script.dialogos import mostrar_error
from Script.ventana_estilo import (
    ALTO_ACCION, ALTO_FILA, ANCHO_ACCION, ANCHO_ESTADO, ANCHO_ESTRUCTURA,
    COLOR_AVISO, COLOR_BASE_OK, COLOR_ESTADO, COLOR_FALTA, COLOR_LINK,
    COLOR_NEUTRO, COLOR_REGISTRO, EVENTOS_RUEDA, LINEAS_POR_MUESCA,
    MENSAJES_POR_TANDA, PIXELES_POR_MUESCA, RELLENO_ACCION, SIMBOLO,
    abrir_en_explorador, formatear_reloj, formato_tiempo, nivel_del_mensaje,
)

# Script.cra (y con el, pandas/openpyxl y nucleo) se carga con la
# ventana ya a la vista, como en Balance_BESS.py.
cra = None
cambio_hora = None
orquestador = None      # marcas del registro y formato de duracion del BESS
bloqueo_nucleo = None   # Bloquear/Desbloquear caso (nucleo/bloqueo.py)

TEXTO_CARGANDO = "Cargando…"


def _cargar():
    global cra, cambio_hora, orquestador, bloqueo_nucleo
    if cra is None:
        from Script.nucleo import bloqueo as _bloqueo
        bloqueo_nucleo = _bloqueo
        from Script import cra as _cra
        from Script.nucleo import cambio_hora as _cambio_hora
        from Script.nucleo import orquestador as _orquestador
        cra, cambio_hora = _cra, _cambio_hora
        orquestador = _orquestador
    return cra


def _ventana_viva(root):
    try:
        return bool(root.winfo_exists())
    except tk.TclError:
        return False


def _cargar_con_la_ventana_viva(root):
    """Carga Script.cra en un hilo; False si cerraron la ventana o fallo."""

    if cra is not None:
        return True

    resultado = {}

    def cargar():
        try:
            _cargar()
        except BaseException as error:
            resultado["error"] = error

    hilo = threading.Thread(target=cargar, daemon=True)
    hilo.start()
    while hilo.is_alive():
        try:
            root.update()
        except tk.TclError:
            return False
        if not _ventana_viva(root):
            return False
        time.sleep(0.01)

    if "error" in resultado:
        mostrar_error(
            "No se pudo cargar el programa",
            "".join(traceback.format_exception(resultado["error"])),
            parent=root,
        )
        root.destroy()
        return False
    return _ventana_viva(root)


# ============================================================
# CONFIG POR PC/USUARIO
#
# Seccion propia ("modo_cra" en __config__/config.json): no pisa lo que
# recuerda Balance_BESS.py ("modo_bess"). Lo que si se LEE del BESS es
# la carpeta de cada periodo: es la misma carpeta del mes.
# ============================================================

CLAVE_CARPETAS_PERIODO = "carpetas_por_periodo"


def get_usuario():
    usuario = (
        os.environ.get("USERNAME")
        or os.environ.get("USER")
        or "desconocido"
    )
    return f"{socket.gethostname()}_{usuario}"


def leer_config():
    # La primera vez se copia la seccion "<equipo>_<usuario>_CRA" de antes.
    return config.seccion_usuario(
        config.MODO_CRA, get_usuario(), clave_antigua=f"{get_usuario()}_CRA"
    )


def guardar_config(data):
    config.actualizar_seccion_usuario(config.MODO_CRA, get_usuario(), data)


def _carpetas_de(bloque):
    datos = bloque.get(CLAVE_CARPETAS_PERIODO) if isinstance(bloque, dict) else None
    if not isinstance(datos, dict):
        return {}
    return {
        str(periodo): str(ruta) for periodo, ruta in datos.items()
        if isinstance(ruta, str) and ruta
    }


def carpetas_por_periodo():
    """{aamm: carpeta} que recuerda el CRA en este PC/usuario."""

    return _carpetas_de(leer_config())


def carpetas_del_bess():
    """
    {aamm: carpeta} que recuerda Balance_BESS.py (solo se lee). Desde
    2026-11-18 el BESS las anota por balance ("2603_Definitivo"): aca se
    vuelven al AAMM, que es lo unico que usa el CRA. Si el mes tiene mas
    de un proceso, gana el que se anoto ultimo.
    """

    try:
        datos = _carpetas_de(config.seccion(config.MODO_BESS).get(get_usuario()))
    except Exception:
        return {}
    return {str(clave).split("_", 1)[0]: ruta for clave, ruta in datos.items()}


def carpetas_del_bess_por_balance():
    """{"2603_Definitivo": carpeta} tal como los anota Balance_BESS.py."""

    try:
        return _carpetas_de(config.seccion(config.MODO_BESS).get(get_usuario()))
    except Exception:
        return {}


def procesos_guardados():
    """Preliminar, Definitivo y los que se agregaron (aca o en el BESS)."""

    procesos = list(PROCESOS_FIJOS)
    extras = list(leer_config().get(CLAVE_PROCESOS) or [])
    try:
        bess = config.seccion(config.MODO_BESS).get(get_usuario()) or {}
        extras += list(bess.get(CLAVE_PROCESOS) or [])
    except Exception:
        pass
    for proceso in extras:
        proceso = limpiar_proceso(proceso)
        if proceso and proceso.lower() not in (p.lower() for p in procesos):
            procesos.append(proceso)
    return procesos


def _existe(ruta):
    try:
        return bool(ruta) and Path(ruta).is_dir()
    except OSError:
        return False


def carpeta_del_balance(aamm, proceso):
    """
    La carpeta de ese balance si ya tiene una que exista; '' si no. Es la
    misma carpeta del mes que la del BESS: se busca primero en lo que
    anoto el CRA y despues en lo que anoto el BESS (por clave
    "AAMM_Proceso"). Una anotada solo con el AAMM (de antes) sirve si su
    nombre dice ese proceso.
    """

    clave = clave_balance(aamm, proceso)
    for carpetas in (carpetas_por_periodo(), carpetas_del_bess_por_balance()):
        ruta = carpetas.get(clave, "")
        if not ruta and proceso:
            vieja = carpetas.get(str(aamm), "")
            if vieja and proceso.lower() in Path(vieja).name.lower():
                ruta = vieja
        if _existe(ruta):
            return ruta
    return ""


def balance_de_la_carpeta(ruta):
    """La clave del balance que ya usa esa carpeta en el CRA, o ''."""

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


def recordar_carpeta(clave, ruta, ademas=None):
    """'ruta' es la carpeta del balance 'clave' (y de ningun otro)."""

    datos = dict(ademas or {})
    if clave and ruta:
        ruta = str(ruta)
        actuales = carpetas_por_periodo()
        carpetas = {
            otra_clave: otra for otra_clave, otra in actuales.items()
            if otra_clave == str(clave) or otra != ruta
        }
        carpetas[str(clave)] = ruta
        if carpetas != actuales:
            datos[CLAVE_CARPETAS_PERIODO] = carpetas
    if datos:
        guardar_config(datos)


# ============================================================
# VENTANA
# ============================================================

def main():
    root = tk.Tk()
    root.title("Balance CRA")
    root.geometry("1180x820")
    cartel = tk.Label(
        root, text=TEXTO_CARGANDO, font=("Segoe UI", 11), fg=COLOR_NEUTRO,
    )
    cartel.pack(expand=True)
    try:
        root.update()
    except tk.TclError:
        return
    if not _cargar_con_la_ventana_viva(root):
        return
    cartel.destroy()

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

    botones_arbol = []
    corriendo = {"activo": False}

    # --------------------------------------------------------
    # BOTONES FIJOS ABAJO (primero, para que no los tape nada)
    # --------------------------------------------------------

    frame_botones = tk.Frame(root)
    frame_botones.pack(side="bottom", fill="x", pady=10)

    # --------------------------------------------------------
    # CANVAS CON SCROLL (igual que Balance_BESS.py)
    # --------------------------------------------------------

    canvas = tk.Canvas(
        root, borderwidth=0, highlightthickness=0, yscrollincrement=1,
    )
    scroll = tk.Scrollbar(root, orient="vertical", command=canvas.yview)
    canvas.configure(yscrollcommand=scroll.set)
    scroll.pack(side="right", fill="y")
    canvas.pack(side="left", fill="both", expand=True)

    contenedor = tk.Frame(canvas)
    ventana_canvas = canvas.create_window((0, 0), window=contenedor, anchor="nw")
    ancho_pintado = {}

    def ajustar_ancho(event):
        if ancho_pintado.get("ancho") != event.width:
            ancho_pintado["ancho"] = event.width
            canvas.itemconfig(ventana_canvas, width=event.width)

    contenedor.bind(
        "<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
    )
    canvas.bind("<Configure>", ajustar_ancho)

    # --------------------------------------------------------
    # RUEDA DEL MOUSE: se desplaza lo que esta debajo del puntero (el
    # registro o la ventana) y el resto de los deltas chicos de un
    # touchpad se acumula (ver Balance_BESS.py).
    # --------------------------------------------------------

    desplazables = {}
    resto_rueda = {"ventana": 0.0, "registro": 0.0}

    def _acumular(clave, cantidad):
        resto = resto_rueda.setdefault(clave, 0.0)
        if resto and (resto > 0) != (cantidad > 0):
            resto = 0.0
        total = cantidad + resto
        entero = int(total)
        resto_rueda[clave] = total - entero
        return entero

    def _muescas(event):
        if getattr(event, "num", None) in (4, 5):
            return -1.0 if event.num == 4 else 1.0
        delta = getattr(event, "delta", 0)
        if not delta:
            return 0.0
        if sys.platform == "darwin":
            return -float(delta)
        return -delta / 120.0

    def _desplazar_ventana(muescas):
        region = canvas.bbox("all")
        if not region or region[3] - region[1] <= canvas.winfo_height():
            return
        pixeles = _acumular("ventana", muescas * PIXELES_POR_MUESCA)
        if pixeles:
            canvas.yview_scroll(pixeles, "units")

    def rueda(event):
        muescas = _muescas(event)
        if not muescas:
            return "break"
        try:
            widget = root.winfo_containing(event.x_root, event.y_root)
        except tk.TclError:
            widget = None
        widget = widget if widget is not None else event.widget
        inicio = widget
        while widget is not None:
            desplazar = desplazables.get(str(widget))
            if desplazar is not None:
                desplazar(muescas)
                return "break"
            widget = getattr(widget, "master", None)
        try:
            if inicio is not None and inicio.winfo_toplevel() is not root:
                return "break"
        except (tk.TclError, AttributeError):
            pass
        _desplazar_ventana(muescas)
        return "break"

    for nombre in EVENTOS_RUEDA:
        root.bind_all(nombre, rueda)

    # --------------------------------------------------------
    # EL BALANCE DEL CASO: AÑO _ MES _ PROCESO (igual que Balance_BESS.py)
    #
    # Tres casillas que se leen "2026_Marzo_Definitivo", la ruta del
    # balance y los botones "Ir a este balance" (azul si el balance ya
    # tiene carpeta, bloqueado si no), "Examinar" (destacado cuando no la
    # tiene) y "Bloquear caso". Elegir en las casillas NO cambia el caso
    # abierto: recien lo cambian los botones. var_aamm / var_proceso son
    # los del caso ABIERTO.
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

    # Bloquear/Desbloquear el CRA del caso abierto: su propio archivo en
    # <mes>/CRA/ (ver nucleo/bloqueo.py); no bloquea al BESS.
    btn_bloqueo = tk.Button(
        fila_botones_balance, text="Bloquear caso", width=16,
    )
    btn_bloqueo.pack(side="left", padx=(16, 0))
    bloqueo = {"info": None}

    var_cambio_hora = tk.StringVar(value="")
    label_cambio_hora = tk.Label(
        frame_periodo, textvariable=var_cambio_hora,
        font=("Segoe UI", 9, "bold"), fg=COLOR_NEUTRO, wraplength=900,
        justify="left", anchor="w",
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
            cra.validar_aamm(aamm)
            texto = cambio_hora.describir_periodo(aamm)
            hay = bool(cambio_hora.cambios_del_periodo(aamm))
        except (cra.ErrorEntrada, ValueError):
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
            propios = [
                limpiar_proceso(p) for p in (leer_config().get(CLAVE_PROCESOS) or [])
            ]
            guardar_config({CLAVE_PROCESOS: [p for p in propios if p] + [nuevo]})
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

    # --------------------------------------------------------
    # DIAGRAMA DE LA ESTRUCTURA DEL CASO
    # --------------------------------------------------------

    frame_arbol = tk.LabelFrame(
        contenedor, text="Estructura del caso", padx=8, pady=8
    )
    frame_arbol.pack(fill="both", expand=True, padx=20, pady=6)

    enc_arbol = tk.Frame(frame_arbol)
    enc_arbol.pack(fill="x")
    for texto, ancho in (
        ("Estructura", ANCHO_ESTRUCTURA), ("Estado", ANCHO_ESTADO),
        ("Accion", ANCHO_ACCION),
    ):
        celda = tk.Frame(enc_arbol, width=ancho, height=16)
        celda.pack(side="left")
        celda.pack_propagate(False)
        tk.Label(
            celda, text=texto, font=("Segoe UI", 8, "bold"), anchor="w",
        ).pack(side="left")
    tk.Label(
        enc_arbol, text="Detalle", font=("Segoe UI", 8, "bold"), anchor="w",
    ).pack(side="left")

    filas_arbol = tk.Frame(frame_arbol)
    filas_arbol.pack(fill="x")

    def _celda(parent, ancho, alto):
        celda = tk.Frame(parent, width=ancho, height=alto)
        celda.pack(side="left")
        celda.pack_propagate(False)
        return celda

    def abrir_ruta_de_fila(ruta, es_carpeta):
        if abrir_en_explorador(ruta, es_archivo=not es_carpeta) is None:
            log(f"No se pudo abrir {ruta}")

    def _fila_arbol(prefijo, fila, boton):
        """Una fila: estructura (link), estado, boton y detalle."""

        negrita = fila["nivel"] == 0
        estilo = "bold" if negrita else "normal"
        ruta = fila.get("ruta")
        es_carpeta = fila.get("es_carpeta", False)
        alto = ALTO_ACCION if boton is not None else ALTO_FILA

        renglon = tk.Frame(filas_arbol)
        renglon.pack(fill="x", pady=1)

        celda_estructura = _celda(renglon, ANCHO_ESTRUCTURA, alto)
        if prefijo:
            tk.Label(
                celda_estructura, text=prefijo, font=("Consolas", 9, estilo),
                anchor="w",
            ).pack(side="left")
        etiqueta = tk.Label(
            celda_estructura, text=fila["texto"],
            font=("Consolas", 9, f"{estilo} underline" if ruta else estilo),
            anchor="w", fg=COLOR_LINK if ruta else "black",
            cursor="hand2" if ruta else "",
        )
        etiqueta.pack(side="left")
        if ruta:
            etiqueta.bind(
                "<Button-1>",
                lambda e, r=ruta, c=es_carpeta: abrir_ruta_de_fila(r, c),
            )

        estado = fila["estado"]
        tk.Label(
            _celda(renglon, ANCHO_ESTADO, alto),
            text=SIMBOLO.get(estado, estado), anchor="w",
            fg=COLOR_ESTADO.get(estado, COLOR_NEUTRO),
            font=("Segoe UI", 9, "bold"),
        ).pack(side="left")

        celda_accion = _celda(renglon, ANCHO_ACCION, alto)
        if boton is not None:
            texto_boton, comando = boton
            widget = tk.Button(
                celda_accion, text=texto_boton, font=("Segoe UI", 8, "bold"),
                bg="#fdf0d5", command=comando,
            )
            widget.pack(
                fill="both", expand=True,
                padx=(0, RELLENO_ACCION * 2), pady=RELLENO_ACCION // 2,
            )
            botones_arbol.append(widget)

        if fila.get("detalle"):
            tk.Label(
                renglon, text=fila["detalle"], anchor="w", fg=COLOR_NEUTRO,
                font=("Segoe UI", 8),
            ).pack(side="left")

    def _boton_de_fila(fila):
        """(texto, comando) del boton de esa fila, o None."""

        id_fila = fila["id"]
        if id_fila == "salida":
            return ("Actualizar", lambda: lanzar(
                cra.generar_balance_cra, {}, "planilla completa"))
        if id_fila == "energia":
            # Medidas_CRA.xlsx desde PRMTE: replica el Formato_Solicitud.
            return ("Generar", lambda: lanzar(
                cra.generar_medidas_cra, {}, "Medidas_CRA.xlsx en CRA/Energia/"))
        if id_fila == "sscc_desempeno":
            return ("Traer", lambda: lanzar(
                cra.traer_fd, {}, "SSCC_Desempeño_* en FD y FMA/"))
        if id_fila == "fp":
            return ("Generar", lambda: lanzar(
                cra.generar_politicas, {"que": ["fp"]}, "fp_<AAMM>"))
        if id_fila == "co":
            return ("Generar", lambda: lanzar(
                cra.generar_politicas, {"que": ["cvar"]}, "cvar_cra_<AAMM>"))
        if id_fila.startswith("hoja_"):
            seccion = id_fila[len("hoja_"):]
            if seccion in cra.SECCIONES:
                hoja = cra.SECCIONES[seccion][0]
                return ("Actualizar", lambda s=seccion, h=hoja: lanzar(
                    cra.generar_balance_cra, {"secciones": [s]},
                    f"hoja '{h}'"))
        return None

    def pintar_arbol(filas):
        for hijo in filas_arbol.winfo_children():
            hijo.destroy()
        botones_arbol.clear()

        if not filas:
            tk.Label(
                filas_arbol,
                text="Elegi el balance arriba (Ir a este balance o Examinar).",
                fg=COLOR_NEUTRO, anchor="w",
            ).pack(fill="x")
            return

        prefijos = _prefijos_arbol([fila["nivel"] for fila in filas])
        for fila, prefijo in zip(filas, prefijos):
            _fila_arbol(prefijo, fila, _boton_de_fila(fila))

        habilitar_botones(not corriendo["activo"])

    def habilitar_botones(activos):
        # Con el caso bloqueado no se corre nada (nucleo/bloqueo.py).
        activos = activos and bloqueo["info"] is None
        for widget in botones_arbol:
            try:
                widget.config(state="normal" if activos else "disabled")
            except tk.TclError:
                pass

    # --------------------------------------------------------
    # ESTADO, TIEMPO Y REGISTRO
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

    frame_log = tk.LabelFrame(contenedor, text="Registro", padx=6, pady=6)
    frame_log.pack(fill="both", expand=True, padx=20, pady=6)

    txt_log = tk.Text(frame_log, height=14, font=("Consolas", 9), wrap="word")
    scroll_log = tk.Scrollbar(frame_log, command=txt_log.yview)
    txt_log.configure(yscrollcommand=scroll_log.set)
    scroll_log.pack(side="right", fill="y")
    txt_log.pack(side="left", fill="both", expand=True)

    def _desplazar_registro(muescas):
        lineas = _acumular("registro", muescas * LINEAS_POR_MUESCA)
        if lineas:
            txt_log.yview_scroll(lineas, "units")

    for widget in (frame_log, txt_log, scroll_log):
        desplazables[str(widget)] = _desplazar_registro
        for nombre in EVENTOS_RUEDA:
            widget.bind(nombre, rueda)

    for nivel, color in COLOR_REGISTRO.items():
        txt_log.tag_configure(nivel, foreground=color)
    txt_log.tag_configure("hora", foreground=COLOR_NEUTRO)
    txt_log.tag_configure(
        "encabezado", background="#244062", foreground="white",
        font=("Segoe UI", 9, "bold"), spacing1=6, spacing3=2,
    )
    txt_log.tag_configure(
        "encabezado_fin", foreground="#375623",
        font=("Segoe UI", 9, "bold"), spacing3=4,
    )

    inicio_log = {"t": None}

    def _insertar_log(mensaje, ahora=None):
        texto = str(mensaje)
        nivel = nivel_del_mensaje(texto)
        ahora = time.time() if ahora is None else ahora
        for linea in texto.split("\n"):
            if linea.startswith(orquestador.ENCABEZADO_REGISTRO):
                txt_log.insert("end", f" {linea} \n", ("encabezado",))
                continue
            if linea.startswith(orquestador.FIN_REGISTRO):
                txt_log.insert("end", f"{linea}\n", ("encabezado_fin",))
                continue
            if linea.strip():
                marca = time.strftime("%H:%M:%S", time.localtime(ahora))
                if inicio_log["t"] is not None:
                    marca += f" +{formatear_reloj(ahora - inicio_log['t'])}"
                txt_log.insert("end", f"[{marca}] ", ("hora",))
            txt_log.insert("end", linea + "\n", (nivel,) if nivel else ())

    def log(mensaje):
        _insertar_log(mensaje)
        txt_log.see("end")

    cola_log = queue.Queue()
    sondeo_log = {"programado": None}

    def vaciar_cola_log(todo=False):
        if sondeo_log["programado"] is not None:
            try:
                root.after_cancel(sondeo_log["programado"])
            except tk.TclError:
                pass
            sondeo_log["programado"] = None
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
        """Deja el registro en <mes>/CRA/Log/Registro_ultima_corrida.txt."""

        base = var_base.get()
        if not base or not os.path.isdir(base):
            return None
        try:
            rutas = cra.resolver_rutas(base, var_aamm.get().strip())
            rutas["log_dir"].mkdir(parents=True, exist_ok=True)
            rutas["registro"].write_text(
                txt_log.get("1.0", "end"), encoding="utf-8"
            )
        except OSError:
            return None
        return rutas["registro"]

    timer = {"corriendo": False, "inicio": 0.0}

    def tick():
        if timer["corriendo"]:
            var_tiempo.set(formato_tiempo(time.time() - timer["inicio"]))
            root.after(500, tick)

    # --------------------------------------------------------
    # REVISION DE LA CARPETA
    # --------------------------------------------------------

    def raiz_bloqueo(base):
        return cra.raiz_cra(base)

    def actualizar_bloqueo():
        """Lee el archivo de bloqueo del CRA del caso abierto y pinta el boton."""

        base = var_base.get()
        hay_caso = bool(base) and Path(base).is_dir()
        bloqueo["info"] = (
            bloqueo_nucleo.info_bloqueo(base, raiz=raiz_bloqueo(base))
            if hay_caso else None
        )
        if bloqueo["info"] is not None:
            btn_bloqueo.config(text="Desbloquear caso", bg="#b00020", fg="white",
                               font=("Segoe UI", 9, "bold"))
        else:
            btn_bloqueo.config(text="Bloquear caso", bg=COLOR_FONDO_BOTON, fg="black",
                               font=("Segoe UI", 9))
        btn_bloqueo.config(
            state="normal" if hay_caso and not corriendo["activo"] else "disabled"
        )
        habilitar_botones(not corriendo["activo"])

    def revisar(*_):
        ruta = var_base.get()
        actualizar_bloqueo()

        if not ruta or not Path(ruta).is_dir():
            pintar_arbol([])
            return

        try:
            filas = cra.revisar_estructura(ruta, var_aamm.get().strip())
        except Exception as error:
            log(f"Error al revisar la carpeta: {error}")
            pintar_arbol([{
                "id": "error", "texto": "Error al revisar (ver el registro)",
                "nivel": 0, "estado": "falta", "detalle": "", "ruta": None,
            }])
            return

        pintar_arbol(filas)
        faltan = [f for f in filas if f["estado"] == "falta"]
        var_estado.set(
            f"Faltan {len(faltan)} entradas requeridas." if faltan
            else "Entradas completas."
        )

    def crear_carpetas_en(base):
        try:
            Path(base).mkdir(parents=True, exist_ok=True)
            creadas = cra.crear_carpetas_caso(base)
        except OSError as error:
            mostrar_error("No se pudo crear", str(error))
            return False
        for sub in creadas:
            log(f"  creada: {Path(base) / sub}")
        log(
            f"Caso preparado en {base} ({len(creadas)} carpeta(s) nuevas)."
            if creadas else f"{base} ya tenia todas sus carpetas."
        )
        revisar()
        return True

    def usar_carpeta(ruta, completar=True):
        """
        Pasa a trabajar sobre esa carpeta como la del balance elegido en
        las casillas: la anota como la de ese balance y repinta. Con
        completar=True ofrece crear las subcarpetas que le falten.
        """

        ruta = str(ruta)
        aamm, proceso, nombre = seleccion()
        var_base.set(ruta)
        var_aamm.set(aamm)
        var_proceso.set(proceso)
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
        faltan = [
            sub for sub in cra.subcarpetas_del_caso(ruta)
            if not (Path(ruta) / sub).is_dir()
        ]
        if faltan and messagebox.askyesno(
            "Faltan carpetas del caso",
            f"A {Path(ruta).name} le faltan {len(faltan)} carpeta(s) del caso:"
            f"\n\n" + "\n".join(f"  {sub}/" for sub in faltan)
            + "\n\n¿Las creo ahora? (vacias; no se toca nada de lo que ya hay)",
            parent=root,
        ):
            crear_carpetas_en(ruta)

    def ir_a_balance():
        """Abre la carpeta que ya tiene el balance elegido."""

        if corriendo["activo"]:
            return
        aamm, proceso, _ = seleccion()
        ruta = carpeta_del_balance(aamm, proceso)
        if not ruta:
            refrescar_balance()
            return
        usar_carpeta(ruta, completar=False)

    def examinar_balance():
        """
        Elige (o crea, desde el mismo dialogo) la carpeta del balance de
        las casillas. Es la misma carpeta del mes que usa el BESS. Si ya
        era la de otro balance en el CRA, se pregunta antes.
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
        if corriendo["activo"]:
            return
        revisar()
        refrescar_balance()
        log("Carpetas revisadas de nuevo.")

    def alternar_bloqueo():
        """Bloquea o desbloquea el CRA del caso abierto (pregunta antes)."""

        base = var_base.get()
        if corriendo["activo"] or not base or not Path(base).is_dir():
            return
        ano, mes = ano_mes_de(var_aamm.get())
        nombre = nombre_balance(ano, mes, var_proceso.get()) or Path(base).name
        raiz = raiz_bloqueo(base)
        actualizar_bloqueo()
        try:
            if bloqueo["info"] is None:
                if not messagebox.askyesno(
                    "Bloquear caso",
                    f"¿Bloquear el CRA de {nombre}?\n\n{base}\n\nMientras "
                    f"este bloqueado no se corre ningun proceso del CRA. Queda "
                    f"un archivo {bloqueo_nucleo.ARCHIVO_BLOQUEO} en "
                    f"{raiz.name}/: lo ve cualquiera que abra la carpeta. El "
                    f"BESS tiene su propio bloqueo.",
                ):
                    return
                bloqueo_nucleo.bloquear(
                    base, quien=get_usuario(), raiz=raiz, herramienta="Balance CRA",
                )
                log(f"Caso bloqueado: {nombre}")
            else:
                if not messagebox.askyesno(
                    "Desbloquear caso",
                    f"¿Desbloquear el CRA de {nombre}?\n\n{base}\n\nSe van a "
                    f"poder volver a correr sus procesos y cambiar sus resultados.",
                    icon="warning",
                ):
                    return
                bloqueo_nucleo.desbloquear(base, raiz=raiz)
                log(f"Caso desbloqueado: {nombre}")
        except OSError as error:
            mostrar_error("No se pudo cambiar el bloqueo", str(error))
        actualizar_bloqueo()

    btn_ir_balance.config(command=ir_a_balance)
    btn_examinar_balance.config(command=examinar_balance)
    btn_bloqueo.config(command=alternar_bloqueo)

    # Un clic fuera de un campo de texto le quita el foco (y aplica el
    # AAMM escrito), como en el BESS.
    campos_de_texto = (tk.Entry, tk.Text, tk.Spinbox, ttk.Entry)

    def soltar_foco(evento):
        widget = evento.widget
        if isinstance(widget, str) or isinstance(widget, campos_de_texto):
            return
        try:
            enfocado = widget.focus_get()
        except (KeyError, tk.TclError):
            return
        if isinstance(enfocado, (tk.Entry, tk.Spinbox, ttk.Entry)):
            widget.winfo_toplevel().focus_set()

    root.bind_all("<Button-1>", soltar_foco, add="+")

    # --------------------------------------------------------
    # LANZADOR (hilo aparte; registro y tiempo de la ventana)
    # --------------------------------------------------------

    TITULO_DE_FUNCION = {
        "generar_balance_cra": "Planilla 5_REMUNERACION_CRA",
        "traer_fd": "Traer FD (SSCC_Desempeño)",
        "generar_politicas": "Generar politicas",
        "generar_medidas_cra": "Generar Medidas_CRA (PRMTE)",
    }

    def lanzar(funcion, kwargs, etiqueta_salida):
        if corriendo["activo"]:
            return

        base, aamm = var_base.get().strip(), var_aamm.get().strip()
        if not base or not Path(base).is_dir():
            messagebox.showwarning(
                "Falta el balance",
                "Elegi primero el balance arriba (Ir a este balance o Examinar).",
            )
            return
        try:
            cra.validar_aamm(aamm)
        except cra.ErrorEntrada as error:
            messagebox.showwarning("Falta el periodo", str(error))
            return
        actualizar_bloqueo()
        if bloqueo["info"] is not None:
            messagebox.showwarning(
                "Caso bloqueado",
                "El CRA de este caso esta bloqueado: no se corre ningun "
                "proceso.\n\nSi de verdad hay que correrlo, usa 'Desbloquear "
                "caso' (arriba, junto a Examinar).",
            )
            return

        corriendo["activo"] = True
        habilitar_botones(False)
        txt_log.delete("1.0", "end")

        titulo = TITULO_DE_FUNCION.get(funcion.__name__, funcion.__name__)
        if kwargs.get("secciones"):
            titulo += ": " + ", ".join(kwargs["secciones"])
        if kwargs.get("que"):
            titulo += ": " + ", ".join(kwargs["que"])
        timer["corriendo"] = True
        timer["inicio"] = time.time()
        inicio_log["t"] = timer["inicio"]
        log(f"{orquestador.ENCABEZADO_REGISTRO}{titulo}")
        log(f"Inicio: {time.strftime('%d-%m-%Y %H:%M:%S')}")
        var_estado.set("Procesando...")
        tick()
        vaciar_cola_log()

        resultado = {"ok": False}

        def trabajo():
            try:
                funcion(
                    base, aamm, **kwargs,
                    registrar=registrar_desde_hilo,
                    progreso=lambda _v: None,
                )
                resultado["ok"] = True
            except cra.ErrorEntrada as error:
                registrar_desde_hilo(f"\nERROR DE ENTRADA:\n{error}")
                resultado["error"] = str(error)
            except Exception as error:
                registrar_desde_hilo(
                    f"\nERROR INESPERADO:\n{error}\n\n{traceback.format_exc()}"
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
                f"{orquestador.FIN_REGISTRO}{titulo}: "
                f"{'listo' if res['ok'] else 'termino con problemas'} en "
                f"{orquestador.formatear_duracion(duracion)}"
            )
            guardado = guardar_registro()
            if guardado:
                log(f"Registro guardado en {guardado}")
            inicio_log["t"] = None
            revisar()
            if res["ok"]:
                var_estado.set("Terminado correctamente.")
                messagebox.showinfo(
                    "Listo", f"Se genero/actualizo:\n{etiqueta_salida}"
                )
            else:
                var_estado.set("Termino con errores.")
                mostrar_error("Error", res.get("error", "Revisa el registro."))

        threading.Thread(target=trabajo, daemon=True).start()

    # --------------------------------------------------------
    # BOTONES FIJOS ABAJO (como en el BESS: Actualizar estado y Salir)
    # --------------------------------------------------------

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
    actualizar_bloqueo()      # el boton de bloqueo con su letra desde el inicio
    log(
        "Elige el balance arriba (año _ mes _ proceso) y entra con 'Ir a "
        "este balance' o, si todavia no tiene carpeta, con 'Examinar' (es "
        "la misma carpeta del mes que usa el BESS). Cada fila del diagrama "
        "trae su propio boton: 'Generar' en el Formato_Solicitud (arma "
        "Medidas_CRA.xlsx desde PRMTE), en FP/ y en CO/, 'Traer' en el "
        "SSCC_Desempeño de FD y FMA/, y 'Actualizar' en la planilla (todas "
        "las hojas) y en cada hoja. Cada archivo y cada carpeta del "
        "diagrama es un link a su ruta."
    )

    if var_base.get():
        ano, mes = ano_mes_de(var_aamm.get())
        log(
            f"Balance abierto: {nombre_balance(ano, mes, var_proceso.get()) or '(sin periodo)'}"
            f" - {var_base.get()}"
        )
        revisar()

    # El caso con que arranca la ventana es el de la ultima vez: se anota
    # como la carpeta de ESE balance, para que "Ir a este balance" lo
    # encuentre.
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
