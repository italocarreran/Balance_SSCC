# -*- coding: utf-8 -*-
"""
Balance CRA (remuneracion CRA).

Ventana unica, con el MISMO formato que Balance_BESS.py (pedido del
usuario 2026-11-16: "la ventana del CRA dejala en el formato del BESS"):
arriba la carpeta del mes y el periodo (AAMM) con su cambio de hora, al
medio el diagrama del caso (Estructura | Estado | Accion | Detalle, con
un boton en cada fila que se puede generar y cada nombre como link a su
carpeta), y abajo el registro con colores, hora y tiempo. Colores,
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

Cada periodo trabaja en SU carpeta, como en el BESS: al cambiar el AAMM
se vuelve sola a la carpeta de ese periodo (la que recuerda el CRA o, si
no, la que recuerda el BESS: es la misma carpeta del mes); si no hay
ninguna, se pregunta (examinar / crear / "esta es la de AAMM").

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
rutas_nucleo = None
cambio_hora = None
orquestador = None      # marcas del registro y formato de duracion del BESS

TEXTO_CARGANDO = "Cargando…"


def _cargar():
    global cra, rutas_nucleo, cambio_hora, orquestador
    if cra is None:
        from Script import cra as _cra
        from Script.nucleo import cambio_hora as _cambio_hora
        from Script.nucleo import orquestador as _orquestador
        from Script.nucleo import rutas as _rutas
        cra, rutas_nucleo, cambio_hora = _cra, _rutas, _cambio_hora
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


def carpeta_recordada(aamm):
    """La carpeta del periodo (la del CRA o, si no, la del BESS), si existe."""

    for recordadas in (carpetas_por_periodo(), carpetas_del_bess()):
        ruta = recordadas.get(str(aamm), "")
        try:
            if ruta and Path(ruta).is_dir():
                return ruta
        except OSError:
            continue
    return ""


def recordar_carpeta(aamm, ruta, ademas=None):
    """'ruta' es la carpeta de 'aamm' (y de ningun otro periodo)."""

    datos = dict(ademas or {})
    if aamm and ruta:
        ruta = str(ruta)
        actuales = carpetas_por_periodo()
        carpetas = {
            periodo: otra for periodo, otra in actuales.items()
            if periodo == str(aamm) or otra != ruta
        }
        carpetas[str(aamm)] = ruta
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
    # CARPETA DEL MES
    # --------------------------------------------------------

    frame_carpeta = tk.LabelFrame(
        contenedor, text="Carpeta del mes (la misma del BESS)", padx=10, pady=8
    )
    frame_carpeta.pack(fill="x", padx=14, pady=(10, 4))

    lbl_base = tk.Label(
        frame_carpeta, textvariable=var_base, wraplength=980, justify="left",
        anchor="w", cursor="hand2", font=("Segoe UI", 9),
    )
    lbl_base.pack(anchor="w")
    lbl_base.bind(
        "<Button-1>",
        lambda e: abrir_en_explorador(var_base.get()) if var_base.get() else None,
    )

    # --------------------------------------------------------
    # PERIODO (AAMM)
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
        fila_aamm, textvariable=var_aamm, width=8, justify="center",
        font=("Consolas", 11, "bold"), validate="key",
        validatecommand=validador_aamm,
    )
    entry_aamm.pack(side="left")

    boton_ir_mes = tk.Button(
        fila_aamm, text="Ir a este mes", bg="#1f5fa8", fg="white",
        font=("Segoe UI", 9, "bold"),
    )
    boton_ir_mes.pack(side="left", padx=8)

    tk.Label(
        fila_aamm,
        text=(
            "4 digitos: año+mes simplificado. Ej. 2608 para agosto de "
            "2026. Con el se buscan los archivos del periodo (energia, "
            "fp_, cvar_cra_, prorrata) y se arma el nombre de la planilla."
        ),
        fg="#444444", font=("Segoe UI", 8), wraplength=760, justify="left",
        anchor="w",
    ).pack(side="left", padx=6)

    var_cambio_hora = tk.StringVar(value="")
    label_cambio_hora = tk.Label(
        frame_periodo, textvariable=var_cambio_hora,
        font=("Segoe UI", 9, "bold"), fg=COLOR_NEUTRO, wraplength=900,
        justify="left", anchor="w",
    )
    label_cambio_hora.pack(fill="x", pady=(6, 0))

    def mostrar_cambio_hora(*_):
        aamm = var_aamm.get().strip()
        try:
            cra.validar_aamm(aamm)
            texto = cambio_hora.describir_periodo(aamm)
            hay = bool(cambio_hora.cambios_del_periodo(aamm))
        except (cra.ErrorEntrada, ValueError):
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
                filas_arbol, text="Selecciona la carpeta del mes.",
                fg=COLOR_NEUTRO, anchor="w",
            ).pack(fill="x")
            return

        prefijos = _prefijos_arbol([fila["nivel"] for fila in filas])
        for fila, prefijo in zip(filas, prefijos):
            _fila_arbol(prefijo, fila, _boton_de_fila(fila))

        habilitar_botones(not corriendo["activo"])

    def habilitar_botones(activos):
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

    def revisar(*_):
        ruta = var_base.get()

        if not ruta or not Path(ruta).is_dir():
            lbl_base.config(fg=COLOR_FALTA)
            pintar_arbol([])
            btn_abrir_carpeta.config(state="disabled")
            return

        lbl_base.config(fg=COLOR_BASE_OK)
        btn_abrir_carpeta.config(state="normal")

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
            f"Faltan {len(faltan)} entradas." if faltan
            else "Entradas completas."
        )

    def ofrecer_completar(ruta, parent=None):
        """Crea, si el usuario quiere, las subcarpetas que le falten."""

        faltan = [
            sub for sub in cra.subcarpetas_del_caso(ruta)
            if not (Path(ruta) / sub).is_dir()
        ]
        if faltan and messagebox.askyesno(
            "Faltan carpetas del caso",
            f"A {Path(ruta).name} le faltan {len(faltan)} carpeta(s):\n\n"
            + "\n".join(f"  {sub}/" for sub in faltan)
            + "\n\n¿Las creo ahora? (vacias; no se toca nada de lo que "
              "ya hay)",
            parent=parent or root,
        ):
            crear_carpetas_en(ruta)

    def usar_carpeta(ruta, completar=True, parent=None):
        ruta = str(ruta)
        var_base.set(ruta)
        aamm = var_aamm.get().strip()
        try:
            cra.validar_aamm(aamm)
        except cra.ErrorEntrada:
            aamm = ""
        recordar_carpeta(aamm, ruta, ademas={"carpeta_base": ruta})
        log(f"Carpeta del periodo {aamm}: {ruta}" if aamm else f"Carpeta: {ruta}")
        revisar()
        if completar:
            ofrecer_completar(ruta, parent)

    def examinar(parent=None, titulo="Selecciona la carpeta del mes"):
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

    def actualizar_vista():
        if corriendo["activo"]:
            return
        revisar()
        log("Carpetas revisadas de nuevo.")

    frame_botones_carpeta = tk.Frame(frame_carpeta)
    frame_botones_carpeta.pack(pady=(6, 0))
    tk.Button(
        frame_botones_carpeta, text="Examinar", command=examinar
    ).pack(side="left", padx=4)
    tk.Button(
        frame_botones_carpeta, text="Actualizar", command=actualizar_vista
    ).pack(side="left", padx=4)

    # --------------------------------------------------------
    # UN PERIODO, UNA CARPETA (como en el BESS)
    # --------------------------------------------------------

    ultimo_aamm = {"valor": var_aamm.get().strip()}
    preguntando = {"activo": False}

    def presentar(top, modal=False):
        top.update_idletasks()
        ancho, alto = top.winfo_reqwidth(), top.winfo_reqheight()
        x = root.winfo_rootx() + (root.winfo_width() - ancho) // 2
        y = root.winfo_rooty() + (root.winfo_height() - alto) // 3
        top.geometry(f"+{max(x, 0)}+{max(y, 0)}")
        top.deiconify()
        if modal:
            top.wait_visibility()
            top.grab_set()

    def volver_al_periodo_anterior():
        var_aamm.set(ultimo_aamm["valor"])
        guardar_config({"aamm": ultimo_aamm["valor"]})
        revisar()

    def ventana_carpeta_del_periodo(aamm, situacion):
        """El periodo no tiene carpeta: examinar, crear o 'esta es'."""

        base = var_base.get()
        top = tk.Toplevel(root)
        top.withdraw()
        top.title(f"Carpeta del periodo {aamm}")
        top.transient(root)
        top.resizable(False, False)

        if situacion == rutas_nucleo.CARPETA_DE_OTRO_PERIODO:
            explicacion = (
                f"La carpeta que esta abierta no es la del periodo {aamm}:"
                f"\n\n    {base or '(ninguna)'}\n\nCada periodo trabaja en "
                f"su propia carpeta. Elegi la de {aamm}, o crea una nueva."
            )
        else:
            explicacion = (
                f"No se puede saber de que periodo es la carpeta que esta "
                f"abierta:\n\n    {base}\n\nSu nombre no trae ningun AAMM. "
                f"Si ESA es la carpeta de {aamm}, decilo con el tercer "
                f"boton; si no, elegi la de {aamm} o crea una nueva."
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
            if examinar(parent=top, titulo=f"Selecciona la carpeta del periodo {aamm}"):
                cerrar(True)

        def con_crear():
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
        if situacion == rutas_nucleo.CARPETA_SIN_PERIODO and base:
            tk.Button(
                pie, text=f"Esta es la de {aamm}", command=con_esta, width=18,
            ).pack(side="left", padx=8)
        tk.Button(
            pie, text="Cancelar", command=lambda: cerrar(False), width=12,
        ).pack(side="right")
        top.protocol("WM_DELETE_WINDOW", lambda: cerrar(False))

        presentar(top, modal=True)
        root.wait_window(top)

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
            cra.validar_aamm(aamm)
        except cra.ErrorEntrada:
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

        # 1. La carpeta de ese periodo (la del CRA o la del BESS).
        recordada = carpeta_recordada(aamm)
        if recordada and recordada != var_base.get():
            log(f"Periodo {aamm}: se abre su carpeta, {recordada}")
            usar_carpeta(recordada)
            return

        # 2. ¿La abierta es la de este periodo?
        recordadas = {**carpetas_del_bess(), **carpetas_por_periodo()}
        situacion = rutas_nucleo.carpeta_corresponde_al_periodo(
            var_base.get(), aamm, recordadas,
        )
        if situacion != rutas_nucleo.CARPETA_DEL_PERIODO:
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

        # 3. Es la de este periodo.
        recordar_carpeta(aamm, var_base.get())
        revisar()
        ofrecer_completar(var_base.get())

    entry_aamm.bind("<FocusOut>", aamm_cambiado)
    entry_aamm.bind("<Return>", lambda e: root.focus_set())
    boton_ir_mes.config(command=lambda: (root.focus_set(), aamm_cambiado()))

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
    }

    def lanzar(funcion, kwargs, etiqueta_salida):
        if corriendo["activo"]:
            return

        base, aamm = var_base.get().strip(), var_aamm.get().strip()
        if not base or not Path(base).is_dir():
            messagebox.showwarning(
                "Falta la carpeta", "Elegi primero la carpeta del mes."
            )
            return
        try:
            cra.validar_aamm(aamm)
        except cra.ErrorEntrada as error:
            messagebox.showwarning("Falta el periodo", str(error))
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
    # CASO NUEVO
    # --------------------------------------------------------

    def crear_carpetas_en(base, usar=True):
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
        if usar:
            usar_carpeta(base, completar=False)
        else:
            revisar()
        return True

    def ventana_nuevo_caso(aamm_objetivo=None):
        """Crea la carpeta de un mes con sus subcarpetas (las del CRA y comunes)."""

        if corriendo["activo"]:
            return False

        aamm = (aamm_objetivo or var_aamm.get()).strip()
        base_actual = Path(var_base.get()) if var_base.get() else None
        padre_inicial = (
            str(base_actual.parent)
            if base_actual and base_actual.parent.is_dir() else str(Path.home())
        )
        try:
            nombre_inicial = rutas_nucleo.nombre_caso_sugerido(aamm, base_actual)
        except Exception:
            nombre_inicial = base_actual.name if base_actual else "Caso nuevo"

        top = tk.Toplevel(root)
        top.withdraw()
        top.title("Crear la carpeta del mes")
        top.transient(root)
        resultado = {"ok": False}
        var_padre = tk.StringVar(value=padre_inicial)
        var_nombre = tk.StringVar(value=nombre_inicial)

        tk.Label(
            top,
            text=(
                f"Se va a crear la carpeta del periodo {aamm or '(sin AAMM)'} "
                f"con las subcarpetas del CRA y las comunes.\nSi el BESS ya "
                f"tiene carpeta para este mes, elegila con Examinar: es la "
                f"misma.\nNo se copia ni se mueve ningun archivo."
            ),
            font=("Segoe UI", 9), anchor="w", justify="left",
        ).pack(fill="x", padx=14, pady=(12, 8))

        frame_padre = tk.LabelFrame(top, text="Donde crearla", padx=10, pady=8)
        frame_padre.pack(fill="x", padx=14, pady=4)
        tk.Entry(
            frame_padre, textvariable=var_padre, font=("Segoe UI", 9), width=70,
        ).pack(side="left", fill="x", expand=True)

        def elegir_padre():
            elegida = filedialog.askdirectory(
                title="Carpeta donde crear el mes",
                initialdir=var_padre.get() or str(Path.home()), parent=top,
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
            text="\n".join(f"{sub}/" for sub in cra.parametros.SUBCARPETAS_CASO),
            font=("Consolas", 9), anchor="w", justify="left", fg=COLOR_NEUTRO,
        ).pack(fill="both", expand=True)

        frame_pie = tk.Frame(top)
        frame_pie.pack(fill="x", padx=14, pady=(4, 12))

        def crear():
            padre, nombre = var_padre.get().strip(), var_nombre.get().strip()
            if not padre or not Path(padre).is_dir():
                messagebox.showwarning(
                    "Falta donde crearla",
                    "Elegi una carpeta que exista para crear el mes adentro.",
                    parent=top,
                )
                return
            if not nombre:
                messagebox.showwarning(
                    "Falta el nombre", "Escribi como se va a llamar la carpeta.",
                    parent=top,
                )
                return
            destino = Path(padre) / nombre
            if destino.is_dir() and not messagebox.askyesno(
                "Ya existe",
                f"{destino} ya existe.\n\nSe le van a crear las subcarpetas "
                f"que le falten (no se toca nada de lo que ya tiene). ¿Seguir?",
                parent=top,
            ):
                return
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

        presentar(top, modal=True)
        root.wait_window(top)
        return resultado["ok"]

    # --------------------------------------------------------
    # BOTONES FIJOS ABAJO
    # --------------------------------------------------------

    tk.Button(
        frame_botones, text="Crear carpeta del caso",
        command=lambda: ventana_nuevo_caso(var_aamm.get().strip()),
    ).pack(side="left", padx=10, expand=True)

    btn_abrir_carpeta = tk.Button(
        frame_botones, text="Abrir carpeta del caso", state="disabled",
        command=lambda: abrir_en_explorador(var_base.get()),
    )
    btn_abrir_carpeta.pack(side="left", padx=10, expand=True)

    tk.Button(frame_botones, text="Salir", command=root.destroy).pack(
        side="left", padx=10, expand=True
    )

    # --------------------------------------------------------
    # ARRANQUE
    # --------------------------------------------------------

    pintar_arbol([])
    log(
        "Selecciona la carpeta del mes (la misma del BESS) e ingresa el "
        "periodo (AAMM). Cada fila del diagrama trae su propio boton: "
        "'Generar' en FP/ y CO/, 'Traer' en el SSCC_Desempeño de FD y FMA/, "
        "y 'Actualizar' en la planilla (todas las hojas) y en cada hoja. "
        "Cada archivo y cada carpeta del diagrama es un link a su ruta."
    )
    if var_aamm.get():
        log(f"Periodo recordado: {var_aamm.get()}")
    if var_base.get():
        log(f"Carpeta recordada: {var_base.get()}")
        revisar()
        try:
            cra.validar_aamm(var_aamm.get())
        except cra.ErrorEntrada:
            pass
        else:
            if rutas_nucleo.carpeta_corresponde_al_periodo(
                var_base.get(), var_aamm.get(), carpetas_por_periodo(),
            ) != rutas_nucleo.CARPETA_DE_OTRO_PERIODO:
                recordar_carpeta(var_aamm.get(), var_base.get())

    root.mainloop()


if __name__ == "__main__":
    main()
