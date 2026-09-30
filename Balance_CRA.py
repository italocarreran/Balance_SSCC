# -*- coding: utf-8 -*-
"""
Balance CRA (remuneracion CRA).

Ventana unica, con el mismo patron que Balance_BESS.py: se elige la
carpeta base del caso, se ingresa el periodo (AAMM) y debajo se dibuja
el arbol del caso con el estado de cada entrada (OK/FALTA/PENDIENTE) y
un boton en cada fila que se puede generar.

    <CARPETA_BASE>/
        Energia/
            Formato_Solicitud_SSAA_SSCC_Hidro_*.xlsx
        FP/
            fp_*.xlsx
        CO/
            cvar_cra_<AAMM>_*.xlsx
        SC y CO/
            Reporte_CRA*.csv
            Cálculo_SobrecostosSSCC_*.xlsm (hoja SOBRECOSTOS)
        FD/
            SSCC_Desempeño_*.xlsx          [Traer]
        Prorrata retiros/
            Prorrata_Retiros_<AAMM>_pre/_def.xlsx
        Balance_CRA.xlsx                   [Actualizar]
            hoja 'PRORRATA_RETIROS'        [Actualizar]  (solo bloque 1)
            hoja 'FD_CTF' / 'FD_CSF' / 'FD_CPF'  [Actualizar]
            hoja 'SC y CO'                 [Actualizar]
            hoja 'CO'                      [Actualizar]
            hoja 'FP'                      [Actualizar]
            hoja 'ENERGIA'                 [Actualizar]

Por ahora solo estan las hojas de entrada con su carga definida en
docs/Trazabilidad_CRA_Periodo_Generico_v5_Auditoria_Formulas.md; el resto se muestra como
PENDIENTE.

Todo el calculo vive en Script/cra: esta ventana solo lo llama.
"""

import os
import socket
import subprocess
import sys
import threading
import time
import traceback
from pathlib import Path

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from Script import config, cra
from Script.arbol import _prefijos_arbol


COLOR_ESTADO = {
    "ok": "#1a6b1a",
    "falta": "#b00020",
    "pendiente": "#a06000",
}
SIMBOLO = {"ok": "OK", "falta": "FALTA", "pendiente": "PENDIENTE"}

ANCHO_ESTRUCTURA = 380     # pixeles
ANCHO_ESTADO = 86
ANCHO_ACCION = 104
ALTO_FILA = 26


# ============================================================
# CONFIG POR PC/USUARIO
#
# Seccion propia ("<host>_<usuario>_CRA"): la carpeta del CRA no pisa
# la que recuerda Balance_BESS.py para el mismo usuario.
# ============================================================

def get_usuario():
    usuario = (
        os.environ.get("USERNAME")
        or os.environ.get("USER")
        or "desconocido"
    )
    return f"{socket.gethostname()}_{usuario}_CRA"


def leer_config():
    return config.seccion(get_usuario())


def guardar_config(data):
    config.actualizar_seccion(get_usuario(), data)


def abrir_en_explorador(ruta, es_carpeta=False):
    """Abre la carpeta de la ruta (la que la contiene si es archivo)."""

    ruta = Path(ruta)
    carpeta = ruta if es_carpeta else ruta.parent
    while not carpeta.exists() and carpeta != carpeta.parent:
        carpeta = carpeta.parent

    try:
        if sys.platform.startswith("win"):
            os.startfile(str(carpeta))          # noqa: S606
        elif sys.platform == "darwin":
            subprocess.Popen(["open", str(carpeta)])
        else:
            subprocess.Popen(["xdg-open", str(carpeta)])
    except OSError:
        pass


def formato_tiempo(segundos):
    minutos, segundos = divmod(int(segundos), 60)
    return f"{minutos:02d}:{segundos:02d}"


# ============================================================
# VENTANA
# ============================================================

def main():
    root = tk.Tk()
    root.title("Balance CRA")
    root.geometry("980x720")

    guardado = leer_config()
    var_base = tk.StringVar(value=guardado.get("carpeta_base", ""))
    var_aamm = tk.StringVar(value=guardado.get("aamm", ""))
    var_estado = tk.StringVar(value="")
    var_tiempo = tk.StringVar(value="00:00")

    corriendo = {"activo": False}
    botones = []
    timer = {"corriendo": False, "inicio": 0.0}

    # --- carpeta base y periodo -----------------------------
    arriba = ttk.Frame(root, padding=8)
    arriba.pack(fill="x")

    ttk.Label(arriba, text="Carpeta base:").grid(row=0, column=0, sticky="w")
    ttk.Entry(arriba, textvariable=var_base, width=90).grid(
        row=0, column=1, sticky="we", padx=4
    )

    def examinar():
        elegida = filedialog.askdirectory(
            initialdir=var_base.get() or None,
            title="Carpeta del caso CRA",
        )
        if elegida:
            var_base.set(elegida)
            revisar()

    ttk.Button(arriba, text="Examinar…", command=examinar).grid(
        row=0, column=2
    )

    ttk.Label(arriba, text="Periodo (AAMM):").grid(
        row=1, column=0, sticky="w", pady=(6, 0)
    )
    entrada_aamm = ttk.Entry(arriba, textvariable=var_aamm, width=8)
    entrada_aamm.grid(row=1, column=1, sticky="w", padx=4, pady=(6, 0))
    entrada_aamm.bind("<Return>", lambda _e: revisar())
    entrada_aamm.bind("<FocusOut>", lambda _e: revisar())
    arriba.columnconfigure(1, weight=1)

    # --- arbol ----------------------------------------------
    marco_arbol = ttk.LabelFrame(root, text="Estructura del caso", padding=6)
    marco_arbol.pack(fill="both", expand=False, padx=8)

    # --- botones de abajo -----------------------------------
    abajo = ttk.Frame(root, padding=(8, 4))
    abajo.pack(fill="x")

    # --- registro, barra y tiempo ---------------------------
    marco_log = ttk.Frame(root, padding=(8, 0, 8, 8))
    marco_log.pack(fill="both", expand=True)

    barra = ttk.Progressbar(marco_log, maximum=100)
    barra.pack(fill="x")
    fila_estado = ttk.Frame(marco_log)
    fila_estado.pack(fill="x")
    ttk.Label(fila_estado, textvariable=var_estado).pack(side="left")
    ttk.Label(fila_estado, textvariable=var_tiempo).pack(side="right")

    txt_log = tk.Text(marco_log, height=12, wrap="word")
    txt_log.pack(fill="both", expand=True)

    def log(mensaje):
        txt_log.insert("end", mensaje + "\n")
        txt_log.see("end")

    def tick():
        if timer["corriendo"]:
            var_tiempo.set(formato_tiempo(time.time() - timer["inicio"]))
            root.after(500, tick)

    def habilitar_botones(habilitar):
        for boton in botones:
            try:
                boton.configure(state="normal" if habilitar else "disabled")
            except tk.TclError:
                pass

    def lanzar(funcion, **kwargs):
        """
        Corre funcion(base, aamm, **kwargs, registrar, progreso) en un
        hilo aparte, reportando al registro y a la barra.
        """

        if corriendo["activo"]:
            return

        base, aamm = var_base.get().strip(), var_aamm.get().strip()
        try:
            cra.validar_aamm(aamm)
        except cra.ErrorEntrada as error:
            messagebox.showerror("Periodo", str(error))
            return
        if not base or not Path(base).is_dir():
            messagebox.showerror("Carpeta base", "Elegi una carpeta base valida.")
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
                root.after(0, var_estado.set, "Procesando...")
                funcion(
                    base, aamm, **kwargs,
                    registrar=lambda m: root.after(0, log, m),
                    progreso=lambda v: root.after(
                        0, barra.configure, {"value": v}
                    ),
                )
                resultado["ok"] = True
            except cra.ErrorEntrada as error:
                root.after(0, log, f"\nERROR DE ENTRADA:\n{error}")
            except Exception as error:
                root.after(
                    0, log,
                    f"\nERROR INESPERADO:\n{error}\n\n{traceback.format_exc()}",
                )
            finally:
                root.after(0, terminar, resultado)

        def terminar(res):
            timer["corriendo"] = False
            corriendo["activo"] = False
            var_estado.set(
                "Terminado correctamente." if res["ok"]
                else "Termino con errores (ver registro)."
            )
            revisar()

        threading.Thread(target=trabajo, daemon=True).start()

    def boton_de_fila(fila):
        """(texto, comando) del boton de esa fila, o None."""

        if fila["id"] == "salida":
            return "Actualizar", lambda: lanzar(cra.generar_balance_cra)

        if fila["id"] == "sscc_desempeno":
            return "Traer", lambda: lanzar(cra.traer_fd)

        if fila["id"].startswith("hoja_"):
            seccion = fila["id"][len("hoja_"):]
            if seccion in cra.SECCIONES:
                return "Actualizar", lambda s=seccion: lanzar(
                    cra.generar_balance_cra, secciones=[s]
                )

        return None

    def revisar():
        base, aamm = var_base.get().strip(), var_aamm.get().strip()
        guardar_config({"carpeta_base": base, "aamm": aamm})

        for hijo in marco_arbol.winfo_children():
            hijo.destroy()
        botones.clear()

        if not base or not Path(base).is_dir():
            ttk.Label(marco_arbol, text="Elegi la carpeta base del caso.").pack(
                anchor="w"
            )
            return

        try:
            filas = cra.revisar_estructura(base, aamm)
        except cra.ErrorEntrada as error:
            ttk.Label(marco_arbol, text=str(error)).pack(anchor="w")
            return

        prefijos = _prefijos_arbol([f["nivel"] for f in filas])

        for fila, prefijo in zip(filas, prefijos):
            renglon = tk.Frame(marco_arbol, height=ALTO_FILA)
            renglon.pack(fill="x")

            celdas = []
            for ancho in (ANCHO_ESTRUCTURA, ANCHO_ESTADO, ANCHO_ACCION):
                celda = tk.Frame(renglon, width=ancho, height=ALTO_FILA)
                celda.pack_propagate(False)
                celda.pack(side="left")
                celdas.append(celda)

            fuente = ("Consolas", 9, "bold" if fila["nivel"] == 0 else "normal")
            etiqueta = tk.Label(
                celdas[0], text=prefijo + fila["texto"], font=fuente,
                anchor="w", cursor="hand2" if fila["ruta"] else "",
            )
            etiqueta.pack(fill="both", expand=True)
            if fila["ruta"]:
                etiqueta.bind(
                    "<Button-1>",
                    lambda _e, r=fila["ruta"], c=fila["es_carpeta"]:
                        abrir_en_explorador(r, c),
                )

            tk.Label(
                celdas[1], text=SIMBOLO[fila["estado"]],
                fg=COLOR_ESTADO[fila["estado"]], anchor="w",
            ).pack(fill="both", expand=True)

            accion = boton_de_fila(fila)
            if accion:
                boton = ttk.Button(celdas[2], text=accion[0], command=accion[1])
                boton.pack(fill="both", expand=True, padx=2, pady=1)
                botones.append(boton)

            if fila["detalle"]:
                tk.Label(
                    renglon, text=fila["detalle"], fg="#555555", anchor="w",
                ).pack(side="left", fill="x")

        habilitar_botones(not corriendo["activo"])

    def crear_carpetas():
        base = var_base.get().strip()
        if not base:
            messagebox.showerror("Carpeta base", "Elegi una carpeta base.")
            return
        creadas = cra.crear_carpetas_caso(base)
        log(
            "Carpetas creadas: " + ", ".join(creadas) if creadas
            else "No faltaba ninguna carpeta."
        )
        revisar()

    ttk.Button(abajo, text="Crear carpetas del caso", command=crear_carpetas).pack(
        side="left"
    )
    ttk.Button(
        abajo, text="Abrir carpeta del caso",
        command=lambda: abrir_en_explorador(var_base.get(), True),
    ).pack(side="left", padx=6)
    ttk.Button(abajo, text="Revisar", command=revisar).pack(side="left")

    revisar()
    root.mainloop()


if __name__ == "__main__":
    main()
