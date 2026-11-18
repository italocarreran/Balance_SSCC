# -*- coding: utf-8 -*-
"""
EXPERIMENTO: ¿cuanta plata cambia si la Componente 2 de RE545 cuenta solo
la inyeccion real? (observacion O2026-0052689, AES Andes)

Hoy, en un bloque con inyeccion, la "Asignacion Edisponible" cuenta la
inyeccion MAS la Energia asignada de la Componente 1 (BQ = BS + BN): con
1 kWh inyectado se descuenta casi la caja entera (Pmax/4). El escenario
cuenta solo la inyeccion (BS), con el mismo orden (de menor a mayor CMg) y
el mismo tope (re545_componentes.C2_SOLO_INYECCION = True).

Como se mide (con el motor real del programa):
  1. Dos copias de la planilla 11_PAGOS_BESS en una carpeta de trabajo:
     BASE y ESCENARIO.
  2. Las dos se recalculan con el codigo actual (Calculo RE545 y
     COMPENSACION_CENTRAL); el ESCENARIO con C2_SOLO_INYECCION = True.
     Recalcular tambien la BASE separa el efecto del cambio de otros
     cambios del programa desde que se genero la planilla.
  3. Se comparan con Comparar_Meses: el reporte queda en
     <carpeta de trabajo>/Comparacion_c2_solo_inyeccion.xlsx, y en
     pantalla sale la diferencia por central.

Uso:
    Con ventana (pedido del usuario: "selecciono el archivo y aprieto
    play"): abrir este archivo en VS Code y darle Play, o doble clic. Se
    elige el 11_, se aprieta "Simular" y la ventana muestra la diferencia
    por central.

    Por consola (desde la carpeta del programa):
    python desarrollo/experimentos/c2_solo_inyeccion.py <11_PAGOS_BESS_AAMM.xlsx> [carpeta]
"""

import shutil
import sys
import time
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[2]
if str(RAIZ) not in sys.path:
    sys.path.insert(0, str(RAIZ))

import pandas as pd  # noqa: E402

from Script import nucleo  # noqa: E402
from Script.nucleo import re545_componentes, reporte  # noqa: E402

SECCIONES = {"re545", "compensacion_central"}


def montos_por_central(ruta):
    """{central: (E Costos, RE545, total)} desde el reporte de la planilla."""

    resumen = reporte.resumen_por_central(ruta, registrar=lambda *_: None)
    return resumen.set_index(reporte.C_CENTRAL)[
        [reporte.C_ECOSTOS, reporte.C_RE545, reporte.C_TOTAL]
    ]


def main(ruta_planilla, trabajo=None, escribir=print):
    ruta_planilla = Path(ruta_planilla).resolve()
    aamm = reporte._periodo(ruta_planilla)
    trabajo = Path(trabajo or ruta_planilla.parent / "Log" / "experimento_c2_solo_inyeccion")
    base, escenario = trabajo / "base", trabajo / "escenario"
    for carpeta in (base, escenario):
        if carpeta.exists():
            shutil.rmtree(carpeta)
        carpeta.mkdir(parents=True)
        shutil.copy2(ruta_planilla, carpeta / ruta_planilla.name)
        # Si la planilla no trae la hoja Centrales, el calculo usa
        # Auxiliares/Centrales.xlsx de al lado.
        auxiliares = ruta_planilla.parent / "Auxiliares"
        if auxiliares.is_dir():
            shutil.copytree(auxiliares, carpeta / "Auxiliares")

    callado = lambda *_: None  # noqa: E731
    original = re545_componentes.C2_SOLO_INYECCION
    try:
        for nombre, carpeta, regla in (("base", base, False), ("escenario", escenario, True)):
            re545_componentes.C2_SOLO_INYECCION = regla
            inicio = time.time()
            escribir(f"Recalculando {nombre} (C2 solo inyeccion = {regla})...")
            nucleo.generar_pagos_bess(carpeta, set(SECCIONES), registrar=callado, aamm=aamm)
            escribir(f"  listo en {time.time() - inicio:.0f} s")
    finally:
        re545_componentes.C2_SOLO_INYECCION = original

    antes = montos_por_central(base / ruta_planilla.name)
    despues = montos_por_central(escenario / ruta_planilla.name)
    tabla = pd.DataFrame({
        "RE545 hoy [CLP]": antes[reporte.C_RE545],
        "RE545 propuesta [CLP]": despues[reporte.C_RE545],
    }).fillna(0.0)
    tabla["Diferencia [CLP]"] = tabla["RE545 propuesta [CLP]"] - tabla["RE545 hoy [CLP]"]
    tabla = tabla.sort_values("Diferencia [CLP]", ascending=False)

    destino, lineas, _ = reporte.comparar_planillas(
        escenario / ruta_planilla.name, base / ruta_planilla.name, registrar=callado,
    )
    final = trabajo / "Comparacion_c2_solo_inyeccion.xlsx"
    shutil.move(destino, final)
    tabla.to_excel(trabajo / "Diferencia_RE545_por_central.xlsx")

    escribir("")
    escribir("RE545 por central, hoy vs. contando solo la inyeccion (mismo codigo):")
    for central, fila in tabla.iterrows():
        if abs(fila["Diferencia [CLP]"]) >= 1:
            escribir(f"  {central}: {fila['RE545 hoy [CLP]']:>15,.0f} -> "
                     f"{fila['RE545 propuesta [CLP]']:>15,.0f}  ({fila['Diferencia [CLP]']:+,.0f})")
    escribir(f"\n  TOTAL RE545: {tabla['RE545 hoy [CLP]'].sum():,.0f} -> "
             f"{tabla['RE545 propuesta [CLP]'].sum():,.0f}  ({tabla['Diferencia [CLP]'].sum():+,.0f})")
    escribir(f"\nReporte: {final}")
    escribir(f"Tabla:   {trabajo / 'Diferencia_RE545_por_central.xlsx'}")
    main.ultima_tabla = tabla
    main.ultima_carpeta = trabajo
    return final


# ============================================================
# VENTANA: elegir el 11_ y apretar "Simular"
# ============================================================

CLAVE_CONFIG = "experimento_c2_solo_inyeccion"


def _recordado():
    try:
        from Script import config
        return config.seccion_usuario(config.MODO_BESS, _usuario()).get(CLAVE_CONFIG, "")
    except Exception:  # noqa: BLE001 - recordar es una comodidad
        return ""


def _recordar(ruta):
    try:
        from Script import config
        config.actualizar_seccion_usuario(config.MODO_BESS, _usuario(), {CLAVE_CONFIG: str(ruta)})
    except Exception:  # noqa: BLE001
        pass


def _usuario():
    import os
    import socket
    return f"{socket.gethostname()}_{os.environ.get('USERNAME') or os.environ.get('USER') or 'desconocido'}"


def ventana():
    import queue
    import threading
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    from Script.ventana_estilo import abrir_en_explorador

    root = tk.Tk()
    root.title("Simular C2 de AES (O2026-0052689)")
    root.geometry("900x620")

    var_ruta = tk.StringVar(value=_recordado())
    var_estado = tk.StringVar(value="Elige el 11_PAGOS_BESS y aprieta Simular.")
    cola = queue.Queue()
    estado = {"corriendo": False, "inicio": 0.0}

    tk.Label(
        root, font=("Segoe UI", 9), justify="left", anchor="w", wraplength=860,
        text=("Recalcula Calculo RE545 y COMPENSACION_CENTRAL dos veces: con la regla "
              "de hoy y contando en la Componente 2 solo la inyeccion real. El 11_ "
              "original no se toca (trabaja sobre dos copias)."),
    ).pack(fill="x", padx=12, pady=(10, 4))

    marco = tk.LabelFrame(root, text="Planilla 11_PAGOS_BESS", padx=8, pady=6)
    marco.pack(fill="x", padx=12, pady=4)
    tk.Entry(marco, textvariable=var_ruta, font=("Segoe UI", 9)).pack(
        side="left", fill="x", expand=True)

    def examinar():
        inicial = Path(var_ruta.get()).parent if var_ruta.get() else Path.home()
        ruta = filedialog.askopenfilename(
            title="Elige el 11_PAGOS_BESS_AAMM.xlsx", parent=root,
            initialdir=str(inicial) if inicial.is_dir() else "",
            filetypes=[("Excel", "*.xlsx"), ("Todos", "*.*")],
        )
        if ruta:
            var_ruta.set(ruta)

    tk.Button(marco, text="Examinar", command=examinar).pack(side="left", padx=(8, 0))

    botones = tk.Frame(root)
    botones.pack(fill="x", padx=12, pady=6)
    btn_simular = tk.Button(
        botones, text="▶  Simular", width=16, bg="#1f5fa8", fg="white",
        font=("Segoe UI", 10, "bold"),
    )
    btn_simular.pack(side="left")
    btn_carpeta = tk.Button(botones, text="Abrir carpeta de resultados", state="disabled")
    btn_carpeta.pack(side="left", padx=8)
    tk.Label(botones, textvariable=var_estado, font=("Segoe UI", 9), anchor="w").pack(
        side="left", padx=8)

    tabla = ttk.Treeview(
        root, columns=("hoy", "propuesta", "dif"), height=10,
    )
    tabla.heading("#0", text="Central")
    tabla.heading("hoy", text="RE545 hoy [CLP]")
    tabla.heading("propuesta", text="RE545 propuesta [CLP]")
    tabla.heading("dif", text="Diferencia [CLP]")
    tabla.column("#0", width=260)
    for columna in ("hoy", "propuesta", "dif"):
        tabla.column(columna, width=190, anchor="e")
    tabla.pack(fill="both", expand=True, padx=12, pady=4)

    registro = tk.Text(root, height=9, font=("Consolas", 9))
    registro.pack(fill="both", padx=12, pady=(4, 10))

    def escribir(texto):
        cola.put(("log", texto))

    def vaciar():
        try:
            while True:
                tipo, dato = cola.get_nowait()
                if tipo == "log":
                    registro.insert("end", dato + "\n")
                    registro.see("end")
                elif tipo == "fin":
                    terminar(dato)
        except queue.Empty:
            pass
        if estado["corriendo"]:
            var_estado.set(f"Simulando... {time.time() - estado['inicio']:.0f} s")
        root.after(300, vaciar)

    def terminar(error):
        estado["corriendo"] = False
        btn_simular.config(state="normal")
        if error:
            var_estado.set("Termino con error (ver abajo).")
            messagebox.showerror("No se pudo simular", error, parent=root)
            return
        resultado = getattr(main, "ultima_tabla", None)
        tabla.delete(*tabla.get_children())
        if resultado is not None:
            for central, fila in resultado.iterrows():
                tabla.insert("", "end", text=str(central), values=(
                    f"{fila['RE545 hoy [CLP]']:,.0f}",
                    f"{fila['RE545 propuesta [CLP]']:,.0f}",
                    f"{fila['Diferencia [CLP]']:+,.0f}",
                ))
            tabla.insert("", "end", text="TOTAL", values=(
                f"{resultado['RE545 hoy [CLP]'].sum():,.0f}",
                f"{resultado['RE545 propuesta [CLP]'].sum():,.0f}",
                f"{resultado['Diferencia [CLP]'].sum():+,.0f}",
            ))
            var_estado.set(
                f"Listo en {time.time() - estado['inicio']:.0f} s. Diferencia total: "
                f"{resultado['Diferencia [CLP]'].sum():+,.0f} CLP"
            )
        btn_carpeta.config(
            state="normal",
            command=lambda: abrir_en_explorador(str(main.ultima_carpeta)),
        )

    def simular():
        ruta = Path(var_ruta.get().strip().strip('"'))
        if not ruta.is_file():
            messagebox.showwarning("Falta el archivo", "Elige el 11_PAGOS_BESS_AAMM.xlsx.", parent=root)
            return
        _recordar(ruta)
        estado["corriendo"] = True
        estado["inicio"] = time.time()
        btn_simular.config(state="disabled")
        btn_carpeta.config(state="disabled")
        registro.delete("1.0", "end")
        tabla.delete(*tabla.get_children())

        def trabajo():
            error = None
            try:
                main(ruta, escribir=escribir)
            except Exception as exc:  # noqa: BLE001 - se muestra en la ventana
                import traceback
                escribir(traceback.format_exc())
                error = str(exc)
            cola.put(("fin", error))

        threading.Thread(target=trabajo, daemon=True).start()

    btn_simular.config(command=simular)
    vaciar()
    root.mainloop()


if __name__ == "__main__":
    if len(sys.argv) < 2:
        ventana()
        sys.exit(0)
    main(*sys.argv[1:3])
