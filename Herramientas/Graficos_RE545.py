# -*- coding: utf-8 -*-
"""
Graficos interactivos del Calculo RE545 de una planilla 11_PAGOS_BESS.

Pedido del usuario (2026-11-18): elegir el archivo 11 ya calculado y ver,
por central, como se arma la compensacion RE545, con graficos que se puedan
acercar y un simulador para ver como afectan la Energia Disponible, la
reserva, la Pmax, etc.

Que hace:
  - lee la hoja "Calculo RE545" y "Centrales" de la planilla (no recalcula
    nada en Python);
  - escribe <carpeta de la planilla>/Log/Graficos_RE545_<AAMM>.html, una
    pagina que se abre con doble clic y funciona sin internet;
  - la abre en el navegador.

La libreria de graficos (Plotly) se embebe en la pagina. La primera vez se
baja de internet y queda guardada en esta carpeta (plotly-<version>.min.js);
si no hay internet, se puede dejar ahi una copia a mano.

Como Comparar_Meses, usa el programa (Script/): tiene que estar en la
carpeta Herramientas del repositorio.

Uso:  python Herramientas/Graficos_RE545.py
"""

import json
import sys
import threading
from pathlib import Path

# El repositorio (la carpeta de arriba) para importar Script.
RAIZ = Path(__file__).resolve().parents[1]
if str(RAIZ) not in sys.path:
    sys.path.insert(0, str(RAIZ))

PATRON_PLANILLA = "11_PAGOS_BESS_*.xlsx"


def carpeta_config():
    """El __config__ mas cercano hacia arriba, o esta misma carpeta."""

    carpeta = Path(__file__).resolve().parent
    for candidata in (carpeta, *carpeta.parents):
        if (candidata / "__config__").is_dir():
            return candidata / "__config__"
    return carpeta


def ruta_recordada():
    return carpeta_config() / "graficos_re545.json"


def leer_recordada():
    try:
        return json.loads(ruta_recordada().read_text(encoding="utf-8")).get("planilla", "")
    except (OSError, ValueError, AttributeError):
        return ""


def guardar_recordada(planilla):
    try:
        ruta_recordada().write_text(
            json.dumps({"planilla": str(planilla)}, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
    except OSError:
        pass  # recordar es una comodidad


def mostrar_error(titulo, mensaje):
    try:
        from Script.dialogos import mostrar_error as mostrar
    except Exception:  # noqa: BLE001 - sin el programa, el de tkinter
        from tkinter import messagebox
        messagebox.showerror(titulo, mensaje)
        return
    mostrar(titulo, mensaje)


def abrir(ruta):
    import os
    import subprocess
    import webbrowser

    try:
        if sys.platform.startswith("win"):
            os.startfile(ruta)  # noqa: S606 - abrir la pagina
        elif sys.platform == "darwin":
            subprocess.Popen(["open", str(ruta)])
        else:
            webbrowser.open(Path(ruta).as_uri())
    except OSError as exc:
        mostrar_error("No se pudo abrir la pagina", str(exc))


def main():
    import tkinter as tk
    from tkinter import filedialog

    root = tk.Tk()
    root.title("Graficos RE545")
    root.geometry("820x460")

    var_planilla = tk.StringVar(value=leer_recordada())
    var_info = tk.StringVar(value="Elige la planilla 11_PAGOS_BESS ya calculada.")
    estado = {"pagina": None}

    marco = tk.LabelFrame(root, text="Planilla 11_PAGOS_BESS", padx=10, pady=6)
    marco.pack(fill="x", padx=14, pady=(10, 2))
    tk.Label(marco, textvariable=var_planilla, anchor="w", justify="left",
             wraplength=640, font=("Segoe UI", 9)).pack(side="left", fill="x", expand=True)

    def examinar():
        actual = var_planilla.get()
        inicial = Path(actual).parent if actual else None
        ruta = filedialog.askopenfilename(
            title="Planilla 11_PAGOS_BESS",
            initialdir=str(inicial) if inicial and inicial.is_dir() else None,
            filetypes=[("Planilla 11_PAGOS_BESS", PATRON_PLANILLA),
                       ("Excel", "*.xlsx"), ("Todos", "*.*")],
        )
        if ruta:
            var_planilla.set(ruta)

    boton_examinar = tk.Button(marco, text="Examinar...", command=examinar)
    boton_examinar.pack(side="right")

    abajo = tk.Frame(root)
    abajo.pack(fill="x", padx=14, pady=8)
    boton_generar = tk.Button(abajo, text="Generar graficos", bg="#0d7472", fg="white",
                              width=16, font=("Segoe UI", 9, "bold"))
    boton_generar.pack(side="left")
    boton_abrir = tk.Button(abajo, text="Abrir pagina", width=14, state="disabled")
    boton_abrir.pack(side="left", padx=8)
    tk.Label(abajo, textvariable=var_info, anchor="w", fg="#444444").pack(side="left", padx=8)

    marco_texto = tk.Frame(root)
    marco_texto.pack(fill="both", expand=True, padx=14, pady=(0, 12))
    caja = tk.Text(marco_texto, wrap="word", font=("Consolas", 9))
    barra = tk.Scrollbar(marco_texto, command=caja.yview)
    caja.configure(yscrollcommand=barra.set)
    barra.pack(side="right", fill="y")
    caja.pack(side="left", fill="both", expand=True)

    def escribir(texto):
        caja.insert("end", texto + "\n")
        caja.see("end")

    def generar():
        planilla = var_planilla.get().strip()
        if not planilla:
            var_info.set("Falta la planilla.")
            return
        guardar_recordada(planilla)
        caja.delete("1.0", "end")
        for boton in (boton_examinar, boton_generar):
            boton.config(state="disabled")
        var_info.set("Generando...")

        def registrar(texto):
            root.after(0, lambda: escribir(str(texto)))

        def trabajo():
            from Script.nucleo import graficos_re545

            try:
                destino = graficos_re545.generar_html(
                    planilla, registrar=registrar,
                    carpetas_plotly=(Path(__file__).resolve().parent,),
                )
                error = None
            except Exception as exc:  # noqa: BLE001 - se muestra al usuario
                destino, error = None, exc
            root.after(0, lambda: terminado(destino, error))

        threading.Thread(target=trabajo, daemon=True).start()

    def terminado(destino, error):
        for boton in (boton_examinar, boton_generar):
            boton.config(state="normal")
        if error is not None:
            var_info.set("No se pudo generar.")
            mostrar_error("No se pudieron generar los graficos", str(error))
            return
        estado["pagina"] = destino
        boton_abrir.config(state="normal")
        var_info.set(f"Listo: {destino.name}")
        abrir(destino)

    boton_generar.config(command=generar)
    boton_abrir.config(command=lambda: estado["pagina"] and abrir(estado["pagina"]))
    root.mainloop()


if __name__ == "__main__":
    main()
