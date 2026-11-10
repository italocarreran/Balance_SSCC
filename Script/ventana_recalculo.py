# -*- coding: utf-8 -*-
"""
La ventana del recalculo que se entrega a los coordinados (antecedentes
SAE): elige el 11_PAGOS_BESS_AAMM.xlsx, tilda que recalcular y lo corre
sobre una copia (ver nucleo/recalculo.py).

Va dentro del archivo Recalcular_11_PAGOS_BESS_AAMM.py que arma
nucleo/antecedentes.py; tambien se puede abrir desde el repo:
    python -m Script.ventana_recalculo
"""

import queue
import sys
import threading
import time
from pathlib import Path

from .dialogos import mostrar_error
from .nucleo.recalculo import SECCIONES_RECALCULO, aamm_del_archivo, buscar_11, recalcular
from .nucleo.utiles import ErrorEntrada

AVISO = (
    "Al ejecutar se recalculan las hojas seleccionadas con los datos del "
    "propio archivo 11_. El resultado queda en una copia "
    "(…_recalculado.xlsx) al lado del archivo; el original no se modifica."
)


def main(ruta_inicial=None):
    import tkinter as tk
    from tkinter import filedialog, messagebox

    root = tk.Tk()
    root.title("Recalcular 11_PAGOS_BESS")
    root.geometry("760x620")

    var_ruta = tk.StringVar(value=str(ruta_inicial) if ruta_inicial else "")
    var_tiempo = tk.StringVar(value="00:00")
    cola = queue.Queue()
    estado = {"corriendo": False, "inicio": 0.0}

    # --- aviso
    tk.Label(
        root, text=AVISO, bg="#fff4ce", fg="#5c4400", justify="left",
        anchor="w", wraplength=700, padx=10, pady=8, font=("Segoe UI", 9),
        relief="groove", bd=1,
    ).pack(fill="x", padx=14, pady=(12, 6))

    # --- archivo
    marco = tk.LabelFrame(root, text="Archivo 11_PAGOS_BESS_AAMM.xlsx", padx=10, pady=6)
    marco.pack(fill="x", padx=14, pady=4)
    tk.Label(
        marco, textvariable=var_ruta, anchor="w", justify="left",
        wraplength=560, font=("Segoe UI", 9),
    ).pack(side="left", fill="x", expand=True)

    def examinar():
        ruta = filedialog.askopenfilename(
            title="Archivo 11_PAGOS_BESS_AAMM.xlsx",
            filetypes=[("Excel", "*.xlsx"), ("Todos", "*.*")],
        )
        if ruta:
            var_ruta.set(ruta)
            pintar_nombres()

    tk.Button(marco, text="Examinar...", command=examinar).pack(side="right")

    # --- que recalcular
    seleccion = tk.LabelFrame(root, text="Qué recalcular", padx=10, pady=6)
    seleccion.pack(fill="x", padx=14, pady=4)
    var_libro = tk.StringVar(value="11_PAGOS_BESS_AAMM_recalculado.xlsx")
    tk.Label(seleccion, textvariable=var_libro, font=("Consolas", 9, "bold"),
             anchor="w").pack(fill="x")
    vars_hojas = {}
    for i, (seccion, hoja) in enumerate(SECCIONES_RECALCULO):
        rama = "└──" if i == len(SECCIONES_RECALCULO) - 1 else "├──"
        var = tk.BooleanVar(value=True)
        vars_hojas[seccion] = var
        tk.Checkbutton(
            seleccion, text=f"{rama} hoja '{hoja}'", variable=var,
            font=("Consolas", 9), anchor="w",
        ).pack(fill="x")
    var_csv = tk.BooleanVar(value=True)
    var_nombre_csv = tk.StringVar(value="SAE_AAMM_recalculado.csv")
    tk.Checkbutton(
        seleccion, textvariable=var_nombre_csv, variable=var_csv,
        font=("Consolas", 9, "bold"), anchor="w",
    ).pack(fill="x", pady=(4, 0))

    def pintar_nombres():
        try:
            aamm = aamm_del_archivo(var_ruta.get())
        except ErrorEntrada:
            aamm = "AAMM"
        var_libro.set(f"11_PAGOS_BESS_{aamm}_recalculado.xlsx")
        var_nombre_csv.set(f"SAE_{aamm}_recalculado.csv")

    pintar_nombres()

    # --- botones y registro
    barra = tk.Frame(root)
    barra.pack(fill="x", padx=14, pady=6)
    boton = tk.Button(barra, text="Ejecutar", bg="#1f5fa8", fg="white",
                      font=("Segoe UI", 10, "bold"), width=14)
    boton.pack(side="left")
    tk.Label(barra, textvariable=var_tiempo, font=("Consolas", 10, "bold"),
             fg="#2d7a2d").pack(side="right")

    marco_log = tk.LabelFrame(root, text="Registro")
    marco_log.pack(fill="both", expand=True, padx=14, pady=(0, 12))
    texto = tk.Text(marco_log, height=12, font=("Consolas", 9), wrap="word")
    barra_log = tk.Scrollbar(marco_log, command=texto.yview)
    texto.configure(yscrollcommand=barra_log.set)
    barra_log.pack(side="right", fill="y")
    texto.pack(side="left", fill="both", expand=True)

    def escribir(mensaje):
        texto.insert("end", f"[{time.strftime('%H:%M:%S')}] {mensaje}\n")
        texto.see("end")

    def bombear():
        try:
            while True:
                tipo, dato = cola.get_nowait()
                if tipo == "log":
                    escribir(dato)
                else:
                    terminar(*dato)
        except queue.Empty:
            pass
        if estado["corriendo"]:
            segundos = int(time.time() - estado["inicio"])
            var_tiempo.set(f"{segundos // 60:02d}:{segundos % 60:02d}")
        root.after(100, bombear)

    def terminar(resultado, error):
        estado["corriendo"] = False
        boton.config(state="normal")
        if error is not None:
            escribir(f"ERROR: {error}")
            mostrar_error("No se pudo recalcular", str(error), parent=root)
            return
        salida, csv = resultado
        listos = "\n".join(str(r) for r in (salida, csv) if r)
        messagebox.showinfo("Listo", f"Se genero:\n{listos}", parent=root)

    def ejecutar():
        ruta = var_ruta.get().strip()
        if not ruta:
            messagebox.showwarning("Falta el archivo", "Elige el archivo 11_ con Examinar.")
            return
        secciones = [s for s, v in vars_hojas.items() if v.get()]
        if not secciones and not var_csv.get():
            messagebox.showwarning("Nada seleccionado", "Tilda al menos una hoja o el CSV.")
            return
        estado.update(corriendo=True, inicio=time.time())
        boton.config(state="disabled")
        texto.delete("1.0", "end")

        def trabajo():
            try:
                resultado = recalcular(
                    ruta, secciones, con_csv=var_csv.get(),
                    registrar=lambda m: cola.put(("log", m)),
                )
                cola.put(("fin", (resultado, None)))
            except Exception as error:  # noqa: BLE001 - se muestra
                cola.put(("fin", (None, error)))

        threading.Thread(target=trabajo, daemon=True).start()

    boton.config(command=ejecutar)
    if not var_ruta.get():
        escribir("Elige el archivo 11_PAGOS_BESS_AAMM.xlsx con 'Examinar...'.")
    else:
        escribir("Revisa lo tildado y aprieta 'Ejecutar'.")
    bombear()
    root.mainloop()


def ejecutar_desde(carpeta_del_codigo, argumentos=None):
    """
    Punto de entrada del archivo de antecedentes. Sin argumentos abre la
    ventana, con el 11_ de la carpeta de arriba (o de la misma) ya elegido.
    'ruta.xlsx --sin-ventana' recalcula todo sin ventana (lo usan las
    pruebas).
    """

    argumentos = list(sys.argv[1:] if argumentos is None else argumentos)
    carpeta = Path(carpeta_del_codigo)
    if "--sin-ventana" in argumentos:
        argumentos.remove("--sin-ventana")
        ruta = Path(argumentos[0]) if argumentos else buscar_11([carpeta.parent, carpeta])
        if ruta is None:
            raise SystemExit("Indica el archivo 11_PAGOS_BESS_AAMM.xlsx.")
        recalcular(ruta, [s for s, _ in SECCIONES_RECALCULO], con_csv=True)
        return
    main(Path(argumentos[0]) if argumentos else buscar_11([carpeta.parent, carpeta]))


if __name__ == "__main__":
    ejecutar_desde(Path.cwd())
