# -*- coding: utf-8 -*-
"""
Comparar los resultados de dos planillas 11_PAGOS_BESS (un mes contra el
anterior, o una version contra otra del mismo mes).

Pedido del usuario (2026-11-10): "un archivo que compare resultados entre
el archivo y el mes o version anterior, los resultados de compensacion
central [...] algo tipo reporte, tal central aumento su ingreso por RE545
dado que las ofertas respecto al otro aumentaron tanto [...] que el
resultado de la ejecucion se guarde en la carpeta Log".

Que hace:
  - por central: lo que recibe por E Costos, por RE545 y en total, en las
    dos planillas, y la diferencia;
  - una frase que explica el cambio con lo que mas se movio en sus
    entradas (lo adjudicado en Subastas, el FD, los dias con oferta
    completa, la energia, el CMg...), en el mismo sentido que la plata;
  - las que no cambiaron van juntas en una sola linea (sin ruido).

El reporte queda en <carpeta de la planilla actual>/Log/
Comparacion_<AAMM actual>_vs_<AAMM anterior>.xlsx, con dos hojas:
"Comparacion" (el resumen y la tabla por central) y "Detalle" (cada
entrada de cada central, anterior y actual).

A diferencia de las otras herramientas, esta usa el programa (Script/):
lee las planillas con los mismos lectores que el calculo. Por eso tiene
que estar en la carpeta Herramientas del repositorio.

Uso:  python Herramientas/Comparar_Meses.py
"""

import json
import re
import sys
import threading
from pathlib import Path

# El repositorio (la carpeta de arriba) para importar Script.
RAIZ = Path(__file__).resolve().parents[1]
if str(RAIZ) not in sys.path:
    sys.path.insert(0, str(RAIZ))

PATRON_PLANILLA = "11_PAGOS_BESS_*.xlsx"


# ------------------------------------------------------------
# Rutas recordadas y la planilla anterior sugerida
# ------------------------------------------------------------

def carpeta_config():
    """El __config__ mas cercano hacia arriba, o esta misma carpeta."""

    carpeta = Path(__file__).resolve().parent
    for candidata in (carpeta, *carpeta.parents):
        if (candidata / "__config__").is_dir():
            return candidata / "__config__"
    return carpeta


def ruta_recordadas():
    return carpeta_config() / "comparar_meses.json"


def leer_recordadas():
    try:
        return json.loads(ruta_recordadas().read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}


def guardar_recordadas(actual, anterior):
    try:
        ruta_recordadas().write_text(
            json.dumps({"actual": str(actual), "anterior": str(anterior)},
                       ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
    except OSError:
        pass  # recordar es una comodidad: si no se puede, no pasa nada


def aamm_anterior(aamm):
    """'2601' -> '2512'."""

    anio, mes = int(aamm[:2]), int(aamm[2:])
    if mes == 1:
        return f"{(anio - 1) % 100:02d}12"
    return f"{anio:02d}{mes - 1:02d}"


def sugerir_anterior(ruta_actual):
    """
    La planilla del mes anterior, buscada en las carpetas vecinas a la del
    caso (hasta dos niveles); None si no se encuentra.
    """

    ruta_actual = Path(ruta_actual)
    hallado = re.search(r"(\d{4})$", ruta_actual.stem)
    if not hallado:
        return None
    nombre = f"11_PAGOS_BESS_{aamm_anterior(hallado.group(1))}.xlsx"
    caso = ruta_actual.parent
    for base in (caso, caso.parent, caso.parent.parent):
        for patron in (nombre, f"*/{nombre}", f"*/*/{nombre}"):
            try:
                encontrados = sorted(base.glob(patron))
            except OSError:
                continue
            if encontrados:
                return encontrados[0]
    return None


# ------------------------------------------------------------
# Cuadro de error con boton "Copiar" (el de Script/dialogos.py)
# ------------------------------------------------------------

def mostrar_error(titulo, mensaje):
    try:
        from Script.dialogos import mostrar_error as mostrar
    except Exception:  # noqa: BLE001 - sin el programa, el de tkinter
        from tkinter import messagebox
        messagebox.showerror(titulo, mensaje)
        return
    mostrar(titulo, mensaje)


# ------------------------------------------------------------
# La ventana
# ------------------------------------------------------------

def main():
    import os
    import subprocess
    import tkinter as tk
    from tkinter import filedialog

    root = tk.Tk()
    root.title("Comparar planillas 11_PAGOS_BESS")
    root.geometry("900x620")

    recordadas = leer_recordadas()
    var_actual = tk.StringVar(value=recordadas.get("actual", ""))
    var_anterior = tk.StringVar(value=recordadas.get("anterior", ""))
    var_info = tk.StringVar(value="Elige la planilla actual y la anterior.")
    estado = {"reporte": None}

    def fila_archivo(texto, variable, al_elegir=None):
        marco = tk.LabelFrame(root, text=texto, padx=10, pady=6)
        marco.pack(fill="x", padx=14, pady=(8, 2))
        tk.Label(
            marco, textvariable=variable, anchor="w", justify="left",
            wraplength=720, font=("Segoe UI", 9),
        ).pack(side="left", fill="x", expand=True)

        def examinar():
            inicial = Path(variable.get()).parent if variable.get() else None
            ruta = filedialog.askopenfilename(
                title=texto, initialdir=str(inicial) if inicial else None,
                filetypes=[("Planilla 11_PAGOS_BESS", PATRON_PLANILLA),
                           ("Excel", "*.xlsx"), ("Todos", "*.*")],
            )
            if ruta:
                variable.set(ruta)
                if al_elegir:
                    al_elegir(ruta)

        boton = tk.Button(marco, text="Examinar...", command=examinar)
        boton.pack(side="right")
        return boton

    def elegida_actual(ruta):
        sugerida = sugerir_anterior(ruta)
        if sugerida is not None:
            var_anterior.set(str(sugerida))
            var_info.set(f"Anterior sugerida: {sugerida.name} (cámbiala si no es).")

    botones = [
        fila_archivo("Planilla actual", var_actual, elegida_actual),
        fila_archivo("Planilla anterior (otro mes u otra versión)", var_anterior),
    ]

    abajo = tk.Frame(root)
    abajo.pack(fill="x", padx=14, pady=8)
    boton_comparar = tk.Button(
        abajo, text="Comparar", bg="#1f5fa8", fg="white", width=14,
        font=("Segoe UI", 9, "bold"),
    )
    boton_comparar.pack(side="left")
    boton_abrir = tk.Button(abajo, text="Abrir reporte", width=14, state="disabled")
    boton_abrir.pack(side="left", padx=8)
    tk.Label(abajo, textvariable=var_info, anchor="w", fg="#444444").pack(
        side="left", padx=8
    )

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

    def comparar():
        actual, anterior = var_actual.get().strip(), var_anterior.get().strip()
        if not actual or not anterior:
            var_info.set("Faltan las dos planillas.")
            return
        if Path(actual).resolve() == Path(anterior).resolve():
            var_info.set("Elegiste la misma planilla dos veces.")
            return
        guardar_recordadas(actual, anterior)
        caja.delete("1.0", "end")
        for boton in (*botones, boton_comparar):
            boton.config(state="disabled")
        var_info.set("Comparando...")

        def trabajo():
            from Script.nucleo import reporte

            avisos = []
            try:
                destino, lineas, tabla = reporte.comparar_planillas(
                    actual, anterior, registrar=avisos.append
                )
                error = None
            except Exception as exc:  # noqa: BLE001 - se muestra al usuario
                destino = lineas = tabla = None
                error = exc
            root.after(0, lambda: terminado(destino, lineas, tabla, avisos, error))

        threading.Thread(target=trabajo, daemon=True).start()

    def terminado(destino, lineas, tabla, avisos, error):
        for boton in (*botones, boton_comparar):
            boton.config(state="normal")
        if error is not None:
            var_info.set("No se pudo comparar.")
            mostrar_error("No se pudo comparar", str(error))
            return
        for linea in lineas:
            escribir(linea)
        escribir("")
        for _, fila in tabla.iterrows():
            if fila["Explicación"]:
                escribir(f"• {fila['Central']}: {fila['Explicación']}")
        avisos = [a.strip() for a in avisos if "[AVISO]" in a]
        if avisos:
            escribir("")
            escribir("Avisos (hojas que no se pudieron leer; el reporte sigue sin ellas):")
            for aviso in avisos:
                escribir(f"  {aviso}")
        escribir("")
        escribir(f"Reporte guardado en: {destino}")
        estado["reporte"] = destino
        boton_abrir.config(state="normal")
        var_info.set(f"Listo: {destino.name}")

    def abrir():
        destino = estado["reporte"]
        if destino is None:
            return
        try:
            if sys.platform.startswith("win"):
                os.startfile(destino)  # noqa: S606 - abrir el reporte
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(destino)])
            else:
                subprocess.Popen(["xdg-open", str(destino)])
        except OSError as exc:
            mostrar_error("No se pudo abrir el reporte", str(exc))

    boton_comparar.config(command=comparar)
    boton_abrir.config(command=abrir)
    root.mainloop()


if __name__ == "__main__":
    main()
