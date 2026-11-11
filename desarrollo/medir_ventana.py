# -*- coding: utf-8 -*-
"""
Mide lo que tarda la ventana de Balance_BESS.py: importar, revisar la
carpeta del caso, repintar el arbol y abrir "Ejecutar todo".

Es la vara del plan desarrollo/PLAN_OPTIMIZACION_VENTANAS_BESS.md: se
corre ANTES y DESPUES de cada fase y se comparan los numeros.

    python desarrollo/medir_ventana.py                    (caso sintetico)
    python desarrollo/medir_ventana.py --caso "D:/..." --aamm 2609
    python desarrollo/medir_ventana.py --latencia-ms 3    (simula unidad de red)

En Linux la parte de la ventana necesita pantalla: xvfb-run -a python ...
Sin tkinter o sin pantalla esa parte se saltea y el resto se mide igual.

No toca la configuracion real (usa una carpeta temporal para __config__)
y sobre un caso real solo mira: no genera ni borra nada. OJO: revisar una
carpeta crea "Subastas/DB subastas/" si falta (lo hace el programa).
"""

import argparse
import builtins
import collections
import os
import sys
import tempfile
import time
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(RAIZ))

LLAMADAS = ("stat", "lstat", "scandir", "listdir", "mkdir")


def _crono():
    inicio = time.perf_counter()
    return lambda: time.perf_counter() - inicio


def _caso_sintetico(nucleo, carpeta, aamm):
    """Un caso completo y chico: todas las entradas y los Access del mes."""

    import pandas as pd
    from Script.Subastas import Ofertas_Adjudicadas as ofertas_adj

    caso = Path(carpeta) / f"Caso {aamm}"
    for sub in nucleo.SUBCARPETAS_CASO:
        (caso / sub).mkdir(parents=True, exist_ok=True)

    libro = caso / "Auxiliares" / "Centrales.xlsx"
    pd.DataFrame({"a": [1]}).to_excel(libro, sheet_name="Resumen BESS", index=False)
    contenido = libro.read_bytes()
    for relativo in (
        "Medidas/Medidas_SAE.xlsx", f"Medidas/SOC_{aamm}.xlsx",
        "Auxiliares/Homologacion.xlsx", "Ofertas/OfertasSSCC.xlsx",
        "Cmg/cmg.xlsx", f"FD y FMA/fma_cpf_{aamm}.xlsx",
        f"FD y FMA/fma_csf_{aamm}.xlsx",
    ):
        if (caso / relativo).parent.is_dir():
            (caso / relativo).write_bytes(contenido)

    with pd.ExcelWriter(caso / nucleo.nombre_salida(aamm)) as escritor:
        for hoja in nucleo.ORDEN_HOJAS_SALIDA:
            pd.DataFrame({"a": range(5)}).to_excel(
                escritor, sheet_name=hoja, index=False
            )

    # El PO de cada dia y un PID: lo que hay en un mes real son decenas,
    # pero el programa PREGUNTA por los 24 x dias posibles.
    for dia, hora, nombre in ofertas_adj.nombres_del_periodo(aamm):
        if hora in (0, 14):
            (caso / "Subastas" / "DB subastas" / nombre).write_bytes(b"x")

    return caso


class _Contador:
    """Cuenta (y si se pide, demora) las llamadas al sistema de archivos."""

    def __init__(self):
        self.por_origen = collections.Counter()
        self.total = 0
        self.latencia = 0.0
        self._originales = {}

    def poner(self):
        for nombre in LLAMADAS:
            self._originales[(os, nombre)] = getattr(os, nombre)
            setattr(os, nombre, self._envolver(getattr(os, nombre)))
        self._originales[(builtins, "open")] = builtins.open
        builtins.open = self._envolver(builtins.open)

    def sacar(self):
        for (modulo, nombre), original in self._originales.items():
            setattr(modulo, nombre, original)
        self._originales.clear()

    def reiniciar(self):
        self.por_origen.clear()
        self.total = 0

    def _envolver(self, original):
        def envuelta(*args, **kwargs):
            self.total += 1
            marco = sys._getframe(1)
            while marco is not None and "Script" not in Path(
                marco.f_code.co_filename
            ).parts:
                marco = marco.f_back
            if marco is not None:
                self.por_origen[
                    f"{Path(marco.f_code.co_filename).name}:"
                    f"{marco.f_code.co_name}"
                ] += 1
            if self.latencia:
                time.sleep(self.latencia)
            return original(*args, **kwargs)
        return envuelta


def _medir_ventana(caso, aamm):
    """Arranque, repintado y dialogo, con la ventana de verdad."""

    try:
        import tkinter as tk
        from unittest import mock
        import Balance_BESS as app
    except ImportError as error:
        print(f"  (sin tkinter: {error})")
        return
    if os.name != "nt" and not os.environ.get("DISPLAY"):
        print("  (sin pantalla: correr con xvfb-run -a)")
        return

    cfg = {
        "carpeta_base": str(caso), "aamm": aamm,
        app.CLAVE_CARPETAS_PERIODO: {aamm: str(caso)},
    }
    ventana = {}
    original = tk.Tk.mainloop
    tk.Tk.mainloop = lambda self: ventana.setdefault("root", self)

    def todos(widget):
        for hijo in widget.winfo_children():
            yield hijo
            yield from todos(hijo)

    def boton(texto, dentro):
        return next(
            w for w in todos(dentro)
            if w.winfo_class() == "Button" and w.cget("text") == texto
        )

    try:
        with mock.patch.object(app, "leer_config", lambda: dict(cfg)), \
                mock.patch.object(app, "guardar_config", lambda datos: None):
            crono = _crono()
            app.main()
            root = ventana["root"]
            root.update()
            print(f"  main() + primer dibujo:            {crono() * 1000:7.0f} ms")
            print(f"  widgets en la ventana principal:   {sum(1 for _ in todos(root)):7d}")

            marco = next(
                w for w in todos(root)
                if w.winfo_class() == "Labelframe"
                and w.cget("text") == "Carpeta base del caso"
            )
            actualizar = boton("Actualizar", marco)
            tiempos = []
            for _ in range(5):
                crono = _crono()
                actualizar.invoke()
                root.update()
                tiempos.append(crono())
            print(f"  boton Actualizar (revisar + pintar): {min(tiempos) * 1000:5.0f} ms")

            crono = _crono()
            boton("Ejecutar todo", root).invoke()
            root.update()
            print(f"  abrir 'Ejecutar todo':             {crono() * 1000:7.0f} ms")
            root.destroy()
    finally:
        tk.Tk.mainloop = original


def main():
    parser = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    parser.add_argument("--caso", help="carpeta de un caso real (solo se mira)")
    parser.add_argument("--aamm", default="2609")
    parser.add_argument(
        "--latencia-ms", type=float, default=3.0,
        help="demora simulada por llamada al disco (unidad de red)",
    )
    args = parser.parse_args()

    temporal = tempfile.TemporaryDirectory()
    os.environ.setdefault("BESS_CARPETA_CONFIG", str(Path(temporal.name) / "cfg"))

    print("IMPORTAR")
    crono = _crono()
    import tkinter  # noqa: F401 - se mide aparte de lo demas
    print(f"  tkinter:                           {crono() * 1000:7.0f} ms")
    crono = _crono()
    from Script import nucleo
    print(f"  Script.nucleo (pandas, openpyxl):  {crono() * 1000:7.0f} ms")
    print("  (detalle: python -X importtime -c \"import Balance_BESS\")")

    caso = (
        Path(args.caso) if args.caso
        else _caso_sintetico(nucleo, temporal.name, args.aamm)
    )
    print(f"\nREVISAR LA CARPETA  ({caso})")

    crono = _crono()
    _, filas = nucleo.revisar_estructura(str(caso), args.aamm)
    print(f"  primera vez (abre los libros):     {crono() * 1000:7.0f} ms   {len(filas)} filas")
    crono = _crono()
    nucleo.revisar_estructura(str(caso), args.aamm)
    print(f"  siguientes (lo que cuesta cada repintado): {crono() * 1000:.0f} ms")

    contador = _Contador()
    contador.poner()
    try:
        nucleo.revisar_estructura(str(caso), args.aamm)
        print(f"  llamadas al sistema de archivos:   {contador.total:7d}")
        for origen, veces in contador.por_origen.most_common(8):
            print(f"      {veces:5d}  {origen}")
        contador.reiniciar()
        contador.latencia = args.latencia_ms / 1000.0
        crono = _crono()
        nucleo.revisar_estructura(str(caso), args.aamm)
        print(
            f"  con {args.latencia_ms:g} ms por llamada (unidad de red simulada): "
            f"{crono():.2f} s de ventana congelada por repintado"
        )
    finally:
        contador.sacar()

    print("\nVENTANA")
    _medir_ventana(caso, args.aamm)
    temporal.cleanup()


if __name__ == "__main__":
    main()
