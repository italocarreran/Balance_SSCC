# -*- coding: utf-8 -*-
"""
Pasar la carpeta de un mes armada ANTES de la division BESS/CRA a la
estructura nueva (pedido del usuario 2026-11-16: "en herramientas agrega
un codigo para reorganizar las carpetas, porque lo usare un par de veces
no mas").

    antes (todo suelto)                 despues
    <mes>/Medidas/                 ->   <mes>/BESS/Medidas/
    <mes>/Auxiliares/              ->   <mes>/BESS/Auxiliares/
    <mes>/Ofertas/  Subastas/      ->   <mes>/BESS/...
    <mes>/Log/                     ->   <mes>/BESS/Log/
    <mes>/11_PAGOS_BESS_AAMM.xlsx  ->   <mes>/BESS/  (tambien SAE_AAMM.csv
                                          y un Balance_BESS.xlsx de antes)
    <mes>/SSCC_Desempeño/          ->   <mes>/FD y FMA/   (comun)
    <mes>/Energia/ FP/ CO/ SC y CO/->   <mes>/CRA/...
    <mes>/Balance_CRA.xlsx y
      5_REMUNERACION_CRA_AAMM.xlsx ->   <mes>/CRA/
    <mes>/FD/ (la del CRA)         ->   sus archivos a <mes>/FD y FMA/
    <mes>/Cmg/, Prorrata retiros/       se quedan (son comunes)

Si en un mismo mes hay de los dos y Auxiliares/ tiene centrales_cra.xlsx,
ese archivo va a CRA/Auxiliares/ y el resto a BESS/Auxiliares/.

Reglas:
  - primero se muestra el PLAN y se confirma; nada se mueve sin eso;
  - nunca se pisa nada: si el destino ya existe (y no es una carpeta
    vacia), el plan lo marca como conflicto y no se mueve NADA;
  - si al mover algo falla (un archivo abierto en Excel), se para ahi y
    dice que quedo movido y que no: se cierra Excel y se vuelve a correr
    (lo ya movido no se vuelve a tocar);
  - los sellos de __config__ se conservan (se identifican por la carpeta
    del mes, no por la de la planilla).

Usa el programa (Script/) para saber que es de quien, asi que tiene que
estar en la carpeta Herramientas del repositorio.

Uso:
    python Herramientas/Reorganizar_Carpetas.py            (ventana)
    python Herramientas/Reorganizar_Carpetas.py <carpeta>  (consola: muestra
                                                            el plan y pregunta)
"""

import sys
from dataclasses import dataclass
from pathlib import Path

# El repositorio (la carpeta de arriba) para importar Script.
RAIZ = Path(__file__).resolve().parents[1]
if str(RAIZ) not in sys.path:
    sys.path.insert(0, str(RAIZ))

from Script.cra import parametros as pc  # noqa: E402
from Script.cra.rutas import raiz_cra  # noqa: E402
from Script.nucleo import carpetas as cc  # noqa: E402
from Script.nucleo import parametros as pn  # noqa: E402
from Script.nucleo.rutas import raiz_bess  # noqa: E402


@dataclass
class Movimiento:
    origen: Path
    destino: Path
    conflicto: str = ""        # vacio = se puede mover


# ------------------------------------------------------------
# El plan
# ------------------------------------------------------------

def _archivos_sueltos(base, patrones):
    vistos = []
    for patron in patrones:
        for ruta in sorted(base.glob(patron)):
            if ruta.is_file() and ruta not in vistos:
                vistos.append(ruta)
    return vistos


def _carpeta_vacia(ruta):
    try:
        return ruta.is_dir() and not any(ruta.iterdir())
    except OSError:
        return False


def _agregar(plan, origen, destino):
    conflicto = ""
    if destino.exists() and not _carpeta_vacia(destino):
        conflicto = f"ya existe {destino}"
    elif destino.exists() and origen.is_file():
        conflicto = f"ya existe una carpeta {destino}"
    plan.append(Movimiento(origen, destino, conflicto))


def planear(carpeta_mes):
    """
    Los movimientos para dejar la carpeta del mes con la estructura
    nueva. Lista vacia si ya esta (o si no hay nada de antes).
    """

    base = Path(carpeta_mes)
    if not base.is_dir():
        raise ValueError(f"No existe la carpeta {base}")

    bess_antiguo = raiz_bess(base) == base
    cra_antiguo = raiz_cra(base) == base
    bess = base / cc.CARPETA_BESS
    cra = base / cc.CARPETA_CRA
    plan = []

    if bess_antiguo:
        for nombre in (pn.CARPETA_MEDIDAS, pn.CARPETA_OFERTAS,
                       pn.CARPETA_SUBASTAS):
            if (base / nombre).is_dir():
                _agregar(plan, base / nombre, bess / nombre)
        for archivo in _archivos_sueltos(base, (
            f"{pn.PREFIJO_SALIDA}*.xlsx", pn.ARCHIVO_SALIDA_ANTIGUO,
            f"{pn.PREFIJO_CSV_SAE}*.csv",
        )):
            _agregar(plan, archivo, bess / archivo.name)

    if cra_antiguo:
        for nombre in (pc.CARPETA_ENERGIA, pc.CARPETA_FP, pc.CARPETA_CO,
                       pc.CARPETA_SC_CO):
            if (base / nombre).is_dir():
                _agregar(plan, base / nombre, cra / nombre)
        for archivo in _archivos_sueltos(base, (
            f"{pc.PREFIJO_SALIDA}*.xlsx", pc.ARCHIVO_SALIDA_ANTIGUO,
        )):
            _agregar(plan, archivo, cra / archivo.name)

    # Log/: del BESS; si solo el CRA es de antes, del CRA (su registro).
    if (base / pn.CARPETA_LOG).is_dir() and (bess_antiguo or cra_antiguo):
        dueno = bess if bess_antiguo else cra
        _agregar(plan, base / pn.CARPETA_LOG, dueno / pn.CARPETA_LOG)

    # Auxiliares/: de quien sea. Con los dos, centrales_cra.xlsx al CRA
    # y el resto al BESS.
    auxiliares = base / pn.CARPETA_AUXILIARES
    if auxiliares.is_dir() and (bess_antiguo or cra_antiguo):
        if bess_antiguo and cra_antiguo:
            for archivo in sorted(auxiliares.iterdir()):
                de_cra = archivo.name.lower() == pc.ARCHIVO_CENTRALES_CRA.lower()
                destino = (cra if de_cra else bess) / pn.CARPETA_AUXILIARES
                _agregar(plan, archivo, destino / archivo.name)
        else:
            dueno = bess if bess_antiguo else cra
            _agregar(plan, auxiliares, dueno / pn.CARPETA_AUXILIARES)

    # El FD comun: "SSCC_Desempeño" (BESS) se renombra; la FD/ del CRA
    # pasa archivo por archivo (puede haber ya una "FD y FMA").
    fd_nueva = base / cc.CARPETA_FD_FMA
    vieja_bess = base / cc.CARPETA_FD_FMA_ANTIGUA
    if vieja_bess.is_dir():
        _agregar(plan, vieja_bess, fd_nueva)
    fd_cra = base / pc.CARPETA_FD_ANTIGUA
    if cra_antiguo and fd_cra.is_dir():
        for archivo in sorted(fd_cra.iterdir()):
            _agregar(plan, archivo, fd_nueva / archivo.name)

    return plan


def describir(plan, base):
    """El plan en texto, una linea por movimiento."""

    base = Path(base)
    if not plan:
        return "No hay nada que reorganizar: la carpeta ya tiene la estructura nueva."

    def corta(ruta):
        return cc.relativa(ruta, base)

    lineas = []
    for mov in plan:
        marca = "  CONFLICTO: " + mov.conflicto if mov.conflicto else ""
        lineas.append(f"{corta(mov.origen)}  ->  {corta(mov.destino)}{marca}")
    conflictos = sum(1 for m in plan if m.conflicto)
    lineas.append("")
    if conflictos:
        lineas.append(
            f"{conflictos} conflicto(s): no se mueve nada hasta resolverlos "
            f"a mano (el destino ya existe y no se pisa)."
        )
    else:
        lineas.append(f"{len(plan)} movimiento(s). Nada se pisa.")
    return "\n".join(lineas)


# ------------------------------------------------------------
# Mover
# ------------------------------------------------------------

def ejecutar(plan, registrar=print):
    """
    Mueve todo lo del plan. Se niega si hay conflictos. Si algo falla
    (abierto en Excel), se para y levanta OSError con lo que quedo
    hecho; lo ya movido no se deshace (correr de nuevo sigue desde ahi).
    """

    conflictos = [m for m in plan if m.conflicto]
    if conflictos:
        raise ValueError(
            "Hay conflictos; no se movio nada:\n"
            + "\n".join(f"  {m.origen} -> {m.conflicto}" for m in conflictos)
        )

    hechos = []
    for mov in plan:
        try:
            if _carpeta_vacia(mov.destino):
                mov.destino.rmdir()
            mov.destino.parent.mkdir(parents=True, exist_ok=True)
            mov.origen.rename(mov.destino)
        except OSError as error:
            raise OSError(
                f"No se pudo mover {mov.origen} (¿abierto en Excel?): {error}\n"
                f"Quedaron movidos {len(hechos)} de {len(plan)}. Cerra lo que "
                f"este abierto y volve a correr la herramienta: sigue desde ahi."
            ) from error
        hechos.append(mov)
        registrar(f"  movido: {mov.origen.name} -> {mov.destino}")

    # Las que se repartieron archivo por archivo (la FD/ del CRA, una
    # Auxiliares/ de los dos) quedan vacias: se borran (solo si vacias).
    for padre in {mov.origen.parent for mov in hechos}:
        if padre.name in (pc.CARPETA_FD_ANTIGUA, pn.CARPETA_AUXILIARES) \
                and _carpeta_vacia(padre):
            padre.rmdir()

    return hechos


# ------------------------------------------------------------
# Consola y ventana
# ------------------------------------------------------------

def _consola(carpeta):
    plan = planear(carpeta)
    print(describir(plan, carpeta))
    if not plan or any(m.conflicto for m in plan):
        return 0 if not plan else 1
    if input("\n¿Mover? (s/n): ").strip().lower() not in ("s", "si", "sí"):
        print("No se movio nada.")
        return 0
    ejecutar(plan)
    print("Listo.")
    return 0


def main():
    import tkinter as tk
    from tkinter import filedialog, messagebox

    try:
        from Script.dialogos import mostrar_error
    except Exception:  # pragma: no cover
        def mostrar_error(titulo, mensaje):
            messagebox.showerror(titulo, mensaje)

    root = tk.Tk()
    root.title("Reorganizar carpetas (BESS / CRA)")
    root.geometry("860x520")

    var_carpeta = tk.StringVar()
    estado = {"plan": []}

    arriba = tk.Frame(root, padx=10, pady=8)
    arriba.pack(fill="x")
    tk.Label(arriba, text="Carpeta del mes:").pack(side="left")
    tk.Entry(arriba, textvariable=var_carpeta).pack(
        side="left", fill="x", expand=True, padx=6
    )

    texto = tk.Text(root, wrap="none", font=("Consolas", 9))
    texto.pack(fill="both", expand=True, padx=10)

    abajo = tk.Frame(root, padx=10, pady=8)
    abajo.pack(fill="x")

    def mostrar(contenido):
        texto.configure(state="normal")
        texto.delete("1.0", "end")
        texto.insert("end", contenido)
        texto.configure(state="disabled")

    def revisar():
        carpeta = var_carpeta.get().strip()
        if not carpeta:
            return
        try:
            estado["plan"] = planear(carpeta)
        except (OSError, ValueError) as error:
            estado["plan"] = []
            mostrar_error("Reorganizar", str(error))
            return
        mostrar(describir(estado["plan"], carpeta))
        hay = estado["plan"] and not any(m.conflicto for m in estado["plan"])
        boton_mover.configure(state="normal" if hay else "disabled")

    def examinar():
        elegida = filedialog.askdirectory(title="Carpeta del mes")
        if elegida:
            var_carpeta.set(elegida)
            revisar()

    def mover():
        plan = estado["plan"]
        if not plan:
            return
        if not messagebox.askyesno(
            "Reorganizar", f"¿Mover {len(plan)} elemento(s)? Nada se pisa."
        ):
            return
        movidos = []
        try:
            ejecutar(plan, registrar=movidos.append)
        except (OSError, ValueError) as error:
            mostrar_error("Reorganizar", str(error))
        else:
            messagebox.showinfo("Reorganizar", f"Listo: {len(movidos)} movido(s).")
        revisar()

    tk.Button(arriba, text="Examinar…", command=examinar).pack(side="left")
    tk.Button(abajo, text="Revisar", command=revisar).pack(side="left")
    boton_mover = tk.Button(abajo, text="Mover", command=mover, state="disabled")
    boton_mover.pack(side="left", padx=6)
    mostrar("Elegi la carpeta de un mes armada antes de la division BESS/CRA.")

    root.mainloop()


if __name__ == "__main__":
    if len(sys.argv) > 1:
        sys.exit(_consola(sys.argv[1]))
    main()
