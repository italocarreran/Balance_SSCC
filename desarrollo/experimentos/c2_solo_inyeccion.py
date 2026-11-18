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


def main(ruta_planilla, trabajo=None):
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
            print(f"Recalculando {nombre} (C2 solo inyeccion = {regla})...")
            nucleo.generar_pagos_bess(carpeta, set(SECCIONES), registrar=callado, aamm=aamm)
            print(f"  listo en {time.time() - inicio:.0f} s")
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

    print()
    print("RE545 por central, hoy vs. contando solo la inyeccion (mismo codigo):")
    for central, fila in tabla.iterrows():
        if abs(fila["Diferencia [CLP]"]) >= 1:
            print(f"  {central}: {fila['RE545 hoy [CLP]']:>15,.0f} -> "
                  f"{fila['RE545 propuesta [CLP]']:>15,.0f}  ({fila['Diferencia [CLP]']:+,.0f})")
    print(f"\n  TOTAL RE545: {tabla['RE545 hoy [CLP]'].sum():,.0f} -> "
          f"{tabla['RE545 propuesta [CLP]'].sum():,.0f}  ({tabla['Diferencia [CLP]'].sum():+,.0f})")
    print(f"\nReporte: {final}")
    print(f"Tabla:   {trabajo / 'Diferencia_RE545_por_central.xlsx'}")
    return final


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    main(*sys.argv[1:3])
