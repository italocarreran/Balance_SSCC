# -*- coding: utf-8 -*-
"""
EXPERIMENTO: ¿cuanto cambia lo que se paga si las horas "No participó" del
FD no se pagaran?

Desde 2026-11-11 el escenario POR DIRECCION es la regla del programa
(Script/nucleo/no_participo.py): con el codigo actual la base ya no paga
esas direcciones, asi que el escenario por direccion da ~0 y --estricto mide
solo lo que agrega castigar tambien la direccion que si respondio. Los
resultados de antes de la regla estan en RESULTADOS_sin_no_participo.md.

Pedido del usuario (2026-11-11): "Crea un parentesis en otra rama para ver
que tanto cambia si dejamos de pagar los 'no participo'".

Como se mide (con el motor real del programa, no una estimacion):
  1. Dos copias de la planilla 11_PAGOS_BESS en una carpeta de trabajo:
     BASE (tal cual) y ESCENARIO.
  2. En la hoja FD del ESCENARIO, toda hora cuya Respuesta dice "No
     participó" pasa a FD = 0, en su direccion: "Respuesta CPF+" -> CPF(+),
     "Respuesta CPF-" -> CPF(-), "Respuesta CSF" -> CSF(+) y CSF(-); y el
     Factor de Desempeño de esa fila tambien.
  3. Las dos copias se recalculan con el codigo actual (Calculo E Costos,
     Calculo RE545 y COMPENSACION_CENTRAL). Recalcular tambien la BASE
     separa el efecto del escenario de los cambios del programa desde que
     se genero la planilla.
  4. Se comparan con Comparar_Meses (responsables en pesos): el reporte
     queda en <carpeta de trabajo>/Comparacion_sin_no_participo.xlsx.

La FMA (y su vector de participacion CSF) no se toca: sale de las salidas
de FMA y su hoja ya esta en la planilla.

Uso:
    python desarrollo/experimentos/sin_no_participo.py <11_PAGOS_BESS_AAMM.xlsx> [carpeta] [--estricto]

Sin --estricto: el escenario principal (por direccion). Con --estricto: la
cota de arriba (ver sin_no_participo).
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
from Script.nucleo import reporte  # noqa: E402
from Script.nucleo.hojas_entrada import leer_fd_consolidado  # noqa: E402
from Script.nucleo.utiles import normalizar  # noqa: E402


def _no_participo(valor):
    return normalizar(valor).startswith("no particip")


def sin_no_participo(bloque, servicio, estricto=False):
    """
    Copia del bloque de la hoja FD con FD = 0 en las horas "No participó".

    OJO, los dos calculos usan el FD distinto (factores_subastas.py y
    ecostos_prorratas.py):
      - E Costos, por DIRECCION: columnas CPF(+), CPF(-), CSF(+), CSF(-);
      - RE545, el Factor de Desempeño del SERVICIO (Fd_CPF / Fd_CSF), sin
        direccion.
    Las direcciones van siempre a 0 en la que no participo. El factor del
    servicio: estricto=False (el escenario principal) solo si NINGUNA de
    sus direcciones participo; estricto=True si CUALQUIERA dice "No
    participó" (castiga horas en que la direccion adjudicada si respondio:
    es la cota de arriba). En CSF hay una sola Respuesta: dan lo mismo.

    Devuelve (bloque, filas con alguna direccion "No participó").
    """

    bloque = bloque.copy()
    columnas = list(bloque.columns)
    respuestas = [c for c in columnas if "respuesta" in normalizar(c)]
    factor = next((c for c in columnas if f"Fd_{servicio}" in str(c)), None)
    cambiadas = pd.Series(False, index=bloque.index)
    todas = pd.Series(True, index=bloque.index)
    for respuesta in respuestas:
        texto = normalizar(respuesta)
        if servicio == "CSF":
            destinos = [f"{servicio}(+)", f"{servicio}(-)"]
        elif "cpf+" in texto.replace(" ", ""):
            destinos = [f"{servicio}(+)"]
        else:
            destinos = [f"{servicio}(-)"]
        filas = bloque[respuesta].map(_no_participo)
        for destino in destinos:
            if destino in bloque:
                bloque.loc[filas, destino] = 0.0
        cambiadas |= filas
        todas &= filas
    if factor is not None and respuestas:
        bloque.loc[cambiadas if estricto else todas, factor] = 0.0
    return bloque, int(cambiadas.sum())


def recalcular(carpeta, aamm, registrar):
    nucleo.generar_pagos_bess(
        carpeta, {"ecostos", "re545", "compensacion_central"},
        registrar=registrar, aamm=aamm,
    )


def main(ruta_planilla, trabajo=None, estricto=False):
    ruta_planilla = Path(ruta_planilla).resolve()
    aamm = reporte._periodo(ruta_planilla)
    trabajo = Path(trabajo or ruta_planilla.parent / "Log" / "experimento_sin_no_participo")
    sufijo = "estricto" if estricto else "por_direccion"
    base, escenario = trabajo / "base", trabajo / f"escenario_{sufijo}"
    hacer_base = not (base / ruta_planilla.name).is_file() or not (base / "Log").is_dir()
    for carpeta in ((base, escenario) if hacer_base else (escenario,)):
        if carpeta.exists():
            shutil.rmtree(carpeta)
        carpeta.mkdir(parents=True)
        shutil.copy2(ruta_planilla, carpeta / ruta_planilla.name)

    callado = lambda *_: None  # noqa: E731

    # El escenario: FD = 0 en las horas "No participó".
    copia = escenario / ruta_planilla.name
    csf, cpf = leer_fd_consolidado(copia, registrar=callado)
    csf, cambiadas_csf = sin_no_participo(csf, "CSF", estricto)
    cpf, cambiadas_cpf = sin_no_participo(cpf, "CPF", estricto)
    print(f"Horas 'No participó' que pasan a FD = 0: CPF {cambiadas_cpf:,}, CSF {cambiadas_csf:,}")
    nucleo.escribir_salida(
        None, copia, [], [], df_fd_csf=csf, df_fd_cpf=cpf,
        ruta_existente=copia, hojas_regenerar={nucleo.HOJA_FD},
        registrar=callado,
    )

    for nombre, carpeta in ((("base", base),) if hacer_base else ()) + (("escenario", escenario),):
        inicio = time.time()
        print(f"Recalculando {nombre}...")
        recalcular(carpeta, aamm, callado)
        print(f"  listo en {time.time() - inicio:.0f} s")

    destino, lineas, tabla = reporte.comparar_planillas(
        escenario / ruta_planilla.name, base / ruta_planilla.name, registrar=callado,
    )
    final = trabajo / f"Comparacion_sin_no_participo_{sufijo}.xlsx"
    shutil.move(destino, final)
    print()
    print(f"Escenario {sufijo} (sin pagar 'No participó') contra la base, con el mismo codigo:")
    for linea in lineas:
        print("  " + linea)
    print()
    for _, fila in tabla.iterrows():
        if fila.get("Principal responsable"):
            print(f"  {fila[reporte.C_CENTRAL]}: {reporte.plata_corta(fila['Total anterior [CLP]'])}"
                  f" → {reporte.plata_corta(fila['Total actual [CLP]'])}"
                  f" ({reporte.plata_corta(fila['Diferencia [CLP]'])})")
    print(f"\nReporte: {final}")
    return final


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    argumentos = [a for a in sys.argv[1:] if a != "--estricto"]
    main(*argumentos[:2], estricto="--estricto" in sys.argv)
