# -*- coding: utf-8 -*-
"""
Por que cambio lo que recibe cada central entre dos planillas 11_PAGOS_BESS,
por CAUSAS (el "puente"), recalculando.

Pedido del usuario (2026-11-20): "la idea de esa pagina es responder por que
cambio entre el pre y el def [...] como el precio de una cuerda: la bencina
200 mas cara, el hilo 100 mas caro, la luz 100 mas barata [...] en total
1000". La hoja Responsables (comparador rapido) descompone formulas (C1, C2,
factores) y mete ruido; esto mide cada causa con el calculo de verdad:

  1. la anterior tal cual;
  2. la anterior recalculada con el programa de hoy  -> "Cambios en el calculo";
  3. se le van pasando, de a una, las hojas de entrada de la actual y se
     recalcula cada vez: lo que se mueve en cada paso es lo que aporta esa
     entrada (Medidas, Ofertas, CMg, FD y FMA, Subastas, Centrales);
  4. los reemplazos del 7_ (columna "Reemplazado" y Cuadro N° 2) van aparte:
     los pasos de entrada se calculan sin reemplazos y al final se encienden;
  5. si la actual guardada no da lo mismo que el ultimo paso (se calculo con
     otra version del programa), esa diferencia va como una causa mas.

La suma de las causas da EXACTO la diferencia de cada central (cada paso
parte del anterior). El orden de los pasos importa un poco cuando dos causas
se cruzan (el FD vale mas si hay mas energia): es siempre el mismo
(ORDEN_PASOS) y se dice en el reporte.

Es lento: ~9 recalculos de E Costos y RE545 en una carpeta temporal. Las
planillas no se tocan.
"""

import re
import shutil
import tempfile
import time
from pathlib import Path
from unittest import mock

import pandas as pd

from . import reasignaciones
from .lector import abrir_excel
from .parametros import (
    HOJA_CALCULO_ECOSTOS, HOJA_CALCULO_RE545, HOJA_CENTRALES, HOJA_CMG, HOJA_FD,
    HOJA_FMA, HOJA_MEDIDORES, HOJA_OFERTAS_SSCC, HOJA_SUBASTAS,
)
from .utiles import ErrorEntrada, _texto_seguro

CALCULO = "Cambios en el cálculo (programa de hoy)"
REEMPLAZOS = "Reemplazos (7_Reasignaciones)"
OTRA_VERSION = "La actual se guardó con otra versión del programa"

# (causa, hojas de la actual que se pasan). En este orden.
ORDEN_PASOS = (
    ("Medidas (energía)", (HOJA_MEDIDORES,)),
    ("Ofertas SSCC", (HOJA_OFERTAS_SSCC,)),
    ("CMg", (HOJA_CMG,)),
    ("FD y FMA (desempeño)", (HOJA_FD, HOJA_FMA)),
    ("Subastas adjudicadas", (HOJA_SUBASTAS,)),
    ("Datos de centrales", (HOJA_CENTRALES,)),
)
CAUSAS = (CALCULO, *(c for c, _ in ORDEN_PASOS), REEMPLAZOS, OTRA_VERSION)
LADOS = ("E Costos", "RE545")


def _aamm(ruta):
    hallado = re.search(r"_(\d{4})(?:_|$)", Path(ruta).stem)
    if not hallado:
        raise ErrorEntrada(f"{Path(ruta).name}: no se encontró el AAMM en el nombre.")
    return hallado.group(1)


def montos_por_central(ruta):
    """{central (como se escribe): {"E Costos": $, "RE545": $}} de una planilla."""

    from .ecostos import FORMULAS_E_COSTOS, NOMBRES_CALCULO_E_COSTOS
    from .re545 import FORMULAS_RE545, NOMBRES_CALCULO_RE545
    from .reporte import _hoja_de_calculo, completar_formulas

    resultado = {}
    with abrir_excel(ruta) as libro:
        for lado, hoja, formulas, nombres in (
            ("E Costos", HOJA_CALCULO_ECOSTOS, FORMULAS_E_COSTOS, NOMBRES_CALCULO_E_COSTOS),
            ("RE545", HOJA_CALCULO_RE545, FORMULAS_RE545, NOMBRES_CALCULO_RE545),
        ):
            if hoja not in libro.sheet_names:
                continue
            df = completar_formulas(_hoja_de_calculo(libro, hoja), formulas, nombres)
            if "Configuracion" not in df or "Monto a compensar" not in df:
                continue
            central = df["Configuracion"].map(_texto_seguro)
            monto = pd.to_numeric(df["Monto a compensar"], errors="coerce").fillna(0.0)
            for nombre, valor in monto.groupby(central).sum().items():
                if nombre:
                    resultado.setdefault(nombre, {l: 0.0 for l in LADOS})[lado] = float(valor)
    return resultado


def _recalcular(base, aamm, con_reemplazos):
    from .proceso import generar_pagos_bess
    from .rutas import nombre_salida

    with mock.patch.object(reasignaciones, "QUITAR_REEMPLAZADAS", con_reemplazos), \
            mock.patch.object(reasignaciones, "SUMAR_REEMPLAZANTES", con_reemplazos):
        generar_pagos_bess(base, {"ecostos", "re545"}, registrar=lambda *_: None, aamm=aamm)
    return montos_por_central(base / "BESS" / nombre_salida(aamm))


def _pasar_hojas(destino, origen, hojas):
    """Pone en 'destino' las hojas 'hojas' de 'origen' (las demas no se tocan)."""

    from .xlsx_rapido import nombres_de_hojas, pegar_hojas

    todas = nombres_de_hojas(origen)
    pegar_hojas(destino, origen, ignorar=[h for h in todas if h not in hojas])


def _hojas_iguales(ruta_a, ruta_b, hojas):
    """Las 'hojas' son byte a byte iguales en los dos libros (y estan en los dos)."""

    import zipfile

    from .xlsx_rapido import _hojas, _textos_compartidos

    with zipfile.ZipFile(ruta_a) as a, zipfile.ZipFile(ruta_b) as b:
        partes_a = {h["name"]: h["parte"] for h in _hojas(a)}
        partes_b = {h["name"]: h["parte"] for h in _hojas(b)}
        # Con textos compartidos, el XML de la hoja tiene indices: hay que
        # comparar tambien la tabla de textos.
        textos_iguales = _textos_compartidos(a) == _textos_compartidos(b)
        for hoja in hojas:
            if hoja not in partes_a or hoja not in partes_b:
                return False
            xml_a, xml_b = a.read(partes_a[hoja]), b.read(partes_b[hoja])
            if xml_a != xml_b or (b"t=\"s\"" in xml_a and not textos_iguales):
                return False
    return True


def _tiene_reemplazos(ruta):
    from .lector import abrir_excel as abrir

    try:
        with abrir(ruta) as libro:
            c1, c2 = reasignaciones.leer_subastas_consolidado(libro, HOJA_SUBASTAS)
    except Exception:  # noqa: BLE001 - sin hoja Subastas: sin reemplazos
        return False
    return reasignaciones.COLUMNA_REEMPLAZADO in c1 or (c2 is not None and len(c2) > 0)


def explicar_por_causas(ruta_actual, ruta_anterior, registrar=print, progreso=None):
    """
    (por_central, pasos): por_central = {central: {causa: {lado: $}}} y
    pasos = [(causa, segundos)]. La suma de las causas de una central es lo
    que recibe en la actual menos lo de la anterior.
    """

    from .rutas import nombre_salida

    ruta_actual, ruta_anterior = Path(ruta_actual), Path(ruta_anterior)
    aamm_anterior, aamm_actual = _aamm(ruta_anterior), _aamm(ruta_actual)
    total_pasos = len(ORDEN_PASOS) + 3
    hechos = [0]
    pasos = []

    def avanzar(causa, inicio):
        hechos[0] += 1
        pasos.append((causa, time.time() - inicio))
        registrar(f"  {causa}: listo ({time.time() - inicio:.0f} s)")
        if progreso:
            progreso(100 * hechos[0] / total_pasos)

    registrar(f"Leyendo {ruta_anterior.name} y {ruta_actual.name}...")
    estados = [("inicio", montos_por_central(ruta_anterior))]
    guardada_actual = montos_por_central(ruta_actual)

    with tempfile.TemporaryDirectory() as carpeta:
        base = Path(carpeta)
        (base / "BESS").mkdir()
        aamm = aamm_anterior
        trabajo = base / "BESS" / nombre_salida(aamm)
        shutil.copy2(ruta_anterior, trabajo)

        inicio = time.time()
        registrar(f"Recalculando {ruta_anterior.name} con el programa de hoy...")
        estados.append((CALCULO, _recalcular(base, aamm, True)))
        avanzar(CALCULO, inicio)

        # Lo que no cambia no se recalcula (aporta 0): los reemplazos de una
        # planilla que no los trae, y las hojas iguales en las dos.
        inicio = time.time()
        if _tiene_reemplazos(ruta_anterior):
            registrar("Sin los reemplazos de la anterior...")
            estados.append((REEMPLAZOS, _recalcular(base, aamm, False)))
        else:
            estados.append((REEMPLAZOS, estados[-1][1]))
        avanzar("Reemplazos de la anterior", inicio)

        for causa, hojas in ORDEN_PASOS:
            inicio = time.time()
            cambia_periodo = HOJA_MEDIDORES in hojas and aamm != aamm_actual
            if _hojas_iguales(ruta_anterior, ruta_actual, hojas) and not cambia_periodo:
                registrar(f"{', '.join(hojas)}: igual en las dos, aporta $0.")
                estados.append((causa, estados[-1][1]))
                avanzar(causa, inicio)
                continue
            registrar(f"Pasando {', '.join(hojas)} de {ruta_actual.name}...")
            _pasar_hojas(trabajo, ruta_actual, hojas)
            if cambia_periodo:
                # Desde aca el caso es del periodo de la actual.
                nuevo = base / "BESS" / nombre_salida(aamm_actual)
                trabajo.rename(nuevo)
                trabajo, aamm = nuevo, aamm_actual
            estados.append((causa, _recalcular(base, aamm, False)))
            avanzar(causa, inicio)

        inicio = time.time()
        if _tiene_reemplazos(ruta_actual):
            registrar("Con los reemplazos de la actual...")
            estados.append((REEMPLAZOS, _recalcular(base, aamm, True)))
        else:
            estados.append((REEMPLAZOS, estados[-1][1]))
        avanzar("Reemplazos de la actual", inicio)

    estados.append((OTRA_VERSION, guardada_actual))

    centrales = sorted({c for _, montos in estados for c in montos}, key=str.upper)
    por_central = {}
    for central in centrales:
        causas = {}
        for (_, antes), (causa, despues) in zip(estados, estados[1:]):
            for lado in LADOS:
                delta = (despues.get(central, {}).get(lado, 0.0)
                         - antes.get(central, {}).get(lado, 0.0))
                por_lado = causas.setdefault(causa, {l: 0.0 for l in LADOS})
                por_lado[lado] += delta
        por_central[central] = causas
    return por_central, pasos, estados[0][1], guardada_actual


# ------------------------------------------------------------
# El reporte
# ------------------------------------------------------------

MINIMO_CLP = 1.0
COLUMNA_ARBOL = "Central / causa"


def _total(por_lado):
    return sum(por_lado.values())


def tabla_mes(por_central):
    """Cada causa con lo que aporta en el mes (E Costos, RE545, total) y el total."""

    filas = []
    for causa in CAUSAS:
        ec = sum(c.get(causa, {}).get("E Costos", 0.0) for c in por_central.values())
        re_ = sum(c.get(causa, {}).get("RE545", 0.0) for c in por_central.values())
        if abs(ec + re_) >= MINIMO_CLP or abs(ec) >= MINIMO_CLP or abs(re_) >= MINIMO_CLP:
            filas.append({"Causa": causa, "Aporte [CLP]": ec + re_,
                          "E Costos [CLP]": ec, "RE545 [CLP]": re_})
    tabla = pd.DataFrame(filas, columns=["Causa", "Aporte [CLP]", "E Costos [CLP]", "RE545 [CLP]"])
    if len(tabla):
        tabla.loc[len(tabla)] = [
            "Total (= diferencia del mes)", tabla["Aporte [CLP]"].sum(),
            tabla["E Costos [CLP]"].sum(), tabla["RE545 [CLP]"].sum(),
        ]
    return tabla


def arbol_por_central(por_central, anterior, actual):
    """
    Una central por bloque: "SAE-X · recibe A → B (±d)" y debajo sus causas
    (solo las que mueven algo), de la que mas mueve a la que menos.
    """

    from .reporte import plata_corta

    def suma(montos, central):
        return sum(montos.get(central, {}).values())

    filas, niveles = [], []
    orden = sorted(
        por_central,
        key=lambda c: -abs(sum(_total(v) for v in por_central[c].values())),
    )
    for central in orden:
        causas = por_central[central]
        antes, ahora = suma(anterior, central), suma(actual, central)
        dif = ahora - antes
        filas.append({
            COLUMNA_ARBOL: (f"{central}  ·  recibe {plata_corta(antes)} → {plata_corta(ahora)} "
                            f"({'+' if dif >= 0 else '-'}{plata_corta(abs(dif))})"),
            "Aporte [CLP]": dif,
            "E Costos [CLP]": sum(v["E Costos"] for v in causas.values()),
            "RE545 [CLP]": sum(v["RE545"] for v in causas.values()),
        })
        niveles.append(0)
        lineas = [(causa, v) for causa, v in causas.items()
                  if abs(_total(v)) >= MINIMO_CLP
                  or any(abs(x) >= MINIMO_CLP for x in v.values())]
        if not lineas:
            filas.append({COLUMNA_ARBOL: "sin cambios", "Aporte [CLP]": 0.0,
                          "E Costos [CLP]": 0.0, "RE545 [CLP]": 0.0})
            niveles.append(1)
            continue
        for causa, v in sorted(lineas, key=lambda x: -abs(_total(x[1]))):
            filas.append({COLUMNA_ARBOL: causa, "Aporte [CLP]": _total(v),
                          "E Costos [CLP]": v["E Costos"], "RE545 [CLP]": v["RE545"]})
            niveles.append(1)
    tabla = pd.DataFrame(filas, columns=[COLUMNA_ARBOL, "Aporte [CLP]", "E Costos [CLP]", "RE545 [CLP]"])
    tabla.attrs["niveles"] = niveles
    return tabla


def matriz(por_central):
    """Central x causa (total de los dos lados), con la diferencia al final."""

    filas = []
    for central, causas in sorted(por_central.items(), key=lambda kv: kv[0].upper()):
        fila = {"Central": central}
        for causa in CAUSAS:
            fila[f"{causa} [CLP]"] = _total(causas.get(causa, {}))
        fila["Diferencia [CLP]"] = sum(_total(v) for v in causas.values())
        filas.append(fila)
    tabla = pd.DataFrame(filas)
    if len(tabla):
        vacias = [c for c in tabla.columns[1:-1] if tabla[c].abs().max() < MINIMO_CLP]
        tabla = tabla.drop(columns=vacias)
        total = {c: tabla[c].sum() for c in tabla.columns if c != "Central"}
        total["Central"] = "Total del mes"
        tabla = pd.concat([tabla, pd.DataFrame([total])], ignore_index=True)
    return tabla


def explicar_planillas(ruta_actual, ruta_anterior, registrar=print, progreso=None):
    """
    El reporte lento, aparte del de comparar_planillas: Log/Causas_<actual>_vs_
    <anterior>.xlsx en la carpeta de la actual. Devuelve (ruta, lineas).
    """

    from .parametros import CARPETA_LOG
    from .reporte import _periodo, escribir_reporte, plata_corta

    ruta_actual, ruta_anterior = Path(ruta_actual), Path(ruta_anterior)
    por_central, pasos, anterior, actual = explicar_por_causas(
        ruta_actual, ruta_anterior, registrar=registrar, progreso=progreso
    )
    mes = tabla_mes(por_central)
    lineas = []
    if len(mes):
        total = mes.iloc[-1]["Aporte [CLP]"]
        lineas.append(
            f"Diferencia del mes: {'+' if total >= 0 else '-'}{plata_corta(abs(total))}. "
            f"Por causa: " + ", ".join(
                f"{'+' if f['Aporte [CLP]'] >= 0 else '-'}{plata_corta(abs(f['Aporte [CLP]']))} {f['Causa']}"
                for _, f in mes.iloc[:-1].sort_values(
                    "Aporte [CLP]", key=abs, ascending=False).iterrows()
            ) + "."
        )
    lineas.append(
        "Cómo se mide: se parte de la anterior, se recalcula con el programa de hoy y "
        "se le van pasando las hojas de la actual de a una (Medidores, Ofertas SSCC, "
        "CMg, FD y FMA, Subastas, Centrales), recalculando cada vez; lo que cambia en "
        "cada paso es lo que aporta esa causa. Los reemplazos se miden aparte. La suma "
        "da exacto la diferencia; el orden de los pasos importa un poco cuando dos "
        "causas se cruzan."
    )
    carpeta = ruta_actual.parent / CARPETA_LOG
    carpeta.mkdir(parents=True, exist_ok=True)
    destino = carpeta / f"Causas_{_periodo(ruta_actual)}_vs_{_periodo(ruta_anterior)}.xlsx"
    escribir_reporte(destino, [
        ("Por causa", mes, arbol_por_central(por_central, anterior, actual),
         "Por qué cambió lo que se recibe en el mes, por causa"),
        ("Matriz", lineas, matriz(por_central),
         f"{ruta_actual.name} contra {ruta_anterior.name}"),
    ])
    registrar(f"Reporte: {destino}")
    return destino, lineas
