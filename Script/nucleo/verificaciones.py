# -*- coding: utf-8 -*-
"""
Verificaciones de coherencia de la hoja Medidores antes de pagar.

Nacen de una observacion a la planilla 11_PAGOS_BESS_2610 sobre
SAE-CRCA-PFV-ELENA, que tenia dos problemas que el calculo no avisaba:

1. SOC-019 -- Carga y descarga que no calzan con el SoC.
   La energia que se toma como descarga (Gen_Unidad > 0) coincidia con
   SUBIDAS del SoC y la que se toma como carga (Gen_Unidad < 0) con
   BAJADAS. En una CRCA, mientras la medida del PFV y la del BESS no
   esten separadas, los flujos del PFV se mezclan con los del BESS: no
   se puede saber cuanto cargo y descargo el BESS, y las curvas del
   costo de oportunidad (y la compensacion) no corresponden a su
   operacion real.

2. CMG-007 -- Central con barra que no esta en la hoja CMg.
   La hoja Centrales le asignaba KIMAL_________220, pero la hoja CMg no
   traia esa barra: todos sus bloques quedan sin CMg, C1 y C2 dan 0 y
   la compensacion por costo de oportunidad queda en 0 sin que nada lo
   diga (CMG-004 solo cuenta filas, no dice de quien ni por que).

Las dos son ALTA: la corrida queda NO APROBADA - REQUIERE REVISION,
pero el calculo sigue (no se cambia ningun numero).

Despues (2026-11-16 (2)), al ver que habia centrales con el signo
cambiado que no salian en el reporte:

3. SOC-020 (MEDIA) -- la central no se pudo verificar (sin SoC, pocos
   bloques): antes se la saltaba en silencio.
4. SOC-021 -- sin SoC, una central con PFV se revisa por el mediodia:
   con el PFV generando el neto tiene que salir como inyeccion.

`signo_por_central` es la base de las alertas y de la columna "SoC vs
medida" de la matriz de centrales (Control_corrida.xlsx).
"""

import pandas as pd

from .alertas import ALTA, MEDIA, Alerta, anotar
from .diccionarios import _normaliza_cuarto
from .utiles import normalizar


# ------------------------------------------------------------
# SOC-019: el signo de la energia contra el del cambio de SoC
# ------------------------------------------------------------

# Bloques minimos con energia y cambio de SoC para opinar: con menos,
# un par de bloques raros decidirian solos.
MINIMO_BLOQUES_SOC = 20

# Sobre los bloques que se pueden evaluar, desde que fraccion
# contraria se alerta. Una BESS sana da casi 0 %; ELENA daba "gran
# parte". Con mas de la mitad al reves los canales no pueden estar
# bien.
UMBRAL_CONTRARIOS_ALTA = 0.50

# Entre este umbral y el de arriba va como MEDIA: hay una mezcla
# importante (tipico de una CRCA con PFV y BESS medidos juntos).
UMBRAL_CONTRARIOS_MEDIA = 0.25

# Energia de un bloque por debajo de esta fraccion de la mayor de la
# central: se toma como ruido y no se evalua.
FRACCION_ENERGIA_MINIMA = 0.01


def _orden_temporal(sub):
    """Ordena las filas de una central en el tiempo."""

    if "intervalo" in sub.columns:
        tiempo = pd.to_datetime(sub["intervalo"], errors="coerce")
        if tiempo.notna().all():
            return sub.assign(_t=tiempo).sort_values("_t", kind="stable")

    if "Cuarto de Hora" in sub.columns:
        cuarto = pd.to_numeric(sub["Cuarto de Hora"], errors="coerce")
        if cuarto.notna().all():
            return sub.assign(_t=cuarto).sort_values("_t", kind="stable")

    columnas = [c for c in ("Dia", "Hora", "Minutos") if c in sub.columns]
    if columnas:
        return sub.sort_values(columnas, kind="stable")

    return sub


def _contar_signos(energia, delta_soc):
    """
    (coherentes, contrarios) entre la energia de cada bloque y el cambio
    de SoC asociado. Descarga (E > 0) debe bajar el SoC; carga (E < 0)
    debe subirlo.
    """

    validos = energia.notna() & delta_soc.notna() & delta_soc.ne(0)
    e = energia[validos]
    d = delta_soc[validos]

    coherentes = int(((e > 0) & (d < 0)).sum() + ((e < 0) & (d > 0)).sum())
    contrarios = int(((e > 0) & (d > 0)).sum() + ((e < 0) & (d < 0)).sum())
    return coherentes, contrarios


def evaluar_soc_central(sub):
    """
    Cuenta, para una central, los bloques donde el signo de Gen_Unidad
    contradice el cambio de SoC.

    No se sabe con certeza si el SoC de un cuarto es el del inicio o el
    del final del bloque: se prueban las dos lecturas (SoC siguiente -
    actual y actual - anterior) y se queda con la MAS favorable. Asi la
    alerta solo sale cuando ninguna de las dos explica la medida.

    Devuelve SIEMPRE un dict: evaluados, coherentes, contrarios,
    fraccion (contrarios / evaluados, o None si no se pudo opinar) y
    motivo (por que no se pudo, o "").
    """

    vacio = dict(evaluados=0, coherentes=0, contrarios=0, fraccion=None)

    sub = _orden_temporal(sub)

    energia = pd.to_numeric(sub["Gen_Unidad"], errors="coerce")
    soc = pd.to_numeric(sub["SoC"], errors="coerce")

    if soc.notna().sum() == 0:
        return dict(vacio, motivo="sin SoC")

    maximo = energia.abs().max()
    if pd.isna(maximo) or maximo == 0:
        return dict(vacio, motivo="sin energia medida")

    energia = energia.where(energia.abs() > maximo * FRACCION_ENERGIA_MINIMA)

    mejor = None
    mas_evaluados = 0
    for delta in (soc.shift(-1) - soc, soc - soc.shift(1)):
        coherentes, contrarios = _contar_signos(energia, delta)
        evaluados = coherentes + contrarios
        mas_evaluados = max(mas_evaluados, evaluados)
        if evaluados < MINIMO_BLOQUES_SOC:
            continue
        fraccion = contrarios / evaluados
        if mejor is None or fraccion < mejor["fraccion"]:
            mejor = dict(
                evaluados=evaluados, contrarios=contrarios,
                coherentes=coherentes, fraccion=fraccion, motivo="",
            )

    if mejor is None:
        return dict(
            vacio, evaluados=mas_evaluados,
            motivo=(
                f"solo {mas_evaluados} bloque(s) con energia y cambio de "
                f"SoC (minimo {MINIMO_BLOQUES_SOC})"
            ),
        )

    return mejor


def diagnosticar_soc_central(sub):
    """Como evaluar_soc_central, pero None si no se pudo opinar."""

    resultado = evaluar_soc_central(sub)
    return None if resultado["fraccion"] is None else resultado


# ------------------------------------------------------------
# SOC-021: sin SoC, una central con PFV se revisa por el mediodia
# ------------------------------------------------------------

# Horas locales (inicio del cuarto) en que el PFV genera de lleno. En
# una CRCA con el PFV y el BESS medidos juntos el neto del medidor a
# esa hora es inyeccion: si sale mayoritariamente negativo, el signo
# esta dado vuelta.
HORAS_MEDIODIA = range(11, 16)


def evaluar_mediodia_pfv(central, sub):
    """
    Para una central con PFV ("PFV" en el nombre): de los bloques de
    11:00 a 15:59 (hora local de 'intervalo') con energia, que fraccion
    es inyeccion (Gen_Unidad > 0). Devuelve dict con evaluados,
    fraccion_inyeccion (o None) y motivo.
    """

    vacio = dict(evaluados=0, fraccion_inyeccion=None)

    if "PFV" not in str(central).upper():
        return dict(vacio, motivo="no tiene PFV")
    if "intervalo" not in sub.columns:
        return dict(vacio, motivo="sin hora local (intervalo)")

    hora = pd.to_datetime(sub["intervalo"], errors="coerce").dt.hour
    energia = pd.to_numeric(sub["Gen_Unidad"], errors="coerce")

    maximo = energia.abs().max()
    if pd.isna(maximo) or maximo == 0:
        return dict(vacio, motivo="sin energia medida")

    util = hora.isin(HORAS_MEDIODIA) & (
        energia.abs() > maximo * FRACCION_ENERGIA_MINIMA
    )
    evaluados = int(util.sum())
    if evaluados < MINIMO_BLOQUES_SOC:
        return dict(
            vacio, evaluados=evaluados,
            motivo=f"solo {evaluados} bloque(s) de mediodia con energia",
        )

    inyeccion = int((energia[util] > 0).sum())
    return dict(
        evaluados=evaluados, fraccion_inyeccion=inyeccion / evaluados,
        motivo="",
    )


# ------------------------------------------------------------
# Lo que se sabe del signo de cada central (alertas y matriz)
# ------------------------------------------------------------

SOC = "SoC"
MEDIODIA = "mediodia PFV"


def signo_por_central(df_medidores):
    """
    central (como viene en Medidores) -> dict con:
        metodo        SOC, MEDIODIA o None (no se pudo verificar)
        coincidencia  fraccion de bloques que calzan (0..1) o None
        evaluados     bloques usados
        motivo        por que no se pudo (si metodo es None) o por que
                      no se uso el SoC (si metodo es MEDIODIA)

    Con SoC manda el SoC; sin el, una central con PFV se revisa por el
    mediodia. Vacio si a Medidores le faltan columnas.
    """

    columnas = {"clave", "Gen_Unidad", "SoC"}
    if df_medidores is None or not columnas.issubset(df_medidores.columns):
        return {}

    resultado = {}
    for central, sub in df_medidores.groupby("clave", sort=True):

        if not str(central).strip():
            continue

        soc = evaluar_soc_central(sub)
        if soc["fraccion"] is not None:
            resultado[central] = dict(
                metodo=SOC, coincidencia=1 - soc["fraccion"],
                evaluados=soc["evaluados"], contrarios=soc["contrarios"],
                motivo="",
            )
            continue

        sol = evaluar_mediodia_pfv(central, sub)
        if sol["fraccion_inyeccion"] is not None:
            resultado[central] = dict(
                metodo=MEDIODIA, coincidencia=sol["fraccion_inyeccion"],
                evaluados=sol["evaluados"],
                contrarios=sol["evaluados"] - round(
                    sol["fraccion_inyeccion"] * sol["evaluados"]
                ),
                motivo=soc["motivo"],
            )
            continue

        motivo = soc["motivo"]
        if "PFV" in str(central).upper():
            motivo += f"; mediodia PFV: {sol['motivo']}"
        resultado[central] = dict(
            metodo=None, coincidencia=None, evaluados=0, contrarios=0,
            motivo=motivo,
        )

    return resultado


def _severidad(coincidencia):
    """ALTA si calza menos de la mitad, MEDIA si menos de 3/4."""

    if coincidencia <= 1 - UMBRAL_CONTRARIOS_ALTA:
        return ALTA
    if coincidencia <= 1 - UMBRAL_CONTRARIOS_MEDIA:
        return MEDIA
    return None


def verificar_soc_vs_energia(df_medidores, registrar=print, hoja="Medidores"):
    """
    SOC-019 / SOC-020 / SOC-021: por central, que la carga y la descarga
    de Medidores vayan en el sentido correcto. Devuelve las alertas.

      SOC-019  con SoC: descarga con subidas de SoC o carga con bajadas.
      SOC-021  sin SoC, central con PFV: al mediodia el neto sale como
               retiro (el PFV deberia estar inyectando).
      SOC-020  no se pudo verificar de ninguna forma (MEDIA): que no
               haya alerta de signo NO quiere decir que este bien.
    """

    alertas = []
    origen = "CONTROL NUEVO (observacion 11_PAGOS_BESS_2610)"
    accion = (
        "revisar canal y Flujo de los puntos de la central en la hoja "
        "'homol' del archivo de homologacion; en una CRCA, separar la "
        "medida del PFV de la del BESS. La compensacion de esta central "
        "no es confiable"
    )

    for central, dato in signo_por_central(df_medidores).items():

        if dato["metodo"] is None:
            alertas.append(anotar(registrar, Alerta(
                "SOC-020", MEDIA, "Coherencia Medidores",
                f"No se pudo verificar el signo de carga/descarga "
                f"({dato['motivo']}). Que no salga SOC-019 no quiere decir "
                f"que el signo este bien.",
                central=str(central),
                valor_encontrado=dato["motivo"],
                accion="revisar a mano el Flujo de la central en 'homol' o "
                       "traer su SoC",
                hoja=hoja, origen_control=origen,
            )))
            continue

        coincidencia = dato["coincidencia"]
        severidad = _severidad(coincidencia)
        if severidad is None:
            continue

        contrario = 1 - coincidencia

        if dato["metodo"] == SOC:
            es_crca = "CRCA" in str(central).upper()
            detalle_crca = (
                " Es una CRCA: mientras las medidas del PFV y del BESS no "
                "esten separadas, los flujos del PFV se mezclan con los "
                "del BESS."
                if es_crca else ""
            )
            alertas.append(anotar(registrar, Alerta(
                "SOC-019", severidad, "Coherencia Medidores",
                f"En {dato['contrarios']:,} de {dato['evaluados']:,} "
                f"bloques ({contrario:.0%}) la energia tomada como "
                f"descarga coincide con subidas del SoC o la tomada como "
                f"carga con bajadas.{detalle_crca} La energia de "
                f"carga/descarga y las curvas del costo de oportunidad no "
                f"representan la operacion real del BESS.",
                central=str(central),
                valor_encontrado=f"{contrario:.1%} de bloques contrarios",
                valor_esperado=f"< {UMBRAL_CONTRARIOS_MEDIA:.0%}",
                accion=accion, hoja=hoja, origen_control=origen,
            )))
        else:
            alertas.append(anotar(registrar, Alerta(
                "SOC-021", severidad, "Coherencia Medidores",
                f"Sin SoC para verificar ({dato['motivo']}). Al mediodia "
                f"(11:00-15:59), con el PFV generando, {dato['contrarios']:,} "
                f"de {dato['evaluados']:,} bloques ({contrario:.0%}) salen "
                f"como retiro: el signo de la medida esta probablemente "
                f"invertido.",
                central=str(central),
                valor_encontrado=f"{coincidencia:.1%} de bloques inyectando",
                valor_esperado=f"> {1 - UMBRAL_CONTRARIOS_MEDIA:.0%}",
                accion=accion, hoja=hoja, origen_control=origen,
            )))

    return alertas


# ------------------------------------------------------------
# CMG-007: la barra de la central tiene que estar en la hoja CMg
# ------------------------------------------------------------

def _barras_y_claves_cmg(dic_cmg):
    barras = set()
    for clave in dic_cmg:
        barras.add(str(clave).split("|", 1)[0])
    return barras


def verificar_barras_con_cmg(
    df_medidores, mapa_barra, dic_cmg, registrar=print,
    hoja_centrales="Centrales", hoja_cmg="CMg",
):
    """
    CMG-007: cada central de Medidores con barra asignada debe tener
    CMg para al menos uno de sus bloques. Si la barra no esta en la hoja
    CMg (o esta pero ningun cuarto cruza), todos sus bloques quedan sin
    CMg, C1 y C2 dan 0 y la compensacion queda en 0.

    Devuelve la lista de alertas.
    """

    alertas = []

    if df_medidores is None or "clave" not in df_medidores.columns:
        return alertas

    barras_cmg = _barras_y_claves_cmg(dic_cmg or {})

    for central, sub in df_medidores.groupby("clave", sort=True):

        if not str(central).strip():
            continue

        barra = mapa_barra.get(normalizar(central), "")
        barra = "" if barra is None else str(barra).strip()
        if not barra:
            continue  # sin barra ya lo avisa MAE-001

        if barra.upper() not in barras_cmg:
            motivo = (
                f"La hoja '{hoja_centrales}' le asigna la barra {barra}, "
                f"pero la hoja '{hoja_cmg}' no trae el costo marginal de "
                f"esa barra"
            )
            encontrado = f"barra {barra} ausente de '{hoja_cmg}'"
        else:
            cuartos = sub["Cuarto de Hora"] if "Cuarto de Hora" in sub else []
            con_cmg = sum(
                1 for cuarto in cuartos
                if f"{barra.upper()}|{_normaliza_cuarto(cuarto)}" in dic_cmg
            )
            if con_cmg:
                continue  # parcial: lo cuenta CMG-004
            motivo = (
                f"La barra {barra} esta en la hoja '{hoja_cmg}', pero "
                f"ninguno de los {len(sub):,} bloques de la central cruza "
                f"por Cuarto de Hora"
            )
            encontrado = f"0 de {len(sub):,} bloques con CMg"

        alertas.append(anotar(registrar, Alerta(
            "CMG-007", ALTA, "Coherencia Centrales/CMg",
            f"{motivo}: los {len(sub):,} bloques del mes quedan sin costo "
            f"marginal, C1 y C2 resultan nulos y la compensacion por costo "
            f"de oportunidad queda en 0.",
            central=str(central),
            clave=barra,
            valor_encontrado=encontrado,
            valor_esperado="CMg para todos los bloques de la central",
            accion=(
                f"agregar la barra {barra} a la hoja '{hoja_cmg}' (o "
                f"corregir la barra de la central en '{hoja_centrales}') "
                f"y recalcular"
            ),
            hoja=hoja_cmg,
            origen_control="CONTROL NUEVO (observacion 11_PAGOS_BESS_2610)",
        )))

    return alertas


def verificar_medidores(
    df_medidores, mapa_barra, dic_cmg, registrar=print,
    hoja_centrales="Centrales",
):
    """Las dos verificaciones juntas. Devuelve todas las alertas."""

    registrar("Verificando coherencia de Medidores (SoC y CMg por central)...")

    alertas = verificar_soc_vs_energia(df_medidores, registrar=registrar)
    alertas += verificar_barras_con_cmg(
        df_medidores, mapa_barra, dic_cmg, registrar=registrar,
        hoja_centrales=hoja_centrales,
    )

    if not alertas:
        registrar("  Sin inconsistencias de signo ni de barras sin CMg.")

    return alertas
