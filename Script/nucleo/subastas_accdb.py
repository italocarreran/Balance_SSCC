# -*- coding: utf-8 -*-
"""
Subastas desde su origen: los Access OfertasSSCCAdj*.
"""

import pandas as pd

from .externos import ofertas_adj
from .fma import calcular_fd_subastas, calcular_fma_subastas
from .hojas_entrada import (
    NOMBRES_SUBASTAS, _contiene_bess_o_sae_sin_bat,
    _ordenar_subastas_por_hora_mes,
)
from .parametros import (
    ARCHIVO_CENTRALES, CARPETA_FD_FMA, HOJA_RESUMEN_BESS,
)
from .utiles import (
    ErrorEntrada, _entero_a_texto, _texto_seguro, normalizar,
)


# ============================================================
# SUBASTAS DESDE SU ORIGEN REAL (los Access OfertasSSCCAdj*)
#
# La hoja "Subastas" del consolidado se armaba leyendo la hoja "DB" de
# la planilla 3 (construir_subastas, arriba). Pero la planilla 3 no es
# el origen: ella misma se arma pegando la salida de entradas_sscc.py,
# que lee los Access OfertasSSCCAdj*.accdb. Lo que sigue corta ese
# intermediario y arma las MISMAS columnas B:Q directamente desde los
# Access, siguiendo el documento de trazabilidad que entrego el
# usuario (subastas_AAMM.xlsx -> DB!B:K) mas lo que ya sabiamos de
# DB!L:Q.
#
# Equivalencias (columna de Subastas <- de donde sale ahora):
#
#   B Concepto      <- SERVICIO tal cual ("CSF(-)", "CPF(+)", ...)
#   C Control       <- SERVICIO[:3]            (formula real =LEFT(C,3))
#   D Sub_Baj       <- signo de SERVICIO       (=IF(LEFT(RIGHT(C,2),1)="+",...))
#   E Fecha         <- DATE(AÑO, MES, DIA)
#   F Año           <- AÑO
#   G Mes           <- MES
#   H Dia           <- DIA
#   I Hora_dia      <- HORA                    (sin sumar ni restar 1)
#   J Hora_mes      <- (DIA-1)*24 + HORA       (+ ajuste de cambio de hora)
#   K Configuración <- CONFIGURACIÓN
#   L Propietario   <- Centrales.xlsx, "Resumen BESS", columna
#                      "Propietario" (el usuario la agrego en la B y
#                      corrio el resto de la hoja una columna)
#   M Clave horaria <- Configuración & Dia & Hora_dia (igual que antes)
#   N Ciclo         <- vacia (se calcula en Calculo E Costos)
#   O Energía SSCC  <- CANTIDAD PONDERADA MW
#   P FD            <- PENDIENTE (era DB!Y; no existe en el Access)
#   Q FMA           <- PENDIENTE (era DB!V; no existe en el Access)
# ============================================================

# De que columna del Access sale "Energía SSCC" (Subastas!O). El
# Access trae las dos cantidades: la cruda (Quantity -> "CANTIDAD MW")
# y la ponderada (Quantity2 -> "CANTIDAD PONDERADA MW", que
# entradas_sscc.py completa con la cruda cuando viene vacia). Se eligio
# la ponderada porque es la que se remunera; esta en una constante
# para poder cambiarla de un lado si el usuario confirma la otra.
COLUMNA_ENERGIA_SSCC_ACCDB = "CANTIDAD PONDERADA MW"

# El Access no trae ni el FD ni el FMA: en la planilla 3 venian pegados
# en DB!Y y DB!V. Los dos se calculan ahora desde la carpeta "FD y FMA"
# -el FMA con calcular_fma_subastas() y el FD con calcular_fd_subastas(),
# mas abajo-, asi que ya no queda ninguna columna de Subastas sin
# origen.


def construir_mapa_propietario(resumen_bess):
    """
    central normalizada -> Propietario, desde la hoja "Resumen BESS"
    de Centrales.xlsx.

    El usuario agrego la columna "Propietario" en la B de esa hoja y
    corrio todo lo demas una columna a la derecha. Como esta hoja
    siempre se leyo por NOMBRE de columna (ver _leer_resumen_bess y
    _mapa_resumen_bess_por_nombre), ese corrimiento no rompe nada.

    Si la columna no esta (un Centrales.xlsx viejo), devuelve un dict
    vacio en vez de reventar: la columna Propietario queda vacia y se
    avisa en el log.
    """

    columna_nombre = columna_propietario = None

    for columna in resumen_bess.columns:
        texto = normalizar(columna)
        if "nombre" in texto and "activ" in texto:
            columna_nombre = columna
        elif "propietario" in texto:
            columna_propietario = columna

    if columna_nombre is None or columna_propietario is None:
        return {}

    mapa = {}

    for _, fila in resumen_bess.iterrows():
        nombre = normalizar(fila[columna_nombre])
        if not nombre:
            continue
        mapa[nombre] = _texto_seguro(fila[columna_propietario])

    return mapa


def _control_desde_servicio(servicio):
    """'CSF(-)' -> 'CSF'. Replica =LEFT(C,3) de la planilla 3."""

    return _texto_seguro(servicio)[:3]


def _sub_baj_desde_servicio(servicio):
    """
    'CSF(+)' -> 'SUBIDA', 'CSF(-)' -> 'BAJADA'.

    Replica al pie de la letra =IF(LEFT(RIGHT($C9,2),1)="+","SUBIDA",
    "BAJADA"): mira el penultimo caracter, no "si contiene un +".
    """

    texto = _texto_seguro(servicio)

    return "SUBIDA" if texto[-2:-1] == "+" else "BAJADA"


def calcular_hora_mes_subastas(dias, horas, dia_cambio_hora=None, ajuste=1):
    """
    Hora_mes = (Dia - 1) * 24 + Hora_dia, mas el ajuste por cambio de
    hora que la planilla 3 aplica con =($F9-1)*24+$G9+IF(F9>$F$2,1,0).

    dia_cambio_hora es el dia del mes en que cambia la hora (la celda
    SUBASTAS!F2 de la planilla). Por omision es None = sin ajuste, que
    es lo correcto en los 10 meses del año en que no hay cambio de
    hora. Queda como parametro y no hardcodeado porque de donde sale
    ese dia cada mes es justamente lo que falta confirmar (ver
    BITACORA.md, "Pendientes abiertos").
    """

    dias = pd.to_numeric(dias, errors="coerce")
    horas = pd.to_numeric(horas, errors="coerce")

    hora_mes = (dias - 1) * 24 + horas

    if dia_cambio_hora is not None:
        hora_mes = hora_mes + (dias > int(dia_cambio_hora)).astype(int) * ajuste

    return hora_mes.astype("Int64")


def construir_subastas_desde_accdb(
    carpeta_subastas,
    aamm,
    mapa_propietarios=None,
    dia_cambio_hora=None,
    tablas_fma=None,
    dic_cpf=None,
    tablas_fd=None,
    dic_unidad_fd=None,
    registrar=print,
):
    """
    Arma la hoja "Subastas" del consolidado (las mismas columnas B:Q
    que construir_subastas) leyendo los Access de
    <CARPETA_BASE>/Subastas/DB subastas/ en vez de la planilla 3.

    El filtro es el mismo de siempre: se quedan solo las filas cuya
    Configuración contiene "BESS" o "SAE".
    """

    try:
        df_crudo, resumen = ofertas_adj.construir_crudo(
            carpeta_subastas, aamm, registrar=registrar
        )
    except ofertas_adj.ErrorSubastas as error:
        raise ErrorEntrada(str(error)) from error

    if resumen["dias_sin_archivo"]:
        registrar(
            f"  [AVISO] sin Access de subastas para el/los dia(s): "
            f"{', '.join(str(d) for d in resumen['dias_sin_archivo'])}"
        )

    # Diagnostico de cambio de hora: si algun dia no trae 24 horas, el
    # Hora_mes corrido deja de ser (Dia-1)*24+Hora para los dias
    # siguientes. No se corrige solo (ver calcular_hora_mes_subastas),
    # pero el usuario tiene que enterarse.
    horas_por_dia = df_crudo.groupby("DIA")["HORA"].max()
    dias_raros = horas_por_dia[horas_por_dia != 24]

    for dia, maximo in dias_raros.items():
        registrar(
            f"  [AVISO] el dia {int(dia)} trae {int(maximo)} horas, no 24 "
            f"(cambio de hora?): revisar Hora_mes."
        )

    # Mismo filtro BESS/SAE de siempre, ahora por nombre de columna.
    textos = df_crudo["CONFIGURACIÓN"].map(_texto_seguro)
    filtrado = df_crudo[textos.map(_contiene_bess_o_sae_sin_bat)]
    filtrado = filtrado.reset_index(drop=True)

    df = pd.DataFrame(index=filtrado.index)

    df["B"] = filtrado["SERVICIO"].map(_texto_seguro)
    df["C"] = filtrado["SERVICIO"].map(_control_desde_servicio)
    df["D"] = filtrado["SERVICIO"].map(_sub_baj_desde_servicio)

    df["F"] = pd.to_numeric(filtrado["AÑO"], errors="coerce").astype("Int64")
    df["G"] = pd.to_numeric(filtrado["MES"], errors="coerce").astype("Int64")
    df["H"] = pd.to_numeric(filtrado["DIA"], errors="coerce").astype("Int64")
    df["I"] = pd.to_numeric(filtrado["HORA"], errors="coerce").astype("Int64")

    df["E"] = pd.to_datetime(
        dict(year=df["F"], month=df["G"], day=df["H"]), errors="coerce"
    )

    df["J"] = calcular_hora_mes_subastas(
        df["H"], df["I"], dia_cambio_hora=dia_cambio_hora
    )

    df["K"] = filtrado["CONFIGURACIÓN"].map(_texto_seguro)

    mapa_propietarios = mapa_propietarios or {}
    df["L"] = df["K"].map(
        lambda central: mapa_propietarios.get(normalizar(central), "")
    )

    df["M"] = (
        df["K"]
        + df["H"].map(_entero_a_texto)
        + df["I"].map(_entero_a_texto)
    )

    df["N"] = pd.NA

    df["O"] = pd.to_numeric(
        filtrado[COLUMNA_ENERGIA_SSCC_ACCDB], errors="coerce"
    )

    df["P"] = pd.NA
    df["Q"] = pd.NA

    df = df[list("BCDEFGHIJKLMNOPQ")]
    df = df.rename(columns=NOMBRES_SUBASTAS)
    df = _ordenar_subastas_por_hora_mes(df)

    # FD (P) y FMA (Q): se calculan aca, con la hoja ya en su orden
    # definitivo, para que queden alineadas fila a fila con lo que se
    # escribe.
    if tablas_fd:
        df[NOMBRES_SUBASTAS["P"]] = calcular_fd_subastas(
            df, tablas_fd, dic_unidad_fd=dic_unidad_fd, registrar=registrar
        )
    else:
        registrar(
            f"  [AVISO] sin archivo SSCC_Desempeño_* en "
            f"{CARPETA_FD_FMA}/: la columna FD queda vacia."
        )

    if tablas_fma:
        df[NOMBRES_SUBASTAS["Q"]] = calcular_fma_subastas(
            df, tablas_fma, dic_cpf=dic_cpf, tablas_fd=tablas_fd,
            dic_unidad_fd=dic_unidad_fd, registrar=registrar,
        )
    else:
        registrar(
            f"  [AVISO] sin salidas de FMA en {CARPETA_FD_FMA}/: la "
            f"columna FMA queda vacia."
        )

    if mapa_propietarios:
        sin_propietario = sorted(
            {
                central
                for central, propietario in zip(df.iloc[:, 9], df.iloc[:, 10])
                if not propietario
            }
        )
        for central in sin_propietario:
            registrar(
                f"  [AVISO] sin Propietario en {ARCHIVO_CENTRALES} "
                f"('{HOJA_RESUMEN_BESS}'): {central}"
            )
    else:
        registrar(
            f"  [AVISO] la hoja '{HOJA_RESUMEN_BESS}' de "
            f"{ARCHIVO_CENTRALES} no tiene columna 'Propietario': la "
            f"columna Propietario de Subastas queda vacia."
        )

    registrar(
        f"  Subastas: {len(df):,} fila(s) desde "
        f"{resumen['archivos_leidos']} Access (filtro Configuración "
        f"contiene BESS/SAE), en el orden del origen."
    )

    return df
