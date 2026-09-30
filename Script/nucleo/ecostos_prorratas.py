# -*- coding: utf-8 -*-
"""
E Costos, etapa 3: AG:AL, AM:AR (FD), AS:AV.
"""

import math
import pandas as pd

from .lectura import ROL_BALANCE_BESS, ROL_FD, mapa_diccionario
from .parametros import HOJA_DICCIONARIO
from .alertas import ALTA, MEDIA, Alerta, anotar, anotar_muchas
from .utiles import (
    _columna_clave_vba, _normaliza_valor_vba, _texto_seguro,
    _tiene_valor, normalizar,
)


# ============================================================
# CALCULO E COSTOS (etapa 3): AG:AL (prorratas), AM:AR (FD), AS, AT,
# AU, AV
#
# El usuario confirmo que la categoria "CTF" (AI, AL, AO, AR) no
# existe en nuestros datos de FD y que en el archivo real esas
# columnas salen en 0 -- coincide exactamente con el VBA original,
# que las deja hardcodeadas en 0 (no dependen de ningun diccionario).
#
# Tambien confirmo que el archivo de Subastas que uso para armar nuestra
# hoja esta corrido una columna respecto del original (el original
# tiene una columna vacia al principio que el nuestro no tiene). Eso
# explica por que el VBA original documentaba "D"/"G,H,I,K" para L:
# aplicando ese corrimiento, esas letras coinciden exactamente con
# Sub_Baj/Configuración+Mes+Dia+Hora_dia -- lo mismo que ya se venia
# usando, ahora con una explicacion clara de la discrepancia.
#
# AW, AX y AZ quedaron fuera de ESTA etapa (dependian de una tabla
# de umbrales de subida/bajada que en su momento no se sabia donde
# vivia) y se resolvieron en la etapa 4, mas abajo.
# ============================================================

def construir_prorrata_sscc(df_subastas):
    """
    Construye la tabla dinamica "Prorrata SSCC" a partir de Subastas
    (confirmado por el usuario, NO es un archivo externo):
        Filas: Configuración, Hora_mes
        Columnas: Control
        Valores: Cuenta de Sub_Baj

    "Cuenta de Sub_Baj" es el AGGFUNC de la tabla dinamica (cuenta
    filas), pero el resultado final NO es esa cuenta cruda: es una
    "prorrata" (proporcion), tal como dice el nombre -- cada celda se
    divide por la suma de su propia fila (entre TODOS los valores de
    Control que aparecen para esa Configuración+Hora_mes), asi que
    cada fila termina sumando 1. Encontrado comparando fila a fila
    contra la planilla 11 real: donde nuestra cuenta cruda daba
    (CPF=1, CSF=1) la planilla real trae (0.5, 0.5); donde daba
    (CPF=2, CSF=1) trae (0.6666..., 0.3333...) -- exactamente
    cuenta/total. Antes de esta correccion se devolvia la cuenta
    cruda tal cual, lo que hacia que Prorratas > 1 aparecieran cuando
    una central tenia mas de una fila del mismo tipo en la misma
    Configuración+Hora_mes (el aviso del usuario: "a veces sale con
    2").
    """

    tabla = (
        df_subastas
        .pivot_table(
            index=["Configuración", "Hora_mes"],
            columns="Control",
            values="Sub_Baj",
            aggfunc="count",
            fill_value=0,
        )
        .reset_index()
    )
    tabla.columns.name = None

    columnas_valor = [
        c for c in tabla.columns if c not in ("Configuración", "Hora_mes")
    ]

    total_fila = tabla[columnas_valor].sum(axis=1)

    tabla[columnas_valor] = (
        tabla[columnas_valor].div(total_fila, axis=0).fillna(0.0)
    )

    return tabla


def construir_dic_prorrata(tabla_prorrata, registrar=print):
    """
    Arma central+hora_mes -> (CPF, CSF) a partir de la tabla dinamica
    Prorrata SSCC (construir_prorrata_sscc()). Busca, entre las
    columnas que dejo el pivot (una por cada valor distinto de
    Subastas!Control), la que contenga "cpf" y la que contenga "csf"
    en el nombre -- INFERIDO (no confirmado letra por letra que
    Control tenga exactamente esos dos valores); si no las encuentra,
    avisa y usa 0 en vez de fallar.
    """

    columnas_valor = [
        c for c in tabla_prorrata.columns
        if c not in ("Configuración", "Hora_mes")
    ]

    columna_cpf = next(
        (c for c in columnas_valor if "cpf" in normalizar(str(c))), None
    )
    columna_csf = next(
        (c for c in columnas_valor if "csf" in normalizar(str(c))), None
    )

    if columna_cpf is None or columna_csf is None:
        registrar(
            f"  [AVISO] La columna 'Control' de Subastas no tiene valores "
            f"identificables como CPF/CSF (encontrados: {columnas_valor}); "
            f"CPF(-)/CSF(-)/CPF(+)/CSF(+) de Prorrata SSCC quedaran en 0."
        )

    dic = {}

    for _, fila in tabla_prorrata.iterrows():

        clave = (
            _normaliza_valor_vba(fila["Configuración"])
            + "¦" + _normaliza_valor_vba(fila["Hora_mes"])
        )

        cpf = float(fila[columna_cpf]) if columna_cpf is not None else 0.0
        csf = float(fila[columna_csf]) if columna_csf is not None else 0.0

        dic[clave] = (cpf, csf)

    return dic


def calcular_prorratas(df_ecostos, dic_prorrata, registrar=print):
    """
    Replica AG/AH (y sus duplicados AJ/AK, ver comentario de
    completar_calculo_e_costos_grupos): busca central+"Hora Mes" en
    el diccionario de construir_dic_prorrata(); si no hay match,
    ambas quedan en 0 (igual que el original, que no distingue
    "sin match" de "cero").
    """

    clave = (
        _columna_clave_vba(df_ecostos["clave"])
        + "¦" + _columna_clave_vba(df_ecostos["Hora Mes"])
    )

    valores = clave.map(lambda k: dic_prorrata.get(k, (0.0, 0.0)))

    sin_match = ~clave.isin(set(dic_prorrata))

    # Una hora suelta sin prorrata puede ser legitima (esa hora no tuvo
    # SSCC); una central que NO aparece en NINGUNA hora es otra cosa:
    # la homologacion contra Subastas!Configuración no esta cruzando y
    # AG/AH -- y con ellas todo el prorrateo de costos SSCC -- quedan
    # en 0 para esa central sin que nada lo diga.
    faltantes_por_central = sorted({
        _texto_seguro(central)
        for central, falta in zip(df_ecostos["clave"], sin_match)
        if falta and _tiene_valor(central)
    })

    con_match_por_central = {
        _texto_seguro(central)
        for central, falta in zip(df_ecostos["clave"], sin_match)
        if not falta
    }

    sin_ninguna = [
        central for central in faltantes_por_central
        if central not in con_match_por_central
    ]

    if sin_ninguna:
        anotar_muchas(
            registrar,
            [
                Alerta(
                    "PRO-001", ALTA, "Calculo E Costos",
                    "Central sin ninguna hora en la Prorrata SSCC.",
                    central=central,
                    valor_esperado="al menos una hora en la Prorrata SSCC",
                    accion="AG/AH quedan en 0 y con ellas todo el "
                           "prorrateo de SSCC de esa central",
                    hoja="Subastas",
                    origen_control="CONTROL NUEVO",
                )
                for central in sin_ninguna
            ],
            f"  [{ALTA}] PRO-001: Calculo E Costos: {len(sin_ninguna):,} "
            f"central(es) no tienen ninguna hora en la Prorrata SSCC "
            f"({', '.join(repr(v) for v in sin_ninguna[:15])}); AG/AH "
            f"quedan en 0 y con ellas todo el prorrateo de SSCC.",
        )
    elif faltantes_por_central:
        anotar(registrar, Alerta(
            "PRO-002", MEDIA, "Calculo E Costos",
            f"{int(sin_match.sum()):,} fila(s) sin prorrata SSCC para su "
            f"central+'Hora Mes'.",
            valor_encontrado=f"{int(sin_match.sum()):,} filas",
            accion="AG/AH quedan en 0 en esas filas",
            hoja="Subastas",
            origen_control="CONTROL NUEVO",
        ))

    ag = valores.map(lambda v: v[0])
    ah = valores.map(lambda v: v[1])

    return ag, ah


def construir_dic_mapeo_diccionario(diccionario):
    """
    Replica CrearDiccionarioPrimerValor(datosDiccionario, 1, 2):
    nombre canonico de la central -> nombre de esa central en el
    archivo de desempeño (el FD), la primera coincidencia gana. Es un
    mapeo DISTINTO de construir_homologacion() (usa toda la fila) y de
    _mapas_homologacion_fge() (va de los otros origenes hacia el
    canonico, no al reves) -- tres lecturas distintas de la misma hoja
    Diccionario, no fusionar.

    FORMATO NUEVO de la hoja (tabla unica con encabezados, ver
    lectura.py): columna "Balance_BESS" -> columna "FD". FORMATO VIEJO
    (bloques lado a lado): columna A -> columna B, que es justamente
    donde vive el bloque "FD".
    """

    dic = mapa_diccionario(diccionario, ROL_BALANCE_BESS, ROL_FD)

    if dic:
        return dic

    for _, fila in diccionario.iterrows():

        clave = fila[0]
        if pd.isna(clave):
            continue

        clave_norm = normalizar(clave)

        if clave_norm and clave_norm not in dic:
            dic[clave_norm] = fila[1]

    return dic


def _calcular_bloque(valor):
    """Replica CalcularBloque: Int((valor-1)/4)+1."""

    return math.floor((valor - 1.0) / 4.0) + 1.0


def construir_dic_fd_bloque(df_fd, columna_id, columna_mas, columna_menos):
    """
    Arma id->(valor_mas, valor_menos) a partir de un bloque de FD
    (CSF o CPF, ya con nombres reales de columna), clave = columna_id
    normalizada, primera coincidencia gana. Replica
    CrearDiccionarioFDAL/CrearDiccionarioFDQAD.
    """

    dic = {}

    for _, fila in df_fd.iterrows():

        clave = normalizar(fila[columna_id])

        if clave and clave not in dic:
            dic[clave] = (fila[columna_mas], fila[columna_menos])

    return dic


def unidades_bloque_fd(df_fd, columna_unidad="Unidad"):
    """
    Las unidades que trae un bloque de la hoja FD, como
    {nombre normalizado: nombre tal cual}. Se usa para distinguir "a
    esta unidad le falta una hora" de "a esta unidad no la tiene nadie
    en este bloque", y para poder listarle al usuario, con su
    ortografia real, los nombres entre los que tiene que elegir.
    """

    if df_fd is None or columna_unidad not in getattr(df_fd, "columns", []):
        return {}

    unidades = {}

    for valor in df_fd[columna_unidad]:
        if _tiene_valor(valor):
            unidades.setdefault(normalizar(valor), _texto_seguro(valor))

    return unidades


def _nombres_unidades(unidades):
    """Los nombres tal cual, venga un dict de unidades_bloque_fd o un set."""

    if isinstance(unidades, dict):
        return list(unidades.values())

    return list(unidades)


def calcular_fd_prorrateado(
    df_ecostos, dic_mapeo, dic_fd_csf, dic_fd_cpf, registrar=print,
    unidades_fd_csf=None, unidades_fd_cpf=None,
):
    """
    Replica AM, AN, AP, AQ: homologa la central ("clave") contra
    Diccionario!A->B; si no esta en el Diccionario, las 4 quedan en
    blanco (pd.NA, equivalente al #N/A del original). Si esta,
    arma una clave "bloque+central homologada" (bloque = CalcularBloque
    de Y para descarga, de AC para carga) y busca esa clave en los
    diccionarios de FD (CPF da AM/AP, CSF da AN/AQ); si no hay match
    en FD, queda en 0 (fiel al original: ahi solo se registra un
    aviso, no se propaga un error).

    unidades_fd_csf / unidades_fd_cpf: los nombres de unidad que trae
    cada bloque de la hoja FD (`unidades_bloque_fd()`). Sirven para
    separar dos cosas MUY distintas que antes se informaban igual:

      - a una unidad que SI esta en el bloque le falta alguna hora
        suelta -> FD-005, una alerta por clave, como pide el catalogo;
      - una unidad no aparece NUNCA en el bloque -> FD-007, UNA alerta
        por central, con la lista de unidades que si estan.

    El segundo caso es una homologacion que no cruza, no un dato
    faltante, y repetirlo una vez por hora es lo que hacia inservible
    la hoja Alertas: en la primera corrida real del usuario, 2.232
    alertas FD-005 eran en realidad 3 hechos (una central sin ninguna
    fila en FD, y otra sin ninguna fila en el bloque CSF). Por eso, en
    ese caso, NO se emiten ademas las 744 FD-005 de esa unidad: son la
    misma frase repetida.
    """

    y_num = pd.to_numeric(df_ecostos["Y"], errors="coerce").fillna(0.0)
    ac_num = pd.to_numeric(df_ecostos["AC"], errors="coerce").fillna(0.0)

    bloque_y = y_num.map(_calcular_bloque)
    bloque_ac = ac_num.map(_calcular_bloque)

    unidades_fd_cpf = unidades_fd_cpf or {}
    unidades_fd_csf = unidades_fd_csf or {}

    am, an, ap, aq = [], [], [], []
    sin_diccionario = set()
    sin_fd_cpf = set()
    sin_fd_csf = set()

    # central -> unidad homologada, para poder nombrar a las dos en el
    # aviso (el usuario piensa en la central, la hoja FD en la unidad).
    unidad_de_central = {}
    ausentes_cpf = {}
    ausentes_csf = {}

    for central, by, bac in zip(df_ecostos["clave"], bloque_y, bloque_ac):

        mapeo = dic_mapeo.get(normalizar(central))

        if mapeo is None:
            sin_diccionario.add(_texto_seguro(central))
            am.append(pd.NA)
            an.append(pd.NA)
            ap.append(pd.NA)
            aq.append(pd.NA)
            continue

        mapeo_texto = _normaliza_valor_vba(mapeo)
        unidad = normalizar(mapeo_texto)

        texto_central = _texto_seguro(central)
        unidad_de_central.setdefault(texto_central, _texto_seguro(mapeo))

        clave_y = normalizar(f"{int(by)}{mapeo_texto}")
        clave_ac = normalizar(f"{int(bac)}{mapeo_texto}")

        cpf_y = dic_fd_cpf.get(clave_y)
        csf_y = dic_fd_csf.get(clave_y)
        csf_ac = dic_fd_csf.get(clave_ac)

        # Si la unidad no esta en el bloque, el faltante no es de esta
        # hora: es de la central entera (FD-007 mas abajo).
        ausente_cpf = unidades_fd_cpf and unidad not in unidades_fd_cpf
        ausente_csf = unidades_fd_csf and unidad not in unidades_fd_csf

        if ausente_cpf:
            ausentes_cpf[texto_central] = ausentes_cpf.get(texto_central, 0) + 1
        elif cpf_y is None:
            sin_fd_cpf.add(clave_y)

        if ausente_csf:
            ausentes_csf[texto_central] = ausentes_csf.get(texto_central, 0) + 1
        else:
            if csf_y is None:
                sin_fd_csf.add(clave_y)
            if csf_ac is None:
                sin_fd_csf.add(clave_ac)

        ap.append(cpf_y[0] if cpf_y is not None else 0.0)
        am.append(cpf_y[1] if cpf_y is not None else 0.0)
        an.append(csf_y[1] if csf_y is not None else 0.0)
        aq.append(csf_ac[0] if csf_ac is not None else 0.0)

    if sin_diccionario:
        anotar_muchas(
            registrar,
            [
                Alerta(
                    "DIC-001", ALTA, "Calculo E Costos",
                    f"Central no presente en el diccionario de "
                    f"nomenclatura ({HOJA_DICCIONARIO} A:B).",
                    central=central,
                    valor_esperado=f"una fila en {HOJA_DICCIONARIO} A:B",
                    accion="AM, AN, AP y AQ quedan vacias para esa central",
                    hoja=HOJA_DICCIONARIO,
                    origen_control="CATALOGO AUX-004",
                )
                for central in sorted(sin_diccionario)
            ],
            f"  [{ALTA}] DIC-001: FD de Calculo E Costos: central(es) no "
            f"presentes en el diccionario de nomenclatura "
            f"({HOJA_DICCIONARIO} A:B): "
            f"{', '.join(repr(v) for v in sorted(sin_diccionario))}. "
            f"AM, AN, AP y AQ quedan vacias para esas centrales.",
        )

    # FD-007 (control nuevo): la unidad homologada no aparece en NINGUNA
    # fila del bloque. Una alerta por central, no una por hora.
    for etiqueta, ausentes, disponibles in (
        ("CPF", ausentes_cpf, unidades_fd_cpf),
        ("CSF", ausentes_csf, unidades_fd_csf),
    ):
        if not ausentes:
            continue
        anotar_muchas(
            registrar,
            [
                Alerta(
                    "FD-007", ALTA, "Calculo E Costos",
                    f"La unidad homologada no aparece en ninguna fila "
                    f"del bloque {etiqueta} de la hoja FD.",
                    central=central,
                    clave=unidad_de_central.get(central, ""),
                    valor_encontrado=f"{filas:,} fila(s) sin {etiqueta}",
                    valor_esperado=(
                        f"la unidad en el bloque {etiqueta} de la hoja FD"
                    ),
                    accion=f"el FD {etiqueta} de esa central queda en 0",
                    hoja="FD",
                    origen_control="CONTROL NUEVO",
                )
                for central, filas in sorted(ausentes.items())
            ],
            f"  [{ALTA}] FD-007: FD de Calculo E Costos: "
            f"{len(ausentes)} central(es) no aparecen en NINGUNA fila del "
            f"bloque {etiqueta} de la hoja FD, asi que su FD {etiqueta} "
            f"queda en 0: "
            + "; ".join(
                f"{central} -> {unidad_de_central.get(central, '')!r}"
                for central in sorted(ausentes)
            )
            + f". Unidades que SI trae ese bloque: "
            f"{', '.join(sorted(_nombres_unidades(disponibles)))}. "
            f"Revisar la homologacion "
            f"en la hoja {HOJA_DICCIONARIO}, o confirmar que esa central "
            f"no presta {etiqueta}.",
        )

    # FD-005 del catalogo: se guardan TODAS las claves faltantes, no
    # una muestra. A la pantalla va el resumen con 10 ejemplos.
    for etiqueta, faltantes in (("CPF", sin_fd_cpf), ("CSF", sin_fd_csf)):
        if not faltantes:
            continue
        anotar_muchas(
            registrar,
            [
                Alerta(
                    "FD-005", ALTA, "Calculo E Costos",
                    f"Clave {etiqueta} homologada que no aparece en la "
                    f"hoja FD.",
                    clave=clave,
                    valor_esperado="una fila en la hoja FD",
                    accion="el valor se completa con 0",
                    hoja="FD",
                    origen_control="CATALOGO FD-005",
                )
                for clave in sorted(faltantes)
            ],
            f"  [{ALTA}] FD-005: FD de Calculo E Costos: "
            f"{len(faltantes):,} clave(s) {etiqueta} homologadas no "
            f"aparecen en la hoja FD; sus valores se completan con 0. "
            f"Ejemplos: {', '.join(sorted(faltantes)[:10])}"
            f"{' (detalle completo en la hoja Alertas)' if len(faltantes) > 10 else ''}.",
        )

    return (
        pd.Series(am, index=df_ecostos.index),
        pd.Series(an, index=df_ecostos.index),
        pd.Series(ap, index=df_ecostos.index),
        pd.Series(aq, index=df_ecostos.index),
    )


def _calcular_costo_ponderado(cantidad1, cantidad2, precio1, precio2, energia):
    """Replica CalcularCostoPonderado (AS/AT)."""

    if pd.isna(energia):
        return pd.NA

    if (cantidad1 + cantidad2) > 0:

        if pd.isna(precio1) or pd.isna(precio2):
            return pd.NA

        factor = float(precio1) * cantidad1 + float(precio2) * cantidad2

    else:
        factor = 1.0

    return factor * float(energia)


def calcular_as_at(df_ecostos):
    """
    Replica AS (energia descarga con FD) y AT (energia carga con
    FD): CalcularCostoPonderado combinando las prorratas (AG/AH), el
    FD homologado (AM/AN para AS, AP/AQ para AT) y la energia
    asignada (AE para AS, AF para AT).
    """

    as_ = [
        _calcular_costo_ponderado(ag, ah, am, an, ae)
        for ag, ah, am, an, ae in zip(
            df_ecostos["AG"], df_ecostos["AH"],
            df_ecostos["AM"], df_ecostos["AN"], df_ecostos["AE"],
        )
    ]

    at = [
        _calcular_costo_ponderado(ag, ah, ap, aq, af)
        for ag, ah, ap, aq, af in zip(
            df_ecostos["AG"], df_ecostos["AH"],
            df_ecostos["AP"], df_ecostos["AQ"], df_ecostos["AF"],
        )
    ]

    return (
        pd.Series(as_, index=df_ecostos.index),
        pd.Series(at, index=df_ecostos.index),
    )


def calcular_au_av(df_ecostos):
    """
    Replica AU (ingreso descarga) y AV (costo carga): promedio de AB
    (donde AE es valido y distinto de 0, dentro del grupo central+
    ventana) multiplicado por AE, solo si la suma GLOBAL de energia
    de descarga por ventana (todas las centrales que comparten esa
    misma "Copia_Ventana"/P) supera 10; analogo para AV con AF/AD,
    umbral de suma global menor a -10.
    """

    ae = df_ecostos["AE"]
    af = df_ecostos["AF"]
    ab = df_ecostos["AB"]
    ad = df_ecostos["AD"]

    grupo = [df_ecostos["clave"], df_ecostos["Copia_Ventana"]]

    ae_valido = ae.notna() & (pd.to_numeric(ae, errors="coerce") != 0)
    af_valido = af.notna() & (pd.to_numeric(af, errors="coerce") != 0)

    ab_contable = ab.where(ae_valido & ab.notna())
    ad_contable = ad.where(af_valido & ad.notna())

    suma_ab = ab_contable.groupby(grupo).transform("sum")
    cantidad_ab = ab_contable.groupby(grupo).transform("count")
    hay_error_ae = ae.isna().groupby(grupo).transform("any")

    suma_ad = ad_contable.groupby(grupo).transform("sum")
    cantidad_ad = ad_contable.groupby(grupo).transform("count")
    hay_error_af = af.isna().groupby(grupo).transform("any")

    promedio_ab = (suma_ab / cantidad_ab.replace(0, pd.NA)).fillna(0.0)
    promedio_ad = (suma_ad / cantidad_ad.replace(0, pd.NA)).fillna(0.0)

    # sumaIP/sumaJP: suma GLOBAL por ventana (P), agrupando TODAS las
    # centrales que comparten esa Copia_Ventana -- no por central+P.
    suma_i_global = (
        df_ecostos["Energia_Positiva"]
        .groupby(df_ecostos["Copia_Ventana"])
        .transform("sum")
    )
    suma_j_global = (
        df_ecostos["Energia_Negativa"]
        .groupby(df_ecostos["Copia_Ventana"])
        .transform("sum")
    )

    condicion_au = (
        (cantidad_ab > 0) & (~hay_error_ae) & (suma_i_global > 10) & ae.notna()
    )
    condicion_av = (
        (cantidad_ad > 0) & (~hay_error_af) & (suma_j_global < -10) & af.notna()
    )

    au = (
        promedio_ab * pd.to_numeric(ae, errors="coerce").fillna(0.0)
    ).where(condicion_au, 0.0)
    av = (
        promedio_ad * pd.to_numeric(af, errors="coerce").fillna(0.0)
    ).where(condicion_av, 0.0)

    return au, av
