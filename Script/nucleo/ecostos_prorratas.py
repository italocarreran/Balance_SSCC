# -*- coding: utf-8 -*-
"""
E Costos, etapa 3: prorrata (AG:AL), FD (AM:AR) y energia ajustada (AS, AT).
"""

import pandas as pd

from .lectura import (
    ROL_BALANCE_BESS, ROL_FD, filas_diccionario, mapa_diccionario,
)
from .parametros import COLUMNA_CENTRAL, HOJA_DICCIONARIO
from .alertas import ALTA, MEDIA, Alerta, anotar, anotar_muchas
from .utiles import (
    _columna_clave_vba, _normaliza_valor_vba, _texto_seguro,
    _tiene_valor, normalizar,
)


# ============================================================
# CALCULO E COSTOS (etapa 3): prorrata (AG:AL), FD (AM:AR) y la
# energia ajustada (AS, AT)
#
# Metodo nuevo (macro SSCC_Calcular_Integrado, desde septiembre 2026):
# el FD se aplica a la ENERGIA MEDIDA de cada cuarto de hora, en su
# propia hora (Hora Mes), antes de acumular y de repartir en la curva:
#
#   prorrata de la hora   CPF(-), CSF(-), CPF(+), CSF(+) por separado
#   FD de la hora         el de cada servicio, solo si su prorrata <> 0
#   factor                = SUMA(prorrata x FD) de los cuatro servicios
#   AS (descarga ajustada) = I x factor
#   AT (carga ajustada)    = J x factor
#
# Solo las filas con Adj SSCC (L) = 1: el resto queda en 0. Antes (ver
# la rama Definitivo-Agosto) la prorrata no tenia direccion, el FD se
# buscaba en la hora reordenada de la curva (Y / AC) y se aplicaba a la
# energia ya repartida (AE / AF).
#
# La categoria "CTF" (AI, AL, AO, AR) no existe en nuestros datos de FD
# y queda en 0, igual que en la macro.
# ============================================================

# Los cuatro servicios con prorrata y FD, en el orden de las columnas
# AG, AH, AJ, AK (prorrata) y AM, AN, AP, AQ (FD).
SERVICIOS_PRORRATA = ("CPF(-)", "CSF(-)", "CPF(+)", "CSF(+)")

# La suma de las prorratas de una hora con SSCC tiene que dar 100%.
TOLERANCIA_PRORRATA = 1e-6


def construir_prorrata_sscc(df_subastas):
    """
    Construye la tabla dinamica "Prorrata SSCC" a partir de Subastas:
        Filas: Configuración, Hora_mes
        Columnas: Concepto (CPF(-), CPF(+), CSF(-), CSF(+), ...)
        Valores: cuenta de filas, dividida por el total de la fila

    Cada celda es la participacion del servicio en esa central+hora:
    la fila suma 1 (100%). Con el metodo nuevo la prorrata lleva
    direccion (-)/(+) por separado; antes se armaba por "Control"
    (CPF/CSF sin direccion) y el mismo valor servia para las dos.
    """

    tabla = (
        df_subastas
        .assign(_concepto=df_subastas["Concepto"].map(_normaliza_valor_vba))
        .pivot_table(
            index=["Configuración", "Hora_mes"],
            columns="_concepto",
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
    Arma central¦hora_mes -> (CPF(-), CSF(-), CPF(+), CSF(+)) a partir
    de la tabla dinamica Prorrata SSCC (construir_prorrata_sscc()). Un
    servicio que no aparece en ninguna fila de Subastas vale 0; si no
    aparece ninguno de los cuatro, se avisa.
    """

    columnas = {
        servicio: servicio if servicio in tabla_prorrata.columns else None
        for servicio in SERVICIOS_PRORRATA
    }

    if not any(columnas.values()):
        otras = [
            c for c in tabla_prorrata.columns
            if c not in ("Configuración", "Hora_mes")
        ]
        registrar(
            f"  [AVISO] La columna 'Concepto' de Subastas no trae ninguno "
            f"de {', '.join(SERVICIOS_PRORRATA)} (encontrados: {otras}); "
            f"la prorrata de Calculo E Costos queda en 0."
        )

    dic = {}

    for _, fila in tabla_prorrata.iterrows():

        clave = (
            _normaliza_valor_vba(fila["Configuración"])
            + "¦" + _normaliza_valor_vba(fila["Hora_mes"])
        )

        dic[clave] = tuple(
            float(fila[columna]) if columna is not None else 0.0
            for columna in columnas.values()
        )

    return dic


def _con_sscc(df_ecostos):
    """Las filas con Adj SSCC (L) = 1; sin columna L, todas."""

    if "L" not in df_ecostos:
        return pd.Series(True, index=df_ecostos.index)

    return pd.to_numeric(df_ecostos["L"], errors="coerce").eq(1)


def calcular_prorratas(df_ecostos, dic_prorrata, registrar=print):
    """
    Replica AG (CPF(-)), AH (CSF(-)), AJ (CPF(+)) y AK (CSF(+)): busca
    central + "Hora Mes" (la hora de la fila) en el diccionario de
    construir_dic_prorrata(). Solo en las filas con L = 1; el resto, y
    las que no tienen prorrata, quedan en 0.

    Devuelve las cuatro Series en ese orden.
    """

    con_sscc = _con_sscc(df_ecostos)

    clave = (
        _columna_clave_vba(df_ecostos["clave"])
        + "¦" + _columna_clave_vba(df_ecostos["Hora Mes"])
    )

    sin_prorrata = (0.0,) * len(SERVICIOS_PRORRATA)
    valores = [
        dic_prorrata.get(k, sin_prorrata) if sscc else sin_prorrata
        for k, sscc in zip(clave, con_sscc)
    ]

    sin_match = con_sscc & ~clave.isin(set(dic_prorrata))

    # Una hora suelta sin prorrata puede ser legitima (esa hora no tuvo
    # SSCC); una central que NO aparece en NINGUNA hora es otra cosa:
    # la homologacion contra Subastas!Configuración no esta cruzando y
    # la energia de esa central queda en 0 sin que nada lo diga.
    faltantes_por_central = sorted({
        _texto_seguro(central)
        for central, falta in zip(df_ecostos["clave"], sin_match)
        if falta and _tiene_valor(central)
    })

    con_match_por_central = {
        _texto_seguro(central)
        for central, falta, sscc in zip(df_ecostos["clave"], sin_match, con_sscc)
        if sscc and not falta
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
                    accion="la prorrata queda en 0 y con ella la energia "
                           "ajustada de esa central",
                    hoja="Subastas",
                    origen_control="CONTROL NUEVO",
                )
                for central in sin_ninguna
            ],
            f"  [{ALTA}] PRO-001: Calculo E Costos: {len(sin_ninguna):,} "
            f"central(es) no tienen ninguna hora en la Prorrata SSCC "
            f"({', '.join(repr(v) for v in sin_ninguna[:15])}); su "
            f"prorrata y su energia ajustada quedan en 0.",
        )
    elif faltantes_por_central:
        anotar(registrar, Alerta(
            "PRO-002", MEDIA, "Calculo E Costos",
            f"{int(sin_match.sum()):,} fila(s) con Adj SSCC = 1 sin "
            f"prorrata SSCC para su central+'Hora Mes'.",
            valor_encontrado=f"{int(sin_match.sum()):,} filas",
            accion="la prorrata queda en 0 en esas filas",
            hoja="Subastas",
            origen_control="CONTROL NUEVO",
        ))

    # Control de la macro nueva: con Adj SSCC = 1, las prorratas de la
    # hora suman 100%. Solo puede fallar si Subastas trae otro Concepto
    # en esa hora (CTF, o un "REVISAR"): esa parte no se paga.
    suma = pd.Series([sum(v) for v in valores], index=df_ecostos.index)
    mal_sumadas = con_sscc & ~sin_match & (suma - 1.0).abs().gt(
        TOLERANCIA_PRORRATA
    )

    if mal_sumadas.any():
        ejemplos = sorted({
            f"{_texto_seguro(c)} hora {_texto_seguro(h)} ({s:.2%})"
            for c, h, s in zip(
                df_ecostos.loc[mal_sumadas, "clave"],
                df_ecostos.loc[mal_sumadas, "Hora Mes"],
                suma[mal_sumadas],
            )
        })
        anotar(registrar, Alerta(
            "PRO-003", MEDIA, "Calculo E Costos",
            f"{int(mal_sumadas.sum()):,} fila(s) con Adj SSCC = 1 cuyas "
            f"prorratas CPF/CSF no suman 100% (ej.: "
            f"{', '.join(ejemplos[:5])}).",
            valor_encontrado=f"{int(mal_sumadas.sum()):,} filas",
            valor_esperado="100%",
            accion="se usa la prorrata tal cual",
            hoja="Subastas",
            origen_control="CONTROL NUEVO",
        ))

    return tuple(
        pd.Series([v[posicion] for v in valores], index=df_ecostos.index)
        for posicion in range(len(SERVICIOS_PRORRATA))
    )


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


def mapa_unidad_fd_a_central(diccionario):
    """
    El mismo par de columnas que construir_dic_mapeo_diccionario(),
    leido al reves: unidad del FD normalizada -> nombre Balance_BESS
    (tal cual). Es lo que pone la columna 'Central' de la hoja FD al
    importarla. La primera central que reclama una unidad gana.
    """

    pares = [
        (fila.get(ROL_BALANCE_BESS), fila.get(ROL_FD))
        for fila in filas_diccionario(diccionario)
    ]

    if not any(_tiene_valor(unidad) for _, unidad in pares):
        # Formato viejo: A -> B (ver construir_dic_mapeo_diccionario).
        pares = [
            (fila[0], fila[1]) for _, fila in diccionario.iterrows()
        ] if diccionario.shape[1] > 1 else []

    mapa = {}

    for central, unidad in pares:
        if _tiene_valor(central) and _tiene_valor(unidad):
            mapa.setdefault(normalizar(unidad), _texto_seguro(central))

    return mapa


def agregar_central_fd(df_fd, mapa_central, columna_unidad="Unidad"):
    """
    Copia del bloque FD con la columna COLUMNA_CENTRAL al final: el
    nombre Balance_BESS de la unidad de cada fila ("" si el
    Diccionario no la reclama).
    """

    df = df_fd.copy()
    df[COLUMNA_CENTRAL] = df[columna_unidad].map(
        lambda unidad: mapa_central.get(normalizar(unidad), "")
    )

    return df


def dic_mapeo_desde_fd(*bloques_fd):
    """
    central normalizada -> unidad, desde la columna COLUMNA_CENTRAL que
    ya trae la hoja FD (ver agregar_central_fd). Es el mismo mapeo que
    construir_dic_mapeo_diccionario() pero sin volver a homologar: la
    central ya viene escrita al lado de cada fila. {} si la hoja es de
    antes de que existiera la columna.
    """

    dic = {}

    for df_fd in bloques_fd:

        if df_fd is None or COLUMNA_CENTRAL not in df_fd.columns:
            continue

        for central, unidad in zip(df_fd[COLUMNA_CENTRAL], df_fd["Unidad"]):
            clave = normalizar(central)
            if clave and _tiene_valor(unidad):
                dic.setdefault(clave, unidad)

    return dic


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


def _fd_numero(valor):
    """El FD como numero; None si viene vacio o no es numerico."""

    numero = pd.to_numeric(pd.Series([valor]), errors="coerce").iloc[0]
    return None if pd.isna(numero) else float(numero)


def calcular_fd_prorrateado(
    df_ecostos, dic_mapeo, dic_fd_csf, dic_fd_cpf, registrar=print,
    unidades_fd_csf=None, unidades_fd_cpf=None,
):
    """
    Replica AM (FD CPF(-)), AN (FD CSF(-)), AP (FD CPF(+)) y AQ
    (FD CSF(+)) del metodo nuevo: homologa la central ("clave") con
    el Diccionario (o la columna Central de la hoja FD) y busca en la
    hoja FD la clave "Hora Mes de la fila + unidad homologada" -- la
    hora de la propia fila, no la hora reordenada de la curva (Y / AC)
    del metodo anterior. CPF sale del bloque CPF, CSF del CSF.

    Solo se busca en las filas con L = 1 y para los servicios con
    prorrata (AG, AH, AJ, AK) distinta de 0; si df_ecostos todavia no
    trae la prorrata, se buscan todos. Lo que no se busca, lo que no
    se encuentra y lo que no es numerico queda en 0 (la macro nueva no
    deja errores: avisa y sigue).

    unidades_fd_csf / unidades_fd_cpf: los nombres de unidad que trae
    cada bloque de la hoja FD (`unidades_bloque_fd()`). Separan dos
    cosas MUY distintas:

      - a una unidad que SI esta en el bloque le falta alguna hora
        suelta -> FD-005, una alerta por clave, como pide el catalogo;
      - una unidad no aparece NUNCA en el bloque -> FD-007, UNA alerta
        por central, con la lista de unidades que si estan.

    El segundo caso es una homologacion que no cruza, no un dato
    faltante, y repetirlo una vez por hora es lo que hacia inservible
    la hoja Alertas (en una corrida real, 2.232 FD-005 eran 3 hechos).

    Devuelve (AM, AN, AP, AQ).
    """

    con_sscc = _con_sscc(df_ecostos)

    def _prorrata(columna):
        if columna not in df_ecostos:
            return pd.Series(1.0, index=df_ecostos.index)
        return pd.to_numeric(df_ecostos[columna], errors="coerce").fillna(0.0)

    pr_cpf_menos, pr_csf_menos = _prorrata("AG"), _prorrata("AH")
    pr_cpf_mas, pr_csf_mas = _prorrata("AJ"), _prorrata("AK")

    unidades_fd_cpf = unidades_fd_cpf or {}
    unidades_fd_csf = unidades_fd_csf or {}

    filas = []
    sin_diccionario = set()
    sin_fd_cpf = set()
    sin_fd_csf = set()
    fd_no_numerico = set()

    # central -> unidad homologada, para poder nombrar a las dos en el
    # aviso (el usuario piensa en la central, la hoja FD en la unidad).
    unidad_de_central = {}
    ausentes_cpf = {}
    ausentes_csf = {}

    for posicion, (central, hora) in enumerate(
        zip(df_ecostos["clave"], df_ecostos["Hora Mes"])
    ):

        fila = [0.0, 0.0, 0.0, 0.0]   # AM, AN, AP, AQ
        filas.append(fila)

        pide = (
            pr_cpf_menos.iloc[posicion] != 0,
            pr_csf_menos.iloc[posicion] != 0,
            pr_cpf_mas.iloc[posicion] != 0,
            pr_csf_mas.iloc[posicion] != 0,
        )

        if not con_sscc.iloc[posicion] or not any(pide):
            continue

        mapeo = dic_mapeo.get(normalizar(central))

        if mapeo is None:
            sin_diccionario.add(_texto_seguro(central))
            continue

        hora_num = pd.to_numeric(pd.Series([hora]), errors="coerce").iloc[0]
        if pd.isna(hora_num):
            continue

        mapeo_texto = _normaliza_valor_vba(mapeo)
        unidad = normalizar(mapeo_texto)

        texto_central = _texto_seguro(central)
        unidad_de_central.setdefault(texto_central, _texto_seguro(mapeo))

        clave_fd = normalizar(f"{int(hora_num)}{mapeo_texto}")

        # (indice en la fila, prorrata pedida, bloque FD, ..., posicion
        # del valor en la tupla (mas, menos) del bloque)
        for indice, pedido, dic_fd, unidades, ausentes, sin_fd, lado in (
            (0, pide[0], dic_fd_cpf, unidades_fd_cpf, ausentes_cpf, sin_fd_cpf, 1),
            (1, pide[1], dic_fd_csf, unidades_fd_csf, ausentes_csf, sin_fd_csf, 1),
            (2, pide[2], dic_fd_cpf, unidades_fd_cpf, ausentes_cpf, sin_fd_cpf, 0),
            (3, pide[3], dic_fd_csf, unidades_fd_csf, ausentes_csf, sin_fd_csf, 0),
        ):
            if not pedido:
                continue

            # Si la unidad no esta en el bloque, el faltante no es de
            # esta hora: es de la central entera (FD-007 mas abajo).
            if unidades and unidad not in unidades:
                ausentes[texto_central] = ausentes.get(texto_central, 0) + 1
                continue

            par = dic_fd.get(clave_fd)
            if par is None:
                sin_fd.add(clave_fd)
                continue

            valor = _fd_numero(par[lado])
            if valor is None:
                fd_no_numerico.add(clave_fd)
                continue

            fila[indice] = valor

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
                    accion="el FD de esa central queda en 0",
                    hoja=HOJA_DICCIONARIO,
                    origen_control="CATALOGO AUX-004",
                )
                for central in sorted(sin_diccionario)
            ],
            f"  [{ALTA}] DIC-001: FD de Calculo E Costos: central(es) no "
            f"presentes en el diccionario de nomenclatura "
            f"({HOJA_DICCIONARIO} A:B): "
            f"{', '.join(repr(v) for v in sorted(sin_diccionario))}. "
            f"Su FD queda en 0.",
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

    if fd_no_numerico:
        anotar(registrar, Alerta(
            "FD-008", ALTA, "Calculo E Costos",
            f"{len(fd_no_numerico):,} clave(s) de la hoja FD con el FD "
            f"vacio o no numerico en un servicio con prorrata (ej.: "
            f"{', '.join(sorted(fd_no_numerico)[:10])}).",
            valor_encontrado=f"{len(fd_no_numerico):,} claves",
            valor_esperado="un FD numerico",
            accion="ese FD se toma como 0",
            hoja="FD",
            origen_control="CONTROL NUEVO",
        ))

    return tuple(
        pd.Series([fila[indice] for fila in filas], index=df_ecostos.index)
        for indice in range(4)
    )


def calcular_energia_ajustada(df_ecostos):
    """
    Replica AS ("Descarga ajustada FD-prorrata kWh") y AT ("Carga
    ajustada FD-prorrata kWh"): la energia medida de la fila (I, J)
    por el factor de su hora,

        factor = AG*AM + AH*AN + AJ*AP + AK*AQ

    (los CTF son 0). El mismo factor para la descarga y la carga; la
    carga conserva su signo negativo. Las filas con L <> 1 tienen
    prorrata 0, asi que quedan en 0.
    """

    def _num(columna):
        return pd.to_numeric(df_ecostos[columna], errors="coerce").fillna(0.0)

    factor = (
        _num("AG") * _num("AM") + _num("AH") * _num("AN")
        + _num("AJ") * _num("AP") + _num("AK") * _num("AQ")
    )

    return (
        _num("Energia_Positiva") * factor,
        _num("Energia_Negativa") * factor,
    )
