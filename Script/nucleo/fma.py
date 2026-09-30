# -*- coding: utf-8 -*-
"""
Subastas!Q (FMA) y Subastas!P (FD).
"""

import pandas as pd

from .externos import desempeno_fd, fma_subastas
from .hojas_entrada import NOMBRES_SUBASTAS
from .lectura import (
    ROL_BALANCE_BESS, ROL_SUBASTAS, TITULOS_DICCIONARIO,
    _bloques_columnas_diccionario, _clave_titulo, mapa_diccionario,
)
from .parametros import (
    ARCHIVO_CENTRALES, CARPETA_FD_FMA, HOJA_DICCIONARIO,
)
from .utiles import (
    ErrorEntrada, _texto_seguro, _tiene_valor, normalizar,
)


# ============================================================
# FMA (Subastas!Q, antes DB!V de la planilla 3)
#
# La formula original de DB!V mira el Concepto de la fila y busca en una
# de tres hojas, que no son un origen: se arman pegando las tres salidas
# de entradas_sscc.py (fma_cpf_*, fma_csf_*, fma_cft_*), que el usuario
# ahora guarda en <CARPETA_BASE>/FD y FMA/. Script/Subastas/Fma.py lee
# esos tres archivos y devuelve tablas normalizadas; aca se hace el
# cruce fila por fila.
#
#   CPF(+) -> FMA_CPF_mas    por Central FMA CPF + Año + Mes + Dia + Hora
#   CPF(-) -> FMA_CPF_menos  idem
#   CSF(+) -> FMA_CSF_mas_base   * Vector de Participacion CSF
#   CSF(-) -> FMA_CSF_menos_base * Vector de Participacion CSF
#   CTF(+) -> FMA_CTF_mas    por unidad/Configuracion + fecha + hora
#   CTF(-) -> FMA_CTF_menos  idem
#   otro   -> 0
#
# Como en la formula original, lo que no se encuentra queda en 0 (el
# IFERROR exterior de DB!V).
# ============================================================

# Titulo del bloque de la hoja Diccionario con la equivalencia
# Configuración -> Central FMA CPF (el documento de trazabilidad la
# llama "diccionario de nomenclatura CPF" y recomienda justamente
# sacarla de la planilla y traerla a Centrales.xlsx). Si el bloque no
# existe, se prueba igual con el nombre tal cual y se avisa.
TITULO_BLOQUE_FMA_CPF = "fma cpf"

# El Vector de Participacion CSF (DB!AC) YA NO ES UN PENDIENTE: sale de
# la hoja "CSF Horario" del archivo SSCC_Desempeño_* (ver
# Script/Fd/Desempeno_Horario.py y el documento de trazabilidad de FD,
# secciones 26 y 27). Este valor se usa solo cuando esa fila no
# aparece en el archivo: la formula original no tiene respaldo para ese
# caso -daria #N/A-, y dejar el FMA en su valor base es mas prudente
# que ponerlo en 0, que seria no pagar. Se avisa en el log cada vez.
VECTOR_PARTICIPACION_CSF_SIN_DATO = 1.0

# Titulo del bloque de la hoja Diccionario con la equivalencia
# Configuración -> unidad como la nombra el archivo de desempeño. Es el
# bloque "FD" que esa hoja ya tenia (el mismo que usa Calculo E
# Costos!AM:AR via construir_dic_mapeo_diccionario).
TITULO_BLOQUE_FD = "fd"

# La formula real de DB!V tiene dos bloques consecutivos para "CTF(+)"
# y, por como estan anidados los SI, el segundo nunca se llega a
# evaluar: CTF(+) termina buscando solo en la primera tabla, mientras
# que CTF(-) si busca en las dos. El documento de trazabilidad pide
# expresamente conservar ese comportamiento en la primera replica y
# recien despues comparar contra una version corregida.
CTF_MAS_BUSCA_EN_LAS_DOS_TABLAS = False


def construir_dic_bloque_diccionario(diccionario, titulo):
    """
    Mapa "como se llama esta central en Subastas" -> "como se llama en
    <titulo>", leyendo la hoja Diccionario de Centrales.xlsx.

    FORMATO NUEVO (tabla unica con encabezados, ver lectura.py): la
    clave sale de la columna "Subastas" Y TAMBIEN de la columna
    "Balance_BESS", porque las dos apuntan a la misma central y quien
    llama busca con la Configuración de la hoja Subastas -que para casi
    todas las centrales coincide con el nombre canonico, pero para
    Tocopilla es BAT_TOCOPILLA-. Cargar las dos es lo que arregla las
    diferencias de FMA CPF y de FD que reporto el usuario: con el
    formato viejo la clave era siempre el nombre canonico y
    BAT_TOCOPILLA no encontraba nada.

    FORMATO VIEJO (bloques lado a lado): se conserva el comportamiento
    de siempre -el bloque cuyo TITULO contiene el texto dado, columna
    izquierda -> columna derecha-. Devuelve {} si no hay ningun bloque
    con ese titulo (es el caso de "FMA CPF", que en el formato viejo
    nunca existio).
    """

    objetivo = _clave_titulo(titulo)
    rol = TITULOS_DICCIONARIO.get(objetivo)

    if rol is not None:

        dic = mapa_diccionario(diccionario, ROL_SUBASTAS, rol)

        if dic:
            # El nombre canonico como clave alternativa, sin pisar lo
            # que ya puso la columna Subastas.
            for clave, valor in mapa_diccionario(
                diccionario, ROL_BALANCE_BESS, rol
            ).items():
                dic.setdefault(clave, valor)

            return dic

    for columnas in _bloques_columnas_diccionario(diccionario):

        if len(columnas) < 2:
            continue

        titulos = [
            _clave_titulo(diccionario.iloc[0, col]) for col in columnas
        ]

        if not any(objetivo in t for t in titulos if t):
            continue

        izquierda, derecha = columnas[0], columnas[1]
        dic = {}

        for indice in range(len(diccionario)):

            clave = normalizar(diccionario.iloc[indice, izquierda])
            valor = diccionario.iloc[indice, derecha]

            if not clave or _clave_titulo(clave) == objetivo:
                continue

            if clave not in dic and _tiene_valor(valor):
                dic[clave] = _texto_seguro(valor)

        return dic

    return {}


def cargar_tablas_fma(carpeta_fd_fma, aamm, registrar=print):
    """
    Lee de <CARPETA_BASE>/FD y FMA/ las tres salidas de FMA del periodo
    y devuelve un dict con las tablas normalizadas. Lo que falte queda
    en None y se avisa: el FMA de ese tipo de servicio quedara en 0,
    igual que hace la formula original cuando no encuentra.
    """

    archivos = fma_subastas.buscar_archivos_fma(carpeta_fd_fma, aamm)
    tablas = {"cpf": None, "csf": None, "ctf_unidad": None,
              "ctf_configuracion": None}

    try:
        if archivos["cpf"]:
            tablas["cpf"] = fma_subastas.tabla_cpf(
                archivos["cpf"], registrar=registrar
            )

        if archivos["csf"]:
            tablas["csf"] = fma_subastas.tabla_csf(
                archivos["csf"], registrar=registrar
            )

        if archivos["ctf"]:
            (
                tablas["ctf_unidad"],
                tablas["ctf_configuracion"],
            ) = fma_subastas.tablas_ctf(
                archivos["ctf"], registrar=registrar
            )

    except fma_subastas.ErrorFma as error:
        raise ErrorEntrada(str(error)) from error

    for tipo in ("cpf", "csf", "ctf"):
        if not archivos[tipo]:
            registrar(
                f"  [AVISO] no se encontro fma_{tipo}_{aamm} en "
                f"{CARPETA_FD_FMA}/: el FMA de las filas {tipo.upper()} "
                f"queda en 0."
            )

    return tablas


def _dic_desde_tabla(tabla, columnas_clave, columnas_valor):
    """
    DataFrame -> dict clave (tupla) -> tupla de valores. Las claves
    numericas se pasan a int para que 3.0 y 3 sean la misma hora.
    """

    if tabla is None or tabla.empty:
        return {}

    dic = {}

    for fila in tabla.itertuples(index=False):
        valores = fila._asdict()

        clave = tuple(
            _clave_numerica_o_texto(valores[columna])
            for columna in columnas_clave
        )

        if clave not in dic:
            dic[clave] = tuple(valores[c] for c in columnas_valor)

    return dic


def _clave_numerica_o_texto(valor):
    """3.0 -> 3; 'El Toro - U1' -> 'el toro - u1'."""

    if isinstance(valor, str):
        return normalizar(valor)

    try:
        if pd.isna(valor):
            return ""
        return int(valor)
    except (TypeError, ValueError):
        return normalizar(valor)


def calcular_fma_subastas(
    df_subastas, tablas, dic_cpf=None, tablas_fd=None, dic_unidad_fd=None,
    registrar=print,
):
    """
    Devuelve la columna FMA (Subastas!Q) para cada fila de Subastas,
    replicando la formula de DB!V (ver el comentario de arriba).

    dic_cpf: Configuración -> nombre de la central como aparece en
    fma_cpf (la equivalencia de nomenclatura). Si no esta, se prueba
    con la Configuración tal cual.

    tablas_fd / dic_unidad_fd: lo que devuelve
    desempeno_fd.construir_tablas_fd() y la equivalencia
    Configuración -> unidad del archivo de desempeño. Se usan para el
    **Vector de Participacion CSF**, que multiplica al FMA de las filas
    CSF (DB!AC). Sin ellos, ese factor queda en 1 y se avisa.
    """

    dic_cpf = dic_cpf or {}
    dic_unidad_fd = dic_unidad_fd or {}
    sin_participacion = 0

    columna_concepto = NOMBRES_SUBASTAS["B"]
    columna_anio = NOMBRES_SUBASTAS["F"]
    columna_mes = NOMBRES_SUBASTAS["G"]
    columna_dia = NOMBRES_SUBASTAS["H"]
    columna_hora = NOMBRES_SUBASTAS["I"]
    columna_config = NOMBRES_SUBASTAS["K"]

    mapa_cpf = _dic_desde_tabla(
        tablas.get("cpf"),
        ["Central", "Año", "Mes", "Dia", "Hora"],
        ["FMA_CPF_mas", "FMA_CPF_menos"],
    )
    mapa_csf = _dic_desde_tabla(
        tablas.get("csf"),
        ["Año", "Mes", "Dia", "Hora"],
        ["FMA_CSF_mas_base", "FMA_CSF_menos_base"],
    )
    mapa_ctf_unidad = _dic_desde_tabla(
        tablas.get("ctf_unidad"),
        ["nombre", "Año", "Mes", "Dia", "Hora"],
        ["FMA_CTF_mas", "FMA_CTF_menos"],
    )
    mapa_ctf_config = _dic_desde_tabla(
        tablas.get("ctf_configuracion"),
        ["nombre", "Año", "Mes", "Dia", "Hora"],
        ["FMA_CTF_mas", "FMA_CTF_menos"],
    )

    valores = []
    sin_dato = {"CPF": 0, "CSF": 0, "CTF": 0, "otro": 0}
    centrales_cpf_sin_equivalencia = set()

    for fila in df_subastas.itertuples(index=False):
        datos = fila._asdict()

        concepto = _texto_seguro(datos[columna_concepto]).upper()
        fecha = (
            _clave_numerica_o_texto(datos[columna_anio]),
            _clave_numerica_o_texto(datos[columna_mes]),
            _clave_numerica_o_texto(datos[columna_dia]),
            _clave_numerica_o_texto(datos[columna_hora]),
        )
        config = _texto_seguro(datos[columna_config])
        config_norm = normalizar(config)

        if concepto.startswith("CPF"):

            central = dic_cpf.get(config_norm)

            if central is None:
                central = config
                if dic_cpf:
                    centrales_cpf_sin_equivalencia.add(config)

            par = mapa_cpf.get((normalizar(central),) + fecha)
            indice = 0 if concepto.endswith("(+)") else 1

            if par is None:
                sin_dato["CPF"] += 1
                valores.append(0.0)
            else:
                valores.append(_numero_o_cero(par[indice]))

        elif concepto.startswith("CSF"):

            par = mapa_csf.get(fecha)
            indice = 0 if concepto.endswith("(+)") else 1

            if par is None:
                sin_dato["CSF"] += 1
                valores.append(0.0)
            else:
                vector = None

                if tablas_fd:
                    unidad = dic_unidad_fd.get(config_norm, config)
                    vector = desempeno_fd.buscar_participacion_csf(
                        tablas_fd, unidad, datos[NOMBRES_SUBASTAS["J"]]
                    )

                if vector is None:
                    vector = VECTOR_PARTICIPACION_CSF_SIN_DATO
                    if tablas_fd:
                        sin_participacion += 1

                valores.append(_numero_o_cero(par[indice]) * float(vector))

        elif concepto.startswith("CTF"):

            clave = (config_norm,) + fecha
            es_mas = concepto.endswith("(+)")
            indice = 0 if es_mas else 1

            par = mapa_ctf_unidad.get(clave)

            # CTF(-) busca en las dos nomenclaturas; CTF(+) solo en la
            # primera (ver CTF_MAS_BUSCA_EN_LAS_DOS_TABLAS).
            if par is None and (
                not es_mas or CTF_MAS_BUSCA_EN_LAS_DOS_TABLAS
            ):
                par = mapa_ctf_config.get(clave)

            if par is None:
                sin_dato["CTF"] += 1
                valores.append(0.0)
            else:
                valores.append(_numero_o_cero(par[indice]))

        else:
            sin_dato["otro"] += 1
            valores.append(0.0)

    if centrales_cpf_sin_equivalencia:
        # Se tolera -el FMA de esas filas queda en 0, igual que el
        # IFERROR de la formula original- pero se dice fuerte: es una
        # celda vacia en el Diccionario, no un dato que no exista.
        registrar(
            f"  [AVISO] {len(centrales_cpf_sin_equivalencia)} central(es) "
            f"sin nombre en la columna '{TITULO_BLOQUE_FMA_CPF.upper()}' "
            f"de la hoja {HOJA_DICCIONARIO} de {ARCHIVO_CENTRALES}: "
            f"{', '.join(repr(c) for c in sorted(centrales_cpf_sin_equivalencia))}. "
            f"Se probo con el nombre tal cual, no se encontro, y el FMA "
            f"de sus filas CPF queda en 0."
        )

    for tipo, cuenta in sin_dato.items():
        if cuenta:
            registrar(
                f"  [AVISO] FMA: {cuenta:,} fila(s) {tipo} sin dato en "
                f"las salidas de FMA, quedan en 0 (igual que la "
                f"formula original)."
            )

    if sin_participacion:
        registrar(
            f"  [AVISO] FMA: {sin_participacion:,} fila(s) CSF sin "
            f"Vector de Participacion en el archivo de desempeño: se "
            f"uso {VECTOR_PARTICIPACION_CSF_SIN_DATO:g}."
        )

    if not tablas_fd:
        registrar(
            "  [AVISO] sin archivo SSCC_Desempeño_*: el FMA de las filas "
            "CSF queda en su valor base (Vector de Participacion = 1)."
        )

    return pd.Series(valores, index=df_subastas.index, dtype="float64")


def calcular_fd_subastas(
    df_subastas, tablas_fd, dic_unidad_fd=None, registrar=print
):
    """
    Devuelve la columna FD (Subastas!P, antes DB!Y) para cada fila,
    siguiendo el documento de trazabilidad de FD:

      CPF -> Fd_CPF de 'CPF Horario'   (probando la alternativa TG/TV)
      CSF -> Fd_CSF de 'CSF Horario'
      CTF -> Fd_CTF de 'CTF Horario'

    buscando por Control + Unidad + Hora_mes. La unidad sale de
    homologar la Configuración contra el bloque "FD" del Diccionario de
    Centrales.xlsx (el mismo que ya usa Calculo E Costos!AM:AR); si no
    esta en el Diccionario se prueba con la Configuración tal cual.

    Lo que no se encuentra queda **vacio** y se cuenta en el log. La
    formula original escribe ahi el texto "ERRORCPF"/"ERRORCSF"/
    "ERRORCTF"; se prefirio dejarlo vacio para no meter texto en una
    columna numerica, pero el log dice cuantas filas fueron.
    """

    dic_unidad_fd = dic_unidad_fd or {}

    columna_control = NOMBRES_SUBASTAS["C"]
    columna_hora_mes = NOMBRES_SUBASTAS["J"]
    columna_config = NOMBRES_SUBASTAS["K"]

    valores = []
    sin_dato = {}
    sin_equivalencia = set()

    for fila in df_subastas.itertuples(index=False):
        datos = fila._asdict()

        control = _texto_seguro(datos[columna_control]).upper()
        config = _texto_seguro(datos[columna_config])
        config_norm = normalizar(config)

        unidad = dic_unidad_fd.get(config_norm)

        if unidad is None:
            unidad = config
            if dic_unidad_fd:
                sin_equivalencia.add(config)

        valor = desempeno_fd.buscar_fd(
            tablas_fd, control, unidad, datos[columna_hora_mes]
        )

        if valor is None:
            sin_dato[control] = sin_dato.get(control, 0) + 1
            valores.append(pd.NA)
        else:
            valores.append(valor)

    for central in sorted(sin_equivalencia):
        registrar(
            f"  [AVISO] sin equivalencia de nomenclatura FD para "
            f"'{central}' (bloque '{TITULO_BLOQUE_FD}' de "
            f"{HOJA_DICCIONARIO} en {ARCHIVO_CENTRALES}): se probo con "
            f"el nombre tal cual."
        )

    for control, cuenta in sorted(sin_dato.items()):
        registrar(
            f"  [AVISO] FD: {cuenta:,} fila(s) {control} sin dato en el "
            f"archivo de desempeño, quedan vacias."
        )

    return pd.Series(valores, index=df_subastas.index, dtype="Float64")


def _numero_o_cero(valor):
    try:
        if pd.isna(valor):
            return 0.0
        return float(valor)
    except (TypeError, ValueError):
        return 0.0
