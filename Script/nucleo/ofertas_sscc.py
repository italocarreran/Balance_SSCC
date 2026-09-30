# -*- coding: utf-8 -*-
"""
Ofertas SSCC y las columnas R, S, T, V de Medidores.
"""

import calendar
import pandas as pd
from pathlib import Path

from .lectura import (
    ROL_BALANCE_BESS, ROL_OFERTAS, ROL_SUBASTAS, filas_diccionario,
)
from .parametros import HOJA_OFERTAS_SSCC, INICIO_VENTANA
from .utiles import (
    ErrorEntrada, _es_numero, _texto_seguro, _tiene_valor, _valor_clave,
)


def _limpiar_nombre_mostrar(valor):
    """Replica OSSCC_LimpiarNombreMostrar: NBSP->espacio, trim, espacios colapsados."""

    if not _tiene_valor(valor):
        return ""

    texto = str(valor).replace("\xa0", " ").strip()

    while "  " in texto:
        texto = texto.replace("  ", " ")

    return texto


def _normalizar_nombre_clave(valor):
    """Replica OSSCC_NormalizarNombreClave: LimpiarNombreMostrar + mayusculas."""

    return _limpiar_nombre_mostrar(valor).upper()


def _contiene_bess_o_sae(texto):
    """Replica OSSCC_ContieneBESSoSAE."""

    texto = texto.upper()
    return "BESS" in texto or "SAE" in texto or "BAT" in texto


def _servicio_termina_en_rs(servicio):
    """Replica OSSCC_ServicioTerminaEnRS."""

    return servicio.strip().upper().endswith("_RS")


def _es_respuesta_si(valor):
    """Replica OSSCC_EsRespuestaSi: normaliza y compara contra 'SI'."""

    if not _tiene_valor(valor):
        return False

    texto = str(valor).strip().upper()
    texto = texto.replace("\xa0", "").replace(" ", "").replace("\t", "")

    for con_tilde, sin_tilde in (
        ("Í", "I"), ("Ì", "I"), ("Ï", "I"), ("Î", "I"),
    ):
        texto = texto.replace(con_tilde, sin_tilde)

    for signo in ".,;:":
        texto = texto.replace(signo, "")

    return texto == "SI"


def _normalizar_periodo(valor):
    """
    Replica OSSCC_NormalizarPeriodo: interpreta 'valor' como un
    periodo horario 1..24. Acepta enteros 1..24, fracciones de dia
    estilo Excel, '24:00' y texto de hora/fecha reconocible. None si
    no se puede interpretar.
    """

    if not _tiene_valor(valor):
        return None

    if _es_numero(valor):
        numero = float(valor)

        if numero == int(numero) and 1 <= numero <= 24:
            return int(numero)

        if 0 <= numero < 1:
            total_segundos = round(numero * 86400)
            hora = (total_segundos // 3600) % 24
            return int(hora) + 1

        return None

    texto = str(valor).strip()

    if texto in ("24:00", "24:00:00"):
        return 24

    try:
        marca = pd.to_datetime(texto)
        return int(marca.hour) + 1
    except (ValueError, TypeError):
        return None


def _valor_oferta_binario(valor):
    """Replica OSSCC_ValorOfertaBinario: 1 solo si el valor es exactamente 1."""

    if not _tiene_valor(valor):
        return 0

    if _es_numero(valor):
        return 1 if float(valor) == 1 else 0

    return 1 if str(valor).strip() == "1" else 0


def construir_resumen_ofertas_sscc(ruta_ofertas, registrar=print):
    """
    Replica Generar_Resumen_Ofertas_SSCC.

    Lee TODAS las hojas de ruta_ofertas (columnas A:I, desde la fila
    2), filtra filas cuyo nombre (A) contenga BESS/SAE/BAT y cuyo
    servicio (H) termine en "_RS", y agrupa por (Nombre, Año, Mes,
    Día). Para cada grupo, una columna por cada servicio _RS
    encontrado en TODO el archivo (1 si en ese grupo el servicio tiene
    ofertado ("Sí") las 24 horas, 0 si no) y una columna final "Oferta
    completa" (1 solo si todos los servicios del grupo estan
    completos).

    Devuelve un DataFrame ordenado por Nombre, Año, Mes, Día.
    """

    hojas = pd.read_excel(ruta_ofertas, sheet_name=None, header=None)

    grupos = {}
    datos_grupo = {}
    servicios_globales = set()

    for df_hoja in hojas.values():

        if df_hoja.shape[0] < 2 or df_hoja.shape[1] < 9:
            continue

        for _, fila in df_hoja.iloc[1:, 0:9].iterrows():

            nombre = _texto_seguro(fila.iloc[0])
            anio, mes, dia = fila.iloc[1], fila.iloc[2], fila.iloc[3]
            servicio = _texto_seguro(fila.iloc[7])
            indicador = fila.iloc[8]

            if not _contiene_bess_o_sae(nombre):
                continue

            if not _servicio_termina_en_rs(servicio):
                continue

            if not (
                _tiene_valor(anio)
                and _tiene_valor(mes)
                and _tiene_valor(dia)
            ):
                continue

            clave_grupo = (
                nombre,
                _valor_clave(anio),
                _valor_clave(mes),
                _valor_clave(dia),
            )

            if clave_grupo not in grupos:
                grupos[clave_grupo] = {}
                datos_grupo[clave_grupo] = (nombre, anio, mes, dia)

            servicios_globales.add(servicio)

            horas_servicio = grupos[clave_grupo].setdefault(
                servicio, set()
            )

            if _es_respuesta_si(indicador):
                periodo = _normalizar_periodo(fila.iloc[4])
                if periodo is not None:
                    horas_servicio.add(periodo)

    if not grupos:
        raise ErrorEntrada(
            "No se encontraron registros que cumplan las condiciones "
            "BESS/SAE/BAT y servicio terminado en _RS en "
            f"{Path(ruta_ofertas).name}"
        )

    lista_servicios = sorted(servicios_globales, key=str.upper)

    filas_salida = []

    for clave_grupo, servicios_grupo in grupos.items():

        nombre, anio, mes, dia = datos_grupo[clave_grupo]
        fila_salida = {
            "Nombre": nombre,
            "Año": anio,
            "Mes": mes,
            "Día": dia,
        }

        todos_completos = True

        for servicio in lista_servicios:
            completo = len(servicios_grupo.get(servicio, set())) == 24
            fila_salida[servicio] = 1 if completo else 0
            if not completo:
                todos_completos = False

        fila_salida["Oferta completa"] = (
            1 if todos_completos and lista_servicios else 0
        )

        filas_salida.append(fila_salida)

    columnas = (
        ["Nombre", "Año", "Mes", "Día"]
        + lista_servicios
        + ["Oferta completa"]
    )

    df_resumen = pd.DataFrame(filas_salida, columns=columnas)

    df_resumen["_orden_nombre"] = df_resumen["Nombre"].str.upper()
    df_resumen = (
        df_resumen
        .sort_values(by=["_orden_nombre", "Año", "Mes", "Día"])
        .drop(columns="_orden_nombre")
        .reset_index(drop=True)
    )

    registrar(
        f"  Resumen Ofertas SSCC: {len(df_resumen):,} registros, "
        f"{len(lista_servicios)} servicio(s) _RS: "
        f"{', '.join(lista_servicios)}"
    )

    return df_resumen


def cargar_resumen_en_medidores(
    df_resumen, claves_medidores, diccionario, registrar=print
):
    """
    Replica OSSCC_CargarResumenEnMedidores.

    Arma la tabla equivalente a Medidores!W:Y (Nombre, Dia, Oferta
    completa) con una fila por central x dia del mes: toma los
    nombres y ofertas de df_resumen, y agrega ademas -con oferta 0
    para todos sus dias- cualquier central de Medidores!clave que no
    este representada ahi ni mediante una equivalencia de
    Diccionario!E:F:G.

    claves_medidores: valores unicos de la columna 'clave' de Medidores.

    Devuelve (df_wxy, (anio, mes), avisos).
    """

    if "Oferta completa" not in df_resumen.columns:
        raise ErrorEntrada(
            "El resumen de Ofertas SSCC no tiene la columna "
            "'Oferta completa'."
        )

    dic_nombres = {}
    dic_ofertas = {}
    periodos = set()

    for _, fila in df_resumen.iterrows():

        nombre_mostrar = _limpiar_nombre_mostrar(fila["Nombre"])
        if not nombre_mostrar:
            continue

        anio, mes, dia = fila["Año"], fila["Mes"], fila["Día"]

        if not (_es_numero(anio) and _es_numero(mes) and _es_numero(dia)):
            continue

        anio_i, mes_i, dia_i = int(anio), int(mes), int(dia)

        if not 1 <= mes_i <= 12:
            continue

        periodos.add((anio_i, mes_i))

        clave_nombre = _normalizar_nombre_clave(nombre_mostrar)
        dic_nombres.setdefault(clave_nombre, nombre_mostrar)

        oferta = _valor_oferta_binario(fila["Oferta completa"])
        dic_ofertas[(clave_nombre, dia_i)] = oferta

    if len(periodos) == 0:
        raise ErrorEntrada(
            "No fue posible determinar el año y mes del resumen de "
            "Ofertas SSCC."
        )

    if len(periodos) > 1:
        raise ErrorEntrada(
            "El resumen de Ofertas SSCC contiene mas de un año o "
            f"mes: {sorted(periodos)}. No es posible construir "
            "Medidores!W:Y porque el destino solo admite un periodo "
            "por vez."
        )

    anio, mes = next(iter(periodos))
    dias_del_mes = calendar.monthrange(anio, mes)[1]

    # Equivalencias Diccionario!E:F:G (columnas 5,6,7 -> indices 4,5,6)
    dic_equivalencias = {}

    if diccionario.shape[1] > 6:
        for _, fila in diccionario.iterrows():
            alias = []
            for indice in (4, 5, 6):
                valor = _normalizar_nombre_clave(fila.iloc[indice])
                if valor and valor not in alias:
                    alias.append(valor)
            for valor in alias:
                dic_equivalencias[valor] = set(alias)

    dic_no_encontrados = {}

    for nombre_crudo in claves_medidores:

        nombre_mostrar = _limpiar_nombre_mostrar(nombre_crudo)
        if not nombre_mostrar:
            continue

        clave_nombre = _normalizar_nombre_clave(nombre_mostrar)

        if clave_nombre in dic_nombres:
            continue

        equivalentes = dic_equivalencias.get(clave_nombre)
        ya_representado = False

        if equivalentes:
            ya_representado = any(
                alias in dic_nombres for alias in equivalentes
            )
        else:
            dic_no_encontrados.setdefault(clave_nombre, nombre_mostrar)

        if not ya_representado:
            dic_nombres.setdefault(clave_nombre, nombre_mostrar)

    if not dic_nombres:
        raise ErrorEntrada(
            "No se encontraron nombres en el resumen de Ofertas SSCC "
            "ni en Medidores (columna clave)."
        )

    avisos = []

    if dic_no_encontrados:
        nombres_avisados = sorted(
            dic_no_encontrados.values(), key=str.upper
        )
        cola = (
            f" ... y {len(nombres_avisados) - 30} mas."
            if len(nombres_avisados) > 30 else ""
        )
        avisos.append(
            f"{len(nombres_avisados)} nombre(s) de Medidores!clave no "
            "se encontraron en Diccionario!E:F:G. Se incorporaron con "
            "oferta 0: " + ", ".join(nombres_avisados[:30]) + cola
        )

    nombres_ordenados = sorted(
        dic_nombres.items(), key=lambda kv: kv[1].upper()
    )

    filas_salida = []
    for clave_nombre, nombre in nombres_ordenados:
        for dia in range(1, dias_del_mes + 1):
            oferta = dic_ofertas.get((clave_nombre, dia), 0)
            filas_salida.append((nombre, dia, oferta))

    df_wxy = pd.DataFrame(
        filas_salida, columns=["Nombre", "Dia", "Oferta completa"]
    )

    registrar(
        f"  Ofertas SSCC por dia (equivalente a Medidores!W:Y): "
        f"{len(df_wxy):,} filas ({len(nombres_ordenados)} "
        f"central(es) x {dias_del_mes} dias)"
    )

    return df_wxy, (anio, mes), avisos


def _mapas_homologacion_fge(diccionario):
    """
    Precalcula los dos mapas "nombre de un origen -> nombre canonico"
    que usa la formula de Medidores!V. Ante nombres repetidos se
    conserva el primero, igual que XLOOKUP con la primera coincidencia.

    FORMATO NUEVO de la hoja Diccionario (tabla unica con encabezados,
    ver lectura.py): mapa_f sale de la columna "Subastas" y mapa_g de
    la columna "Ofertas", las dos hacia "Balance_BESS". Es exactamente
    el mismo par de mapas que daba el formato viejo, solo que las
    columnas ahora se encuentran por su nombre en vez de por su
    posicion.

    FORMATO VIEJO (bloques lado a lado): Diccionario!F y Diccionario!G
    (columnas 6 y 7, indices 5 y 6) hacia Diccionario!E (indice 4).
    """

    filas = filas_diccionario(diccionario)

    if filas:

        mapa_f = {}
        mapa_g = {}

        for fila in filas:

            valor_e = _texto_seguro(fila.get(ROL_BALANCE_BESS))

            for mapa, rol in ((mapa_f, ROL_SUBASTAS), (mapa_g, ROL_OFERTAS)):
                clave = _normalizar_nombre_clave(fila.get(rol))
                if clave and clave not in mapa:
                    mapa[clave] = valor_e

        return mapa_f, mapa_g

    mapa_f = {}
    mapa_g = {}

    if diccionario.shape[1] > 6:
        for _, fila in diccionario.iterrows():
            valor_e = _texto_seguro(fila.iloc[4])
            clave_f = _normalizar_nombre_clave(fila.iloc[5])
            clave_g = _normalizar_nombre_clave(fila.iloc[6])

            if clave_f and clave_f not in mapa_f:
                mapa_f[clave_f] = valor_e

            if clave_g and clave_g not in mapa_g:
                mapa_g[clave_g] = valor_e

    return mapa_f, mapa_g


def _homologar_fge(nombre, mapa_f, mapa_g):
    """
    Replica el homologador de Medidores!V:
    XLOOKUP(nombre, Diccionario!F, Diccionario!E,
        XLOOKUP(nombre, Diccionario!G, Diccionario!E, 0, 0), 0)
    """

    clave = _normalizar_nombre_clave(nombre)

    if clave in mapa_f:
        return mapa_f[clave]

    if clave in mapa_g:
        return mapa_g[clave]

    return "0"


def calcular_r(df_medidores, df_wxy, diccionario, registrar=print):
    """
    Replica Medidores!R = VLOOKUP(B&G, V:Y, 4, FALSE).

    V (la tabla auxiliar, aca solo en memoria) es Dia & homologado(W)
    via Diccionario!F/G->E, construida a partir de df_wxy (equivalente
    a Medidores!W:Y). El cruce se hace por Dia + nombre homologado en
    mayusculas, separados por '|' (Excel concatena sin separador; acá
    se agrega uno para no confundir, p.ej., dia=1+'0ABC' con
    dia=10+'ABC' - un caso limite que en Excel tampoco se distingue).

    Devuelve (serie_r, avisos).
    """

    mapa_f, mapa_g = _mapas_homologacion_fge(diccionario)

    tabla_v = {}
    for _, fila in df_wxy.iterrows():
        homologado = _homologar_fge(fila["Nombre"], mapa_f, mapa_g)
        clave_v = f"{int(fila['Dia'])}|{homologado.upper()}"
        tabla_v.setdefault(clave_v, fila["Oferta completa"])

    claves_medidores = (
        df_medidores["Dia"].astype("Int64").astype(str)
        + "|"
        + df_medidores["clave"].astype(str).str.strip().str.upper()
    )

    r = claves_medidores.map(tabla_v).astype("Int64")

    avisos = []
    no_encontrados = int(r.isna().sum())

    if no_encontrados:
        avisos.append(
            f"{no_encontrados:,} fila(s) de Medidores no encontraron "
            "coincidencia (Dia + central homologada) en la tabla de "
            "Ofertas SSCC al calcular la columna R. Revisar que todas "
            "las centrales de Medidores!clave esten en la hoja "
            "Diccionario de Centrales.xlsx (columnas Subastas y "
            "Ofertas hacia Balance_BESS; en el formato viejo, E:F:G)."
        )

    registrar(f"  Columna R (Oferta_Completa_Dia): {no_encontrados:,} sin match")

    return r, avisos


def calcular_s(ventana, r_valor):
    """
    Replica Medidores!S:
        =IF(L3=L2, S2, IF(R3=1, 1, 2))

    S se mantiene igual a la fila anterior mientras la Ventana (L) no
    cambie; al cambiar, se recalcula segun si R vale 1. No se reinicia
    aparte por central: en la planilla original tampoco lo hace, se
    apoya en que L ya cambia al cambiar de central.
    """

    ventana = pd.Series(ventana).reset_index(drop=True)
    r_valor = pd.Series(r_valor).reset_index(drop=True)

    cambia_ventana = ventana.ne(ventana.shift())

    # R sin match (ver calcular_r) se trata como "no es 1", igual que
    # cualquier valor de R distinto de 1 en la formula de Excel.
    es_uno = r_valor.fillna(-1).eq(1)
    valor_si_cambia = es_uno.map({True: 1, False: 2})

    s = valor_si_cambia.where(cambia_ventana)

    return s.ffill().astype("Int64")


def construir_resumen_ventana_oferta(
    clave, ventana, oferta_r, inicio_ventana=INICIO_VENTANA, registrar=print
):
    """
    Replica Resumir_Medidores_Central_Ventana_Oferta_Completa.

    Agrupa Medidores por (clave=Central, Ventana=L), suma R, y marca
    "Completa" segun la oferta esperada para esa ventana: 96 en las
    ventanas normales, (inicio_ventana-1)*4 en la ventana 0, y
    (25-inicio_ventana)*4 en la ultima ventana del mes (la de mayor
    valor numerico encontrado en L).

    Devuelve un DataFrame con columnas Central, Ventana T, Oferta,
    Completa.
    """

    if not 1 <= inicio_ventana <= 24:
        raise ErrorEntrada(
            "INICIO_VENTANA debe ser una hora entre 1 y 24."
        )

    oferta_esperada_inicial = (inicio_ventana - 1) * 4
    oferta_esperada_final = (25 - inicio_ventana) * 4

    ventana_numerica = pd.to_numeric(pd.Series(ventana), errors="coerce")
    ultima_ventana = ventana_numerica.max()

    if pd.isna(ultima_ventana):
        raise ErrorEntrada(
            "No se encontraron valores numericos en la columna "
            "Ventana (L) para resumir Ofertas SSCC."
        )

    df = pd.DataFrame(
        {
            "Central": pd.Series(clave).reset_index(drop=True),
            "Ventana T": ventana_numerica.reset_index(drop=True),
            "Oferta": pd.to_numeric(
                pd.Series(oferta_r), errors="coerce"
            ).reset_index(drop=True),
        }
    )
    df = df.dropna(subset=["Central", "Ventana T"])

    resumen = (
        df.groupby(["Central", "Ventana T"], as_index=False)["Oferta"]
        .sum()
    )

    def oferta_esperada(v):
        if abs(v) < 1e-6:
            return oferta_esperada_inicial
        if abs(v - ultima_ventana) < 1e-6:
            return oferta_esperada_final
        return 96

    esperada = resumen["Ventana T"].map(oferta_esperada)
    resumen["Completa"] = (
        (resumen["Oferta"] - esperada).abs() < 1e-6
    ).astype("int64")

    resumen = (
        resumen
        .sort_values(by=["Central", "Ventana T"])
        .reset_index(drop=True)
    )

    registrar(
        f"  Resumen Ventana Oferta: {len(resumen):,} grupo(s) "
        f"central x ventana; ultima ventana detectada "
        f"{ultima_ventana:g}"
    )

    return resumen


def calcular_t(clave, ventana, resumen_ventana_oferta):
    """
    Replica Medidores!T = 1 - Completa(central=G, ventana=L), buscando
    en el resumen central+ventana+oferta completa (AB:AE).

    Devuelve (serie_t, cantidad_sin_match).
    """

    df = pd.DataFrame(
        {
            "Central": pd.Series(clave).reset_index(drop=True),
            "Ventana T": pd.Series(ventana).reset_index(drop=True),
        }
    )

    cruzado = df.merge(
        resumen_ventana_oferta[["Central", "Ventana T", "Completa"]],
        on=["Central", "Ventana T"],
        how="left",
    )

    t = (1 - cruzado["Completa"]).astype("Int64")

    return t, int(t.isna().sum())


# ============================================================
# LA HOJA "Ofertas SSCC" DEL CONSOLIDADO
# ============================================================

# Titulo de cada una de las dos tablas de la hoja. Los usa escritura.py
# para escribirlas y leer_ofertas_sscc_consolidado() para encontrarlas:
# si alguien cambia uno, las dos puntas se mueven juntas. El nombre
# de la hoja sale de parametros.py (ORDEN_HOJAS_SALIDA), y se
# reexporta aca porque medio modulo la busca por este nombre.
TITULO_OFERTAS_POR_DIA = (
    "Ofertas SSCC por dia (equivalente a Medidores!W:Y)"
)
TITULO_RESUMEN_VENTANA = (
    "Resumen ventana oferta (equivalente a Medidores!AB:AE)"
)


def leer_ofertas_sscc_consolidado(ruta_consolidado, registrar=print, libro=None):
    """
    Las dos tablas de la hoja "Ofertas SSCC" de
    la planilla de salida: (df_wxy, df_resumen_ventana).

    libro: un pd.ExcelFile ya abierto del consolidado. Se pasa cuando
    hay que leer VARIAS hojas del mismo archivo (generar_pagos_bess
    lee cuatro): sin esto, cada pd.read_excel vuelve a parsear el
    libro entero, que es el archivo mas grande del caso.

    La hoja las guarda una al lado de la otra, cada una con su titulo
    arriba y su propia fila de encabezados debajo (ver
    _escribir_tabla_con_titulo). Se las ubica por ese titulo, no por
    una posicion fija de columna: si maniana se agrega una tabla mas o
    cambia el ancho de la primera, esto sigue andando.
    """

    ruta_consolidado = Path(ruta_consolidado)
    fuente = ruta_consolidado if libro is None else libro

    try:
        crudo = pd.read_excel(
            fuente, sheet_name=HOJA_OFERTAS_SSCC, header=None
        )
    except ValueError as error:
        raise ErrorEntrada(
            f"{ruta_consolidado.name} no tiene la hoja "
            f"'{HOJA_OFERTAS_SSCC}' todavia. Genera "
            f"esa hoja primero (tildando "
            f"'Ofertas SSCC')."
        ) from error

    if crudo.empty:
        raise ErrorEntrada(
            f"La hoja '{HOJA_OFERTAS_SSCC}' de {ruta_consolidado.name} "
            f"esta vacia. Actualizala con su boton en la ventana."
        )

    def _tabla(titulo):

        columna = None

        for col in range(crudo.shape[1]):
            if _texto_seguro(crudo.iloc[0, col]) == titulo:
                columna = col
                break

        if columna is None:
            raise ErrorEntrada(
                f"La hoja '{HOJA_OFERTAS_SSCC}' de "
                f"{ruta_consolidado.name} no tiene la tabla "
                f"'{titulo}'. Actualizala con su boton en la ventana."
            )

        # El ancho es el bloque de encabezados no vacios que arranca en
        # esa columna (las tablas quedan separadas por columnas vacias).
        ancho = 0
        while (
            columna + ancho < crudo.shape[1]
            and _texto_seguro(crudo.iloc[1, columna + ancho])
        ):
            ancho += 1

        encabezados = [
            _texto_seguro(crudo.iloc[1, columna + i]) for i in range(ancho)
        ]

        datos = crudo.iloc[2:, columna:columna + ancho].reset_index(drop=True)
        datos = datos.set_axis(encabezados, axis=1)

        # La tabla mas corta viene con filas de relleno vacias abajo
        # (las que ocupa la otra): se cortan.
        con_datos = datos.map(lambda v: bool(_texto_seguro(v))).any(axis=1)
        ultima = con_datos[con_datos].index.max()

        if pd.isna(ultima):
            return datos.iloc[:0]

        return datos.iloc[: int(ultima) + 1]

    df_wxy = _tabla(TITULO_OFERTAS_POR_DIA)
    df_resumen_ventana = _tabla(TITULO_RESUMEN_VENTANA)

    # Vuelven de Excel como texto/object: las columnas con las que se
    # cruza contra Medidores tienen que ser numeros otra vez.
    for columna in ("Dia", "Oferta completa"):
        if columna in df_wxy.columns:
            df_wxy[columna] = pd.to_numeric(df_wxy[columna], errors="coerce")

    for columna in ("Ventana T", "Oferta", "Completa"):
        if columna in df_resumen_ventana.columns:
            df_resumen_ventana[columna] = pd.to_numeric(
                df_resumen_ventana[columna], errors="coerce"
            )

    registrar(
        f"  Ofertas SSCC desde el consolidado: {len(df_wxy):,} fila(s) "
        f"central x dia, {len(df_resumen_ventana):,} grupo(s) "
        f"central x ventana."
    )

    return df_wxy, df_resumen_ventana
