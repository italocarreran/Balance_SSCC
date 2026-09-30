# -*- coding: utf-8 -*-
"""
El arbol de carpetas y archivos que dibuja la ventana.
"""

import openpyxl
import pandas as pd
from pathlib import Path

from .externos import (
    Homologacion, extrae_cmg, fma_subastas, ofertas_adj,
)
from .parametros import (
    ARCHIVO_CENTRALES, ARCHIVO_CMG, ARCHIVO_CONTROL, ARCHIVO_MEDIDAS_SAE,
    ARCHIVO_SALIDA, CARPETA_AUXILIARES, CARPETA_CMG,
    CARPETA_DB_SUBASTAS, CARPETA_FD_FMA, CARPETA_MEDIDAS,
    CARPETA_OFERTAS, CARPETA_PRORRATA_RETIROS, CARPETA_SUBASTAS,
    HOJA_CALCULO_ECOSTOS, HOJA_CALCULO_RE545, HOJA_CMG,
    HOJA_COMPENSACION_CENTRAL, HOJA_DICCIONARIO, HOJA_FD,
    HOJA_MEDIDORES, HOJA_OFERTAS_SSCC, HOJA_PRORRATA_RETIROS,
    HOJA_RESUMEN, HOJA_RESUMEN_BESS, HOJA_SUBASTAS, ORDEN_HOJAS_SALIDA,
)
from .origenes import origen as origen_de
from .prorrata_retiros import buscar_archivo_prorrata
from .rutas import (
    buscar_archivo_ofertas, buscar_archivo_sscc_desempeno,
    buscar_soc, resolver_rutas, validar_aamm,
)
from .utiles import ErrorEntrada, normalizar


# ============================================================
# VALIDACION DE ESTRUCTURA
# ============================================================

# Lo que ya se leyo de cada archivo, por (ruta, fecha de
# modificacion, tamaño): la ventana llama a revisar_estructura() en
# cada repintado (al elegir carpeta, al cambiar el AAMM, al terminar
# cualquier boton), y sin esto cada repintado volvia a ABRIR
# Centrales.xlsx, el archivo de homologacion y las dos salidas -que
# son las planillas grandes- solo para preguntarles que hojas tienen.
# Pedido del usuario: "que no se abran planillas innecesarias".
#
# La clave incluye mtime y tamaño: si el archivo cambia (lo regenero
# el programa, o el usuario lo edito en Excel), la entrada vieja deja
# de servir sola.
_CACHE_HOJAS = {}
_CACHE_MAXIMO = 64


def _firma(ruta):
    """(ruta, mtime, tamaño) o None si el archivo no esta."""

    try:
        est = ruta.stat()
    except OSError:
        return None

    return (str(ruta), est.st_mtime_ns, est.st_size)


def _con_cache(ruta, etiqueta, leer):
    """
    leer(ruta) una sola vez por version del archivo. 'etiqueta'
    separa dos lecturas distintas del mismo archivo (los nombres de
    hoja vs. que hojas tienen datos).
    """

    firma = _firma(ruta)

    if firma is None:
        return None

    clave = (etiqueta,) + firma

    if clave not in _CACHE_HOJAS:
        if len(_CACHE_HOJAS) >= _CACHE_MAXIMO:
            _CACHE_HOJAS.clear()
        _CACHE_HOJAS[clave] = leer(ruta)

    return _CACHE_HOJAS[clave]


def hojas_de(ruta):
    """
    Nombres de hoja de un Excel, o None si no se puede abrir (no
    existe, esta abierto por Excel, corrupto). Se usa para mostrar el
    estado hoja por hoja de las dos SALIDAS, que se generan por
    partes: cada hoja puede estar o no estar.

    Se cachea por version del archivo (ver _con_cache).
    """

    ruta = Path(ruta)

    if not ruta.is_file():
        return None

    def _leer(ruta):
        try:
            with pd.ExcelFile(ruta) as libro:
                return list(libro.sheet_names)
        except Exception:
            return None

    return _con_cache(ruta, "hojas", _leer)


def _fila(id_fila, etiqueta, nivel, estado, detalle="", ruta=None,
          es_carpeta=False, origen=None):
    """
    Una fila del diagrama de la ventana.

    id_fila: identificador estable (la ventana lo usa para saber que
    boton va en que fila; nucleo no sabe nada de botones).
    nivel: 0 = raiz del caso, 1 = adentro de una carpeta/archivo,
    2 = adentro de un archivo que esta adentro de una carpeta. Lo
    decide nucleo porque es estructura, no presentacion: como
    dibujarlo (prefijos, colores) es cosa de la ventana.
    ruta: la ruta de ESA fila en el disco; la ventana la usa para que
    el nombre sea un link. Siempre se abre una CARPETA (la propia si la
    fila es una carpeta, la que contiene al archivo si es un archivo):
    nunca se abre el archivo, para no arrancar Excel sin que se lo
    pidan.
    es_carpeta: si la fila es una carpeta (y no un archivo).
    origen: id de Script/nucleo/origenes.py cuando lo que hay en esa
    fila viene de afuera del caso (o se arma con insumos que vienen de
    afuera). La ventana lo muestra como "Origen: <etiqueta>" en el
    detalle, con la etiqueta como link.
    """

    return {
        "id": id_fila,
        "etiqueta": etiqueta,
        "nivel": nivel,
        "estado": estado,
        "detalle": detalle,
        "ruta": str(ruta) if ruta else "",
        "es_carpeta": bool(es_carpeta),
        "origen": origen_de(origen) if origen else None,
    }


def hojas_con_datos(ruta):
    """
    {nombre de hoja: tiene datos} de un Excel, o None si no se puede
    abrir. "Tiene datos" = mas de una fila usada: una hoja preservada
    que nunca se genero queda con una sola celda vacia (ver
    _preservar_o_avisar en escribir_salida), y en el diagrama tiene
    que verse como PENDIENTE, no como generada.
    """

    ruta = Path(ruta)

    if not ruta.is_file():
        return None

    def _leer(ruta):
        try:
            libro = openpyxl.load_workbook(ruta, read_only=True)
        except Exception:
            return None

        try:
            return {
                hoja.title: (hoja.max_row or 0) > 1
                for hoja in libro.worksheets
            }
        finally:
            libro.close()

    return _con_cache(ruta, "hojas_con_datos", _leer)


def _filas_de_hojas(ruta_archivo, secciones, nivel):
    """
    Una fila por hoja de la planilla de salida: 'ok' si la hoja ya
    existe en el archivo, 'pendiente' si todavia no se genero.

    secciones: SECCIONES_SALIDA, o sea las dos mitades (entradas y
    calculo) ya en el orden en que quedan las hojas del libro -de fin
    a inicio-, cada una con el prefijo de id que dice que boton la
    genera ("consolidado" / "pagos").

    La salida se desglosa como el resto del arbol -el archivo como
    "carpeta", sus hojas adentro- y cada hoja trae su propio boton en
    la ventana: por eso no hace falta ninguna ventana intermedia para
    elegir que recalcular.
    """

    hojas = hojas_con_datos(ruta_archivo)
    con_datos = (
        {normalizar(nombre) for nombre, tiene in hojas.items() if tiene}
        if hojas is not None else set()
    )

    filas = []

    for prefijo_id, (id_seccion, etiqueta, _, nombres_hoja) in secciones:

        presentes = [
            nombre for nombre in nombres_hoja
            if normalizar(nombre) in con_datos
        ]

        if hojas is None:
            estado, detalle = "pendiente", "todavia no generada"
        elif len(presentes) == len(nombres_hoja):
            estado, detalle = "ok", "generada"
        elif presentes:
            estado, detalle = "pendiente", "generada a medias"
        else:
            estado, detalle = "pendiente", "todavia no generada"

        filas.append(
            _fila(
                f"{prefijo_id}:{id_seccion}",
                f"hoja '{etiqueta}'",
                nivel,
                estado,
                detalle,
            )
        )

    return filas


def revisar_estructura(carpeta_base, aamm=None):
    """
    Revisa la carpeta base y devuelve (rutas, filas), donde cada fila
    es el dict que arma _fila(): id, etiqueta, nivel, estado
    ('ok'/'falta'/'pendiente') y detalle.

    aamm: periodo ingresado por el usuario en la ventana (4 digitos,
    ej. '2607'). Sin un AAMM valido no se pueden buscar ni el SoC ni
    el CSV de CMg; eso se dice en el detalle de ESAS filas, no en una
    fila propia del periodo (el AAMM se ingresa arriba, en su campo,
    y no es parte de la estructura de carpetas).
    """

    rutas = resolver_rutas(carpeta_base)
    filas = []

    def agregar(id_fila, etiqueta, nivel, existe, detalle="", ruta=None,
                es_carpeta=False, origen=None):
        filas.append(
            _fila(
                id_fila, etiqueta, nivel, "ok" if existe else "falta",
                detalle, ruta=ruta, es_carpeta=es_carpeta, origen=origen,
            )
        )

    try:
        aamm_valido = validar_aamm(aamm)
    except ErrorEntrada:
        aamm_valido = None

    rutas["aamm"] = aamm_valido

    agregar(
        "base", "Carpeta base", 0, rutas["base"].is_dir(),
        str(rutas["base"]), ruta=rutas["base"], es_carpeta=True,
    )

    # ---- Medidas/ -------------------------------------------------
    agregar(
        "medidas_dir", f"{CARPETA_MEDIDAS}/", 0,
        rutas["medidas_dir"].is_dir(),
        ruta=rutas["medidas_dir"], es_carpeta=True,
    )

    # Medidas_SAE.xlsx ya no se deja a mano: lo arma el programa desde
    # las dos APIs del Coordinador (ver generar_medidas_sae).
    filas.append(
        _fila(
            "medidas_sae",
            ARCHIVO_MEDIDAS_SAE,
            1,
            "ok" if rutas["medidas_sae"].is_file() else "falta",
            (
                "se regenera con el boton ->"
                if rutas["medidas_sae"].is_file()
                else "se genera con el boton -> (baja las medidas del mes)"
            ),
            ruta=rutas["medidas_sae"],
        )
    )

    # El SoC del periodo vive aca adentro (nivel 1), no es una entrada
    # suelta: su nombre solo tiene que contener "SOC" y el AAMM.
    if aamm_valido:
        try:
            archivo_soc = buscar_soc(rutas["medidas_dir"], aamm_valido)
            filas.append(
                _fila(
                    "soc", archivo_soc.name, 1, "ok",
                    f"periodo {aamm_valido}", ruta=archivo_soc,
                )
            )
            rutas["soc"] = archivo_soc
        except ErrorEntrada as error:
            filas.append(
                _fila(
                    "soc",
                    f"SoC del periodo {aamm_valido}",
                    1,
                    "falta",
                    str(error).split("\n")[0],
                    ruta=rutas["medidas_dir"],
                    es_carpeta=True,
                )
            )
            rutas["soc"] = None
    else:
        filas.append(
            _fila(
                "soc",
                "SoC del periodo",
                1,
                "falta",
                "ingresa el periodo (AAMM) arriba para poder buscarlo",
                ruta=rutas["medidas_dir"],
                es_carpeta=True,
            )
        )
        rutas["soc"] = None

    # ---- Auxiliares/ ----------------------------------------------
    agregar(
        "auxiliares_dir", f"{CARPETA_AUXILIARES}/", 0,
        rutas["auxiliares_dir"].is_dir(),
        ruta=rutas["auxiliares_dir"], es_carpeta=True,
    )
    agregar(
        "centrales", ARCHIVO_CENTRALES, 1, rutas["centrales"].is_file(),
        ruta=rutas["centrales"],
    )

    if rutas["centrales"].is_file():

        hojas = hojas_de(rutas["centrales"])

        if hojas is None:
            filas.append(
                _fila(
                    "centrales:hojas",
                    "hojas de Centrales.xlsx",
                    2,
                    "falta",
                    "no se pudo abrir el archivo",
                )
            )
        else:
            hojas_norm = {normalizar(hoja) for hoja in hojas}
            for hoja in (HOJA_RESUMEN_BESS, HOJA_DICCIONARIO):
                agregar(
                    f"centrales:{hoja}", f"hoja '{hoja}'", 2,
                    normalizar(hoja) in hojas_norm,
                )

    # Archivo de homologacion (punto de medida + canal -> clave), que
    # alimenta la descarga de Medidas_SAE.xlsx.
    archivo_homol = Homologacion.buscar_archivo_homologacion(
        rutas["auxiliares_dir"]
    )
    rutas["homologacion"] = archivo_homol

    if archivo_homol:
        filas.append(
            _fila(
                "homologacion", archivo_homol.name, 1, "ok",
                f"en {CARPETA_AUXILIARES}/", ruta=archivo_homol,
            )
        )

        hojas_homol = hojas_de(archivo_homol)
        hojas_homol_norm = (
            {normalizar(h) for h in hojas_homol} if hojas_homol else set()
        )

        agregar(
            f"homologacion:{Homologacion.HOJA_HOMOL}",
            f"hoja '{Homologacion.HOJA_HOMOL}'", 2,
            normalizar(Homologacion.HOJA_HOMOL) in hojas_homol_norm,
            "punto de medida + canal -> clave",
        )

        # "Gen real" es opcional: un caso donde ninguna central venga
        # de la API de operacion real es valido.
        tiene_gen_real = (
            normalizar(Homologacion.HOJA_GEN_REAL) in hojas_homol_norm
        )
        filas.append(
            _fila(
                f"homologacion:{Homologacion.HOJA_GEN_REAL}",
                f"hoja '{Homologacion.HOJA_GEN_REAL}'", 2,
                "ok" if tiene_gen_real else "pendiente",
                (
                    "centrales que se agregan desde la API de "
                    "operacion real"
                    if tiene_gen_real
                    else "opcional: sin ella no se agrega ninguna "
                         "central de operacion real"
                ),
            )
        )

    else:
        filas.append(
            _fila(
                "homologacion", "Archivo *Homologacion*", 1, "falta",
                f"ningun Excel de {CARPETA_AUXILIARES}/ tiene "
                f"'homologacion' en el nombre",
                ruta=rutas["auxiliares_dir"], es_carpeta=True,
            )
        )

    # ---- Ofertas/ -------------------------------------------------
    agregar(
        "ofertas_dir", f"{CARPETA_OFERTAS}/", 0,
        rutas["ofertas_dir"].is_dir(),
        ruta=rutas["ofertas_dir"], es_carpeta=True,
    )

    archivo_ofertas = buscar_archivo_ofertas(rutas["ofertas_dir"])
    rutas["ofertas"] = archivo_ofertas

    if archivo_ofertas:
        filas.append(
            _fila(
                "ofertas", archivo_ofertas.name, 1, "ok",
                f"en {CARPETA_OFERTAS}/", ruta=archivo_ofertas,
            )
        )
    else:
        filas.append(
            _fila(
                "ofertas", "Archivo *OfertasSSCC*", 1, "falta",
                f"ningun archivo en {CARPETA_OFERTAS}/ contiene "
                f"'OfertasSSCC' en el nombre",
                ruta=rutas["ofertas_dir"], es_carpeta=True,
            )
        )

    # ---- Cmg/ -----------------------------------------------------
    # Dos archivos, en orden de uso: primero se trae el CSV 15-minutal
    # de la unidad de red (boton "Traer"), y con ese CSV ya al
    # lado se genera cmg.xlsx ("Generar").
    agregar(
        "cmg_dir", f"{CARPETA_CMG}/", 0, rutas["cmg_dir"].is_dir(),
        ruta=rutas["cmg_dir"], es_carpeta=True,
    )

    rutas["cmg_csv"] = (
        extrae_cmg.ruta_csv_local(rutas["cmg_dir"], aamm_valido)
        if aamm_valido else None
    )

    if rutas["cmg_csv"] is None:
        filas.append(
            _fila(
                "cmg_csv", "cmg<AAMM>_def_15minutal.csv", 1, "falta",
                "ingresa el periodo (AAMM) arriba para poder traerlo",
                ruta=rutas["cmg_dir"], es_carpeta=True,
                origen="cmg_csv",
            )
        )
    else:
        filas.append(
            _fila(
                "cmg_csv",
                rutas["cmg_csv"].name,
                1,
                "ok" if rutas["cmg_csv"].is_file() else "falta",
                (
                    f"en {CARPETA_CMG}/"
                    if rutas["cmg_csv"].is_file()
                    else "se baja de la unidad de red con el boton ->"
                ),
                ruta=rutas["cmg_csv"],
                origen="cmg_csv",
            )
        )

    filas.append(
        _fila(
            "cmg_xlsx",
            ARCHIVO_CMG,
            1,
            "ok" if rutas["cmg"].is_file() else "pendiente",
            (
                "se regenera desde el CSV de arriba ->"
                if rutas["cmg"].is_file()
                else "se genera desde el CSV de arriba ->"
            ),
            ruta=rutas["cmg"],
            origen="cmg_xlsx",
        )
    )

    # ---- FD y FMA/ ------------------------------------------------
    # Se muestra el nombre de la carpeta que REALMENTE se esta usando
    # (puede ser la vieja si el caso todavia no se renombro).
    agregar(
        "sscc_dir", f"{rutas['sscc_desempeno_dir'].name}/", 0,
        rutas["sscc_desempeno_dir"].is_dir(),
        ruta=rutas["sscc_desempeno_dir"], es_carpeta=True,
    )

    archivo_sscc = buscar_archivo_sscc_desempeno(rutas["sscc_desempeno_dir"])
    rutas["sscc_desempeno"] = archivo_sscc

    if archivo_sscc:
        filas.append(
            _fila(
                "sscc", archivo_sscc.name, 1, "ok",
                f"en {CARPETA_FD_FMA}/", ruta=archivo_sscc,
                origen="fd",
            )
        )
    else:
        filas.append(
            _fila(
                "sscc", "Archivo SSCC_Desempeño_*", 1, "falta",
                f"no esta en {CARPETA_FD_FMA}/: se baja del DCO con el "
                f"boton 'Traer' de esta fila",
                ruta=rutas["sscc_desempeno_dir"], es_carpeta=True,
                origen="fd",
            )
        )

    # Las tres salidas de FMA viven en la misma carpeta (pedido del
    # usuario). No bloquean: si falta alguna, el FMA de ese tipo de
    # servicio queda en 0, igual que la formula original.
    archivos_fma = (
        fma_subastas.buscar_archivos_fma(
            rutas["sscc_desempeno_dir"], aamm_valido
        )
        if aamm_valido else {"cpf": None, "csf": None, "ctf": None}
    )
    rutas["fma"] = archivos_fma

    for tipo, etiqueta in (
        ("cpf", "fma_cpf"), ("csf", "fma_csf"), ("ctf", "fma_cft"),
    ):
        archivo = archivos_fma.get(tipo)

        if archivo:
            filas.append(
                _fila(
                    f"fma_{tipo}", archivo.name, 1, "ok",
                    f"alimenta Subastas!FMA ({tipo.upper()}); se rehace "
                    f"con el boton ->",
                    ruta=archivo,
                    origen=f"fma_{tipo}",
                )
            )
        elif not aamm_valido:
            filas.append(
                _fila(
                    f"fma_{tipo}", f"{etiqueta}<AAMM>.xlsx", 1, "pendiente",
                    "ingresa el periodo (AAMM) arriba para buscarlo",
                    ruta=rutas["sscc_desempeno_dir"], es_carpeta=True,
                    origen=f"fma_{tipo}",
                )
            )
        else:
            filas.append(
                _fila(
                    f"fma_{tipo}", f"{etiqueta}_{aamm_valido}.xlsx", 1,
                    "pendiente",
                    (
                        "se arma con el boton -> desde los reportes "
                        "diarios del DCO"
                        if tipo == "cpf" else
                        "se arma con el boton -> desde los reportes del "
                        "AGC (se copian a agcface/)"
                        if tipo == "csf" else
                        "se arma con el boton -> desde el CTF_AAMM.csv "
                        "del DCO"
                    ),
                    ruta=rutas["sscc_desempeno_dir"] / f"{etiqueta}_{aamm_valido}.xlsx",
                    origen=f"fma_{tipo}",
                )
            )

    # ---- Subastas/ ------------------------------------------------
    agregar(
        "subastas_dir", f"{CARPETA_SUBASTAS}/", 0,
        rutas["subastas_dir"].is_dir(),
        ruta=rutas["subastas_dir"], es_carpeta=True,
    )

    # Los Access son AHORA el origen de la hoja Subastas. La carpeta la
    # crea el programa (pedido del usuario), no la persona.
    if rutas["subastas_dir"].is_dir():
        ofertas_adj.asegurar_carpeta_db(rutas["subastas_dir"])

    accdb = (
        ofertas_adj.accdb_presentes(rutas["subastas_dir"], aamm_valido)
        if aamm_valido and rutas["db_subastas_dir"].is_dir()
        else []
    )
    rutas["accdb_subastas"] = [ruta for _, _, ruta in accdb]

    if not rutas["db_subastas_dir"].is_dir():
        detalle_db = "no se pudo crear; revisa permisos"
        estado_db = "falta"
    elif accdb:
        dias = len({dia for dia, _, _ in accdb})
        detalle_db = (
            f"{len(accdb)} Access del periodo, {dias} dia(s) - "
            f"se refrescan con el boton ->"
        )
        estado_db = "ok"
    elif not aamm_valido:
        detalle_db = "ingresa el periodo (AAMM) para revisar que hay"
        estado_db = "pendiente"
    else:
        detalle_db = (
            "vacia para el periodo: se traen de la unidad de red con "
            "el boton ->"
        )
        estado_db = "pendiente"

    filas.append(
        _fila(
            "db_subastas", f"{CARPETA_DB_SUBASTAS}/", 1, estado_db,
            detalle_db, ruta=rutas["db_subastas_dir"], es_carpeta=True,
            origen="subastas",
        )
    )

    # ---- Prorrata de retiros --------------------------------------
    agregar(
        "prorrata_dir", f"{CARPETA_PRORRATA_RETIROS}/", 0,
        rutas["prorrata_retiros_dir"].is_dir(),
        ruta=rutas["prorrata_retiros_dir"], es_carpeta=True,
    )
    try:
        archivo_prorrata = buscar_archivo_prorrata(
            rutas["prorrata_retiros_dir"], aamm_valido
        )
        rutas["prorrata_retiros"] = archivo_prorrata
        filas.append(_fila(
            "prorrata_archivo",
            archivo_prorrata.name if archivo_prorrata else "Prorrata_Retiros_AAMM_pre/def.xlsx",
            1, "ok" if archivo_prorrata else "falta",
            "fuente: hoja 'Prorrata 15min'" if archivo_prorrata else "deja aqui el Excel del periodo",
            ruta=archivo_prorrata or rutas["prorrata_retiros_dir"],
            es_carpeta=not archivo_prorrata,
        ))
    except ErrorEntrada as error:
        rutas["prorrata_retiros"] = None
        filas.append(_fila(
            "prorrata_archivo", "Prorrata_Retiros_AAMM_pre/def.xlsx", 1,
            "falta", str(error).split("\n")[0],
            ruta=rutas["prorrata_retiros_dir"], es_carpeta=True,
        ))

    # ---- Salida ---------------------------------------------------
    # UNA sola planilla (el usuario pidio combinar las dos que habia),
    # desglosada igual que Centrales.xlsx: el archivo y, adentro, una
    # fila por hoja, en el mismo orden en que quedan las hojas del
    # libro -de fin a inicio: el Resumen arriba y las entradas abajo-.
    # Cada hoja se actualiza por separado desde su propio boton; si el
    # archivo todavia no existe, se crea al actualizar la primera hoja.
    filas_hojas = _filas_de_hojas(rutas["salida"], SECCIONES_SALIDA, 1)

    # El archivo esta "ok" solo si TODAS sus hojas tienen datos: que el
    # .xlsx exista no dice nada (se crea entero, con las hojas que
    # todavia no se generaron vacias).
    completas = all(fila["estado"] == "ok" for fila in filas_hojas)

    if not rutas["salida"].is_file():
        detalle_salida = "salida: se crea al actualizar la primera hoja ->"
    elif completas:
        detalle_salida = "salida: se actualiza hoja por hoja ->"
    else:
        detalle_salida = "salida: le faltan hojas por generar ->"

    filas.append(
        _fila(
            "salida", ARCHIVO_SALIDA, 0,
            "ok" if (rutas["salida"].is_file() and completas) else "pendiente",
            detalle_salida,
            ruta=rutas["salida"],
        )
    )
    filas.extend(filas_hojas)

    # El control de la corrida (Alertas, Ejecucion, Log) vive en su
    # propio archivo, al lado de la planilla: no se genera solo, lo
    # deja cada corrida.
    filas.append(
        _fila(
            "control", ARCHIVO_CONTROL, 0,
            "ok" if rutas["control"].is_file() else "pendiente",
            "control de la corrida: Alertas, Ejecucion y Log",
            ruta=rutas["control"],
        )
    )

    return rutas, filas


# ============================================================
# PROCESO COMPLETO
#
# UNA salida (ARCHIVO_SALIDA), escrita por dos mitades independientes:
#
#   - generar_consolidado(): las hojas de ENTRADA (Medidores, Ofertas
#     SSCC, CMg, FD, Subastas).
#   - generar_pagos_bess(): las hojas de CALCULO -- lee Medidores y
#     Subastas ya escritas en la misma planilla, no las recalcula.
#
# Las dos reciben un set de "secciones activas": lo que entra se
# recalcula y lo que queda afuera se preserva tal cual estaba en el
# archivo (ver escribir_salida/hojas_regenerar). En la ventana, cada
# seccion es una fila-hoja del diagrama con su boton "Actualizar", y
# el boton del archivo manda todas juntas.
#
# SECCIONES_CONSOLIDADO agrupa cada seccion con las hojas que produce.
# "medidores" junta Medidas_SAE, SoC, Centrales (Diccionario) y
# OfertasSSCC porque construir_medidores() necesita los 4 juntos: no
# se pueden actualizar por separado a ese nivel de detalle sin
# recalcular con datos parcialmente viejos.
#
# El tercer elemento de cada tupla (la descripcion de que lee esa
# seccion) ya no se muestra en la ventana -antes era el texto debajo
# de cada casilla-: queda como documentacion del contrato de cada
# seccion, que es donde hay que mirarlo al tocar una.
# ============================================================

SECCIONES_CONSOLIDADO = (
    (
        "medidores",
        "Medidores",
        f"Usa {ARCHIVO_MEDIDAS_SAE}, el SoC del periodo y "
        f"{ARCHIVO_CENTRALES}. Ya NO usa OfertasSSCC: las columnas "
        f"que salian de ahi viven en la hoja 'Ofertas SSCC'.",
        (HOJA_MEDIDORES,),
    ),
    (
        "ofertas_sscc",
        "Ofertas SSCC",
        f"Usa el archivo *OfertasSSCC* de {CARPETA_OFERTAS}/ y la hoja "
        f"'Medidores' (centrales y ventanas). Guarda las dos tablas de "
        f"las que salen las columnas R, S y T de la planilla original: "
        f"'Medidores' ya no las trae ni depende de este archivo.",
        (HOJA_OFERTAS_SSCC,),
    ),
    (
        "cmg",
        "CMg",
        f"Usa {ARCHIVO_CMG}.",
        (HOJA_CMG,),
    ),
    (
        "fd",
        "FD",
        f"Usa el archivo {CARPETA_FD_FMA}/ (hojas CPF/CSF "
        f"Horario).",
        (HOJA_FD,),
    ),
    (
        "subastas",
        "Subastas",
        f"Usa los Access de {CARPETA_SUBASTAS}/{CARPETA_DB_SUBASTAS}/ "
        f"(su origen real), y de {CARPETA_FD_FMA}/ las salidas de FMA "
        f"(fma_cpf/fma_csf/fma_cft) y el SSCC_Desempeño_* (columna FD y "
        f"Vector de Participacion CSF), mas {ARCHIVO_CENTRALES} "
        f"(Propietario + nomenclaturas).",
        (HOJA_SUBASTAS,),
    ),
)


SECCIONES_PAGOS = (
    (
        "ecostos",
        "Calculo E Costos",
        f"Usa las hojas 'Medidores', 'Ofertas SSCC', 'FD' y 'Subastas' "
        f"de esta misma planilla, mas {ARCHIVO_CENTRALES} y "
        f"{ARCHIVO_CMG}. El FD homologado de AM:AR sale de la hoja "
        f"'FD' del consolidado, no de releer el SSCC_Desempeño_*.",
        (HOJA_CALCULO_ECOSTOS,),
    ),
    (
        "re545",
        "Calculo RE545",
        f"Usa las hojas 'Medidores', 'Ofertas SSCC' y 'Subastas' de "
        f"esta misma planilla, mas {ARCHIVO_CENTRALES} y {ARCHIVO_CMG} -- "
        f"no necesita el FD.",
        (HOJA_CALCULO_RE545,),
    ),
    (
        "compensacion_central",
        HOJA_COMPENSACION_CENTRAL,
        "Resume la compensacion de las dos hojas de calculo por central y "
        "ciclo/ventana, y el total que recibe cada empresa (el Propietario "
        f"de '{HOJA_RESUMEN_BESS}' de {ARCHIVO_CENTRALES}).",
        (HOJA_COMPENSACION_CENTRAL,),
    ),
    (
        "prorrata_retiros",
        HOJA_PRORRATA_RETIROS,
        f"Usa la hoja 'Prorrata 15min' del Excel de {CARPETA_PRORRATA_RETIROS}/ y las dos hojas de calculo.",
        (HOJA_PRORRATA_RETIROS,),
    ),
    (
        "resumen",
        HOJA_RESUMEN,
        "Consolida por empresa cuanto RECIBE, PAGA y su NETO.",
        (HOJA_RESUMEN,),
    ),
)


# Las hojas de la planilla, en el orden en que quedan en el libro (de
# FIN A INICIO: el Resumen primero, las entradas al final) y con el
# prefijo de id que dice quien las genera: "consolidado" para las
# entradas, "pagos" para el calculo. Es lo que recorre el diagrama de
# la ventana, y sale de dar vuelta las dos tuplas de arriba -- que
# estan escritas en el orden en que se CALCULAN, que es justo el
# inverso del orden en que se LEEN.
SECCIONES_SALIDA = (
    tuple(("pagos", seccion) for seccion in reversed(SECCIONES_PAGOS))
    + tuple(
        ("consolidado", seccion)
        for seccion in reversed(SECCIONES_CONSOLIDADO)
    )
)

# Control de coherencia: el orden de las filas del diagrama y el de
# las hojas del libro tienen que ser el mismo. Si alguien agrega una
# hoja en un lado y se olvida del otro, revienta al importar, no en la
# ventana del usuario.
_HOJAS_DE_LAS_SECCIONES = tuple(
    hoja
    for _, (_, _, _, hojas) in SECCIONES_SALIDA
    for hoja in hojas
)

assert _HOJAS_DE_LAS_SECCIONES == ORDEN_HOJAS_SALIDA, (
    "El orden de las hojas del diagrama no coincide con "
    "ORDEN_HOJAS_SALIDA: " + str(_HOJAS_DE_LAS_SECCIONES)
)
