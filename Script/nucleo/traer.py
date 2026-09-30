# -*- coding: utf-8 -*-
"""
Botones 'Traer'/'Generar': lo que el programa baja o arma.
"""

from pathlib import Path

from .externos import (
    extrae_cmg, indicadores_dco, indices_fma, ofertas_adj,
)
from .diccionarios import construir_mapa_barra
from .lectura import leer_centrales
from .parametros import (
    ARCHIVO_CENTRALES, CARPETA_FD_FMA, HOJA_CMG_ORIGEN,
    HOJA_RESUMEN_BESS,
)
from .rutas import resolver_rutas, validar_aamm
from .utiles import ErrorEntrada


# ============================================================
# cmg.xlsx: TRAER EL CSV Y GENERARLO
#
# La logica en si vive en Script/Cmg/Extrae_CMG_barras.py (que no
# conoce la estructura del caso: recibe rutas y barras). Aca queda
# solo lo que SI es del caso: resolver las rutas, sacar las barras de
# Centrales.xlsx y escribir el Excel.
# ============================================================

def traer_csv_cmg(carpeta_base, aamm, registrar=print, progreso=None):
    """
    Copia el CSV 15-minutal del periodo desde la unidad de red a
    <CARPETA_BASE>/Cmg/ (boton "Traer cmg_15min"). Devuelve la ruta
    local del CSV.
    """

    aamm = validar_aamm(aamm)
    rutas = resolver_rutas(carpeta_base)

    if not rutas["base"].is_dir():
        raise ErrorEntrada(f"No se encontro la carpeta base {rutas['base']}")

    if progreso:
        progreso(10)

    try:
        destino = extrae_cmg.traer_csv_15min(
            rutas["cmg_dir"], aamm, registrar=registrar
        )
    except extrae_cmg.ErrorCmg as error:
        raise ErrorEntrada(str(error)) from error
    except OSError as error:
        raise ErrorEntrada(
            f"No se pudo copiar el CSV de CMg: {error}"
        ) from error

    if progreso:
        progreso(100)

    registrar(f"Listo: {destino}")

    return destino


def traer_subastas(carpeta_base, aamm, registrar=print, progreso=None):
    """
    Copia los Access de subastas del periodo (OfertasSSCCAdj*.accdb)
    desde la unidad de red a <CARPETA_BASE>/Subastas/DB subastas/
    (boton "Traer subastas"). Crea la carpeta si no existe. Devuelve
    la ruta de esa carpeta.
    """

    aamm = validar_aamm(aamm)
    rutas = resolver_rutas(carpeta_base)

    if not rutas["base"].is_dir():
        raise ErrorEntrada(f"No se encontro la carpeta base {rutas['base']}")

    if not rutas["subastas_dir"].is_dir():
        raise ErrorEntrada(
            f"No se encontro la carpeta {rutas['subastas_dir']}"
        )

    if progreso:
        progreso(5)

    registrar(
        f"Trayendo subastas de {ofertas_adj.ruta_origen()} "
        f"(periodo {aamm})..."
    )

    try:
        copiados, salteados, _ = ofertas_adj.traer_accdb(
            rutas["subastas_dir"], aamm, registrar=registrar
        )
    except ofertas_adj.ErrorSubastas as error:
        raise ErrorEntrada(str(error)) from error
    except OSError as error:
        raise ErrorEntrada(
            f"No se pudieron copiar los Access de subastas: {error}"
        ) from error

    if progreso:
        progreso(100)

    registrar(
        f"Listo: {len(copiados)} copiado(s) y {len(salteados)} ya al dia "
        f"en {rutas['db_subastas_dir']}"
    )

    return rutas["db_subastas_dir"]


def traer_fd(carpeta_base, aamm, version=None, registrar=print, progreso=None):
    """
    Copia el FD del periodo (lo que el DCO publica como
    SSCC_Disponibilidad_CSF_* / SSCC_Desempeño_*) desde el arbol de
    indicadores del DCO a <CARPETA_BASE>/FD y FMA/, y descomprime el
    zip si lo que vino es un zip (boton "Traer FD"). Devuelve la ruta
    de esa carpeta.
    """

    aamm = validar_aamm(aamm)
    rutas = resolver_rutas(carpeta_base)

    if not rutas["base"].is_dir():
        raise ErrorEntrada(f"No se encontro la carpeta base {rutas['base']}")

    destino = rutas["sscc_desempeno_dir"]

    if not destino.is_dir():
        raise ErrorEntrada(
            f"No se encontro la carpeta {CARPETA_FD_FMA}/ en "
            f"{rutas['base']}"
        )

    if progreso:
        progreso(5)

    registrar(
        f"Trayendo el FD de {indicadores_dco.RAIZ_DCO_INDICADORES} "
        f"(periodo {aamm})..."
    )

    try:
        copiados, extraidos, carpeta_version = indicadores_dco.traer_fd(
            destino, aamm, version=version, registrar=registrar
        )
    except indicadores_dco.ErrorFd as error:
        raise ErrorEntrada(str(error)) from error
    except OSError as error:
        raise ErrorEntrada(f"No se pudo traer el FD: {error}") from error

    if progreso:
        progreso(100)

    registrar(
        f"Listo: {len(copiados)} archivo(s) desde {carpeta_version.name}"
        + (f" y {len(extraidos)} descomprimido(s)" if extraidos else "")
        + f" en {destino}"
    )

    return destino


def generar_fma(
    carpeta_base, aamm, tipos=None, version=None, registrar=print,
    progreso=None
):
    """
    Arma la(s) salida(s) de FMA pedidas y las deja en
    <CARPETA_BASE>/FD y FMA/, que es de donde las lee despues la hoja
    Subastas. Cada una tiene su propio boton "Generar" en la ventana,
    asi que lo normal es que venga un solo tipo.

    tipos: subconjunto de indices_fma.TIPOS_FMA ("cpf", "csf", "ctf");
    None = las tres.

    OJO: el FMA no se copia ya hecho, se CONSTRUYE -- el CPF desde los
    reportes diarios del DCO, el CSF desde los reportes del AGC (que se
    copian antes a 'FD y FMA/agcface/') y el CTF desde el CTF_AAMM.csv
    del propio DCO. Ver Script/Fd/Indices_FMA.py.
    """

    aamm = validar_aamm(aamm)
    rutas = resolver_rutas(carpeta_base)

    if not rutas["base"].is_dir():
        raise ErrorEntrada(f"No se encontro la carpeta base {rutas['base']}")

    destino = rutas["sscc_desempeno_dir"]

    if not destino.is_dir():
        raise ErrorEntrada(
            f"No se encontro la carpeta {CARPETA_FD_FMA}/ en "
            f"{rutas['base']}"
        )

    if progreso:
        progreso(5)

    etiquetas = ", ".join(sorted(t.upper() for t in (tipos or indices_fma.TIPOS_FMA)))
    registrar(f"Generando FMA {etiquetas} del periodo {aamm}...")

    try:
        escritos, faltantes, versiones = indices_fma.generar_fma(
            destino, aamm, tipos=tipos, version=version, registrar=registrar
        )
    except (indices_fma.ErrorIndicesFma, indicadores_dco.ErrorFd) as error:
        raise ErrorEntrada(str(error)) from error
    except OSError as error:
        raise ErrorEntrada(f"No se pudo generar el FMA: {error}") from error

    for faltante in faltantes:
        registrar(f"  [AVISO] no se pudo armar el FMA de {faltante}")

    if progreso:
        progreso(100)

    detalle_version = ", ".join(
        f"{tipo.upper()} desde {nombre}" for tipo, nombre in sorted(versiones.items())
    )

    registrar(
        f"Listo: {', '.join(escritos.values())} en {destino}"
        + (f" ({detalle_version})" if detalle_version else "")
    )

    return destino


def generar_cmg(
    carpeta_base, aamm, ruta_csv=None, registrar=print, progreso=None
):
    """
    Genera/actualiza <CARPETA_BASE>/Cmg/cmg.xlsx a partir del CSV
    15-minutal que ya esta en esa misma carpeta (se trae con
    traer_csv_cmg / boton "Traer cmg_15min").

    Las barras a filtrar salen de "Resumen BESS" de Centrales.xlsx,
    via construir_mapa_barra(): la MISMA fuente que alimenta
    Calculo E Costos!Barra, asi que las dos puntas no se pueden
    desincronizar.

    ruta_csv: opcional, para forzar otro CSV.
    """

    def avanzar(valor):
        if progreso:
            progreso(valor)

    aamm = validar_aamm(aamm)
    rutas = resolver_rutas(carpeta_base)

    if not rutas["base"].is_dir():
        raise ErrorEntrada(f"No se encontro la carpeta base {rutas['base']}")

    if not rutas["centrales"].is_file():
        raise ErrorEntrada(
            f"No se encontro {rutas['centrales']} (de ahi salen las "
            f"barras a filtrar)."
        )

    ruta_csv = (
        Path(ruta_csv) if ruta_csv
        else extrae_cmg.ruta_csv_local(rutas["cmg_dir"], aamm)
    )

    if not ruta_csv.is_file():
        raise ErrorEntrada(
            f"No esta el CSV 15-minutal del periodo {aamm} en la "
            f"carpeta del caso:\n{ruta_csv}\n\n"
            f"Usa primero el boton 'Traer cmg_15min' de esa fila."
        )

    avanzar(10)

    registrar(f"Leyendo {ARCHIVO_CENTRALES} (hoja '{HOJA_RESUMEN_BESS}')...")
    resumen, _ = leer_centrales(rutas["centrales"])
    barras = barras_desde_resumen_bess(resumen)
    registrar(f"  barras a filtrar: {len(barras)}")
    for barra in barras:
        registrar(f"    {barra}")

    avanzar(25)

    registrar(f"Leyendo {ruta_csv.name}...")

    try:
        df_salida, resumen_dias = extrae_cmg.construir_cmg_desde_csv(
            ruta_csv, barras, registrar=registrar
        )
        avanzar(70)
        extrae_cmg.validar_layout(df_salida, registrar=registrar)
    except extrae_cmg.ErrorCmg as error:
        raise ErrorEntrada(str(error)) from error

    rutas["cmg_dir"].mkdir(parents=True, exist_ok=True)

    registrar(f"Escribiendo {rutas['cmg']}...")
    df_salida.to_excel(rutas["cmg"], sheet_name=HOJA_CMG_ORIGEN, index=False)

    avanzar(95)

    registrar(f"  registros exportados: {len(df_salida):,}")
    registrar(
        f"  maximo Cuarto de Hora: "
        f"{int(df_salida[extrae_cmg.COLUMNA_CUARTO].max())}"
    )
    registrar(f"  dias en el archivo: {len(resumen_dias)}")

    anomalos = extrae_cmg.resumen_dias_anomalos(resumen_dias)

    if anomalos:
        registrar("  dias que NO tienen 24 horas (cambio de hora):")
        for linea in anomalos:
            registrar(f"    {linea}")

    avanzar(100)
    registrar(f"Listo: {rutas['cmg']}")

    return rutas["cmg"]


def barras_desde_resumen_bess(resumen_bess):
    """
    Lista de barras de inyeccion (sin repetir, en el orden en que
    aparecen) de la hoja "Resumen BESS" de Centrales.xlsx. Reusa
    construir_mapa_barra() -- misma deteccion de columna por nombre
    normalizado, mismo .strip() -- para que el filtro de cmg.xlsx y la
    homologacion de Calculo E Costos!Barra miren exactamente el mismo
    dato.
    """

    barras = []
    vistas = set()

    for barra in construir_mapa_barra(resumen_bess).values():

        if not barra:
            continue

        clave = barra.upper()

        if clave in vistas:
            continue

        vistas.add(clave)
        barras.append(barra)

    if not barras:
        raise ErrorEntrada(
            f"La columna 'Barra inyección' de la hoja "
            f"'{HOJA_RESUMEN_BESS}' de {ARCHIVO_CENTRALES} no tiene "
            f"ninguna barra cargada: sin barras no se puede filtrar el "
            f"CSV de CMg."
        )

    return barras
