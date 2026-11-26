# -*- coding: utf-8 -*-
"""
Medidas_CRA.xlsx: la energia del CRA armada desde PRMTE, igual que el
Medidas_SAE.xlsx del BESS.

Pedido del usuario (2026-11-26): "el archivo Formato_Solicitud_SSAA tenga
un boton para generarlo, la info viene de PRMTE asi que tenemos que
replicar el archivo de los SAE, que genere una salida que se llame
Medidas_CRA. Misma idea, y la salida en este formato por ahora, para
comparar: Unidad, Punto De Medida, Año, Mes, DIA, FECHA_HORA, kWhD, kWhR.
En auxiliares del CRA va a haber un homologacion tambien igual que el de
SAE".

Mismos pasos que nucleo.generar_medidas_sae, con las mismas piezas de
Script/Medidas:

  1. lee el Excel de homologacion de <mes>/CRA/Auxiliares/ (hoja 'homol':
     Punto de Medida + Canal -> clave, que aca es la Unidad);
  2. descarga las medidas por punto de medida (API de medidas, por
     lotes y reanudable, en <mes>/CRA/Energia/_trabajo/);
  3. limpia la descarga igual que el SAE (fechas, hora inexistente del
     cambio de hora, huecos) y cruza con la homologacion.

Lo que cambia es la salida: en vez de sumar por clave con el signo del
Flujo, deja UNA FILA POR UNIDAD + PUNTO DE MEDIDA + CUARTO DE HORA, con
la energia entregada (kWhD) y la recibida (kWhR) en dos columnas, segun
el Canal de la homologacion (kWhD / kWhR), tal como las trae la API
(sin Flujo): es la forma del Formato_Solicitud_SSAA_SSCC_Hidro_*, para
compararlos. La hoja ENERGIA sigue leyendo el Formato_Solicitud.

La hoja "Gen real" de la homologacion (centrales desde la API de
operacion real) todavia no se usa aca: si tiene filas, se avisa.
"""

import pandas as pd

from ..nucleo.escritura import escribir_tabla_rapida
from ..nucleo.externos import (
    Claves_Balance, Descarga_PRMTE, ErrorMedidas, Homologacion,
)
from ..nucleo.medidas_sae import armar_puntos_fallidos, escribir_puntos_fallidos
from ..nucleo.utiles import ErrorEntrada, normalizar
from . import parametros as p
from .rutas import resolver_rutas, validar_aamm


def _nada(*_args, **_kwargs):
    pass


def tipo_de_canal(canal):
    """'kWhD' -> 'kWhD', 'kWhR' -> 'kWhR', otro -> None (no se sabe)."""

    texto = normalizar(canal)
    if texto.endswith("d"):
        return p.CAMPO_MEDIDAS_KWHD
    if texto.endswith("r"):
        return p.CAMPO_MEDIDAS_KWHR
    return None


def construir_medidas_cra(df_crudo, df_homol, registrar=_nada):
    """
    De la descarga de PRMTE a (df_medidas_cra, calendario, huecos).

    df_medidas_cra trae p.COLUMNAS_MEDIDAS_CRA, ordenado por Unidad,
    Punto De Medida y FECHA_HORA.
    """

    cb = Claves_Balance
    df = cb.normalizar_fechas(df_crudo, registrar=registrar)
    df, descartadas = cb.descartar_hora_inexistente(df)
    if not descartadas.empty:
        registrar(
            f"  filas descartadas por la hora inexistente del cambio de hora "
            f"o sin intervaloUtc: {len(descartadas):,}"
        )

    esperados = cb.cuartos_de_hora_del_mes(df)
    registrar(f"  cuartos de hora del mes: {esperados:,}")
    _, incompletos = cb.detectar_incompletos(df, esperados)
    huecos = cb.detallar_huecos(df)
    if not incompletos.empty:
        registrar(
            f"  puntos de medida incompletos: {len(incompletos):,} (se "
            f"incluyen con lo que traen; detalle en Puntos_fallidos.xlsx)"
        )

    df = df[df["principal"] == True].copy()  # noqa: E712
    if df.empty:
        raise ErrorMedidas(
            "La descarga no trae ningun registro principal: no hay con que "
            "armar Medidas_CRA.xlsx."
        )
    df["canalVal"] = df["canalVal1"].fillna(0) + df["canalVal3"].fillna(0)

    cruzado = df.merge(
        df_homol,
        left_on=["idPuntoMedida", "slugCanal"],
        right_on=["Punto de Medida", "Canal"],
        how="inner",
    )
    if cruzado.empty:
        raise ErrorMedidas(
            "Ningun punto de medida descargado cruza con la homologacion "
            "del CRA (idPuntoMedida + slugCanal contra 'Punto de Medida' + "
            "'Canal'). Revisa el archivo de CRA/Auxiliares/."
        )

    cruzado["_tipo"] = cruzado["Canal"].map(tipo_de_canal)
    sin_tipo = cruzado[cruzado["_tipo"].isna()]
    if not sin_tipo.empty:
        canales = sorted(set(sin_tipo["Canal"].astype(str)))
        registrar(
            f"  AVISO: canal(es) {', '.join(canales)} de la homologacion no "
            f"terminan en D ni en R: no se sabe si son kWhD o kWhR y quedan "
            f"fuera ({len(sin_tipo):,} fila(s))."
        )
        cruzado = cruzado[cruzado["_tipo"].notna()]

    registrar(
        f"  registros homologados: {len(cruzado):,} "
        f"({cruzado['clave'].nunique()} unidad(es))"
    )

    calendario = cb.construir_calendario(cruzado)
    cb.avisar_hora_repetida(calendario, registrar)

    local = cb.COL_INTERVALO_LOCAL
    tabla = (
        cruzado
        .pivot_table(
            index=["clave", "Punto de Medida", local, cb.COL_INTERVALO_UTC],
            columns="_tipo", values="canalVal", aggfunc="sum",
        )
        .reindex(columns=[p.CAMPO_MEDIDAS_KWHD, p.CAMPO_MEDIDAS_KWHR])
        .fillna(0.0)
        .reset_index()
        .sort_values(["clave", "Punto de Medida", cb.COL_INTERVALO_UTC])
    )
    tabla.columns.name = None

    fechas = pd.to_datetime(tabla[local])
    salida = pd.DataFrame({
        p.CAMPO_MEDIDAS_UNIDAD: tabla["clave"].astype(str),
        p.CAMPO_MEDIDAS_PUNTO: tabla["Punto de Medida"],
        p.CAMPO_MEDIDAS_ANIO: fechas.dt.year,
        p.CAMPO_MEDIDAS_MES: fechas.dt.month,
        p.CAMPO_MEDIDAS_DIA: fechas.dt.day,
        p.CAMPO_MEDIDAS_FECHA_HORA: fechas,
        p.CAMPO_MEDIDAS_KWHD: tabla[p.CAMPO_MEDIDAS_KWHD],
        p.CAMPO_MEDIDAS_KWHR: tabla[p.CAMPO_MEDIDAS_KWHR],
    })

    return salida.reset_index(drop=True), calendario, huecos


def generar_medidas_cra(carpeta_base, aamm, registrar=_nada, progreso=_nada):
    """
    Boton "Generar" de la fila del Formato_Solicitud_SSAA_*: arma
    <mes>/CRA/Energia/Medidas_CRA.xlsx desde PRMTE y, al lado,
    Puntos_fallidos.xlsx (vacio si no falto nada). Devuelve la ruta.
    """

    anio, mes = validar_aamm(aamm)
    periodo = f"{anio}{mes:02d}"
    rutas = resolver_rutas(carpeta_base, aamm)

    archivo_homol = Homologacion.buscar_archivo_homologacion(rutas["auxiliares"])
    if not archivo_homol:
        raise ErrorEntrada(
            f"No se encontro el archivo de homologacion en {rutas['auxiliares']}: "
            f"ningun Excel de esa carpeta tiene 'homologacion' en el nombre "
            f"(el mismo formato que el del BESS: hoja 'homol' con Punto de "
            f"Medida, Canal, clave y Flujo; la clave es la Unidad)."
        )

    registrar(f"Periodo: {anio}-{mes:02d} ({aamm})")
    try:
        registrar(f"Leyendo {archivo_homol.name}...")
        df_homol = Homologacion.leer_homologacion(archivo_homol)
        puntos = Homologacion.puntos_de_medida(df_homol)
        registrar(
            f"  puntos de medida a consultar: {len(puntos):,} "
            f"({df_homol['clave'].nunique()} unidad(es))"
        )
        if Homologacion.leer_gen_real(archivo_homol):
            registrar(
                f"  AVISO: la hoja '{Homologacion.HOJA_GEN_REAL}' tiene "
                f"centrales, pero Medidas_CRA todavia no las agrega (solo "
                f"PRMTE por punto de medida)."
            )
        progreso(5)

        registrar("Descargando medidas por punto de medida...")
        df_crudo, _ = Descarga_PRMTE.descargar(
            puntos, periodo, rutas["trabajo_medidas"],
            registrar=registrar, progreso=progreso, desde=5, hasta=80,
        )

        registrar("Separando kWhD y kWhR por unidad y punto de medida...")
        df, calendario, huecos = construir_medidas_cra(
            df_crudo, df_homol, registrar=registrar
        )
        df_fallidos = armar_puntos_fallidos(
            puntos, df_homol, df_crudo, huecos, calendario
        )
    except ErrorMedidas as error:
        raise ErrorEntrada(str(error)) from error
    progreso(85)

    rutas["energia"].mkdir(parents=True, exist_ok=True)
    registrar(f"Escribiendo {rutas['medidas_cra']}...")
    try:
        escribir_tabla_rapida(df, rutas["medidas_cra"], p.HOJA_MEDIDAS_CRA)
        escribir_puntos_fallidos(rutas["puntos_fallidos"], df_fallidos)
    except PermissionError as error:
        raise ErrorEntrada(
            f"No se pudo guardar en {rutas['energia']}: hay un archivo abierto "
            f"en Excel. Cerralo y volve a intentar."
        ) from error

    registrar(f"  filas: {len(df):,}")
    registrar(f"  unidades: {df[p.CAMPO_MEDIDAS_UNIDAD].nunique()}")
    registrar(
        f"  kWhD total: {df[p.CAMPO_MEDIDAS_KWHD].sum():,.3f}   "
        f"kWhR total: {df[p.CAMPO_MEDIDAS_KWHR].sum():,.3f}"
    )
    if df_fallidos.empty:
        registrar("Puntos_fallidos.xlsx: sin problemas (vacio).")
    else:
        registrar(
            f"AVISO: Medidas_CRA.xlsx se genero con informacion incompleta. "
            f"Detalle en {rutas['puntos_fallidos']}."
        )

    progreso(100)
    registrar(f"Listo: {rutas['medidas_cra']}")
    return rutas["medidas_cra"]
