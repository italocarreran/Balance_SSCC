# -*- coding: utf-8 -*-
"""
Medidas_CRA.xlsx: la energia del CRA armada desde PRMTE, igual que el
Medidas_SAE.xlsx del BESS.

Pedido del usuario (2026-11-26): "el archivo Formato_Solicitud_SSAA tenga
un boton para generarlo, la info viene de PRMTE asi que tenemos que
replicar el archivo de los SAE, que genere una salida que se llame
Medidas_CRA [...] En auxiliares del CRA va a haber un homologacion tambien
igual que el de SAE", y despues: "que la salida de CRA tenga el mismo
formato que [...] los SAE".

Es EXACTAMENTE el proceso del SAE (nucleo.medidas_sae.armar_medidas), con
la homologacion de <mes>/CRA/Auxiliares/ y su propia carpeta de trabajo:

  1. la hoja 'homol' (Punto de Medida + Canal -> clave, con el Flujo;
     varios paños en una linea con '+', ver Homologacion.expandir_punto);
  2. la descarga de PRMTE por punto de medida, por lotes y reanudable,
     en <mes>/CRA/Energia/_trabajo/;
  3. calendario de cuartos de hora y suma por clave con el signo del Flujo;
  4. las centrales de la hoja 'Gen real' (API de operacion real), si hay.

Salida con las mismas columnas que Medidas_SAE.xlsx (COLUMNAS_AI) y
Puntos_fallidos.xlsx al lado. La hoja ENERGIA sigue leyendo el
Formato_Solicitud.
"""

from ..nucleo.escritura import escribir_tabla_rapida
from ..nucleo.externos import Homologacion
from ..nucleo.medidas_sae import (
    armar_medidas, escribir_puntos_fallidos, informar_escritura,
)
from ..nucleo.utiles import ErrorEntrada
from . import parametros as p
from .rutas import resolver_rutas, validar_aamm


def _nada(*_args, **_kwargs):
    pass


def generar_medidas_cra(carpeta_base, aamm, registrar=_nada, progreso=_nada):
    """
    Boton "Generar" de la fila del Formato_Solicitud_SSAA_*: arma
    <mes>/CRA/Energia/Medidas_CRA.xlsx desde PRMTE y, al lado,
    Puntos_fallidos.xlsx (vacio si no falto nada). Devuelve la ruta.
    """

    validar_aamm(aamm)
    rutas = resolver_rutas(carpeta_base, aamm)

    archivo_homol = Homologacion.buscar_archivo_homologacion(rutas["auxiliares"])
    if not archivo_homol:
        raise ErrorEntrada(
            f"No se encontro el archivo de homologacion en {rutas['auxiliares']}: "
            f"ningun Excel de esa carpeta tiene 'homologacion' en el nombre "
            f"(el mismo formato que el del BESS: hoja 'homol' con Punto de "
            f"Medida, Canal, clave y Flujo)."
        )

    df, df_fallidos = armar_medidas(
        archivo_homol, aamm, rutas["trabajo_medidas"],
        registrar=registrar, progreso=progreso, origen="Medidas_CRA",
    )

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
    informar_escritura(
        df, df_fallidos, rutas["medidas_cra"], rutas["puntos_fallidos"],
        registrar,
    )

    progreso(100)
    registrar(f"Listo: {rutas['medidas_cra']}")
    return rutas["medidas_cra"]
