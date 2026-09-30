# -*- coding: utf-8 -*-
"""
Medidas_SAE.xlsx: los cuatro pasos de un viaje.
"""

import calendar
import pandas as pd

from .externos import (
    Claves_Balance, Descarga_PRMTE, ErrorMedidas, Generacion_Real,
    Homologacion,
)
from .parametros import (
    ARCHIVO_MEDIDAS_SAE, COLUMNAS_AI, HOJA_MEDIDAS_SAE,
)
from .rutas import periodo_desde_aamm, resolver_rutas, validar_aamm
from .utiles import ErrorEntrada


# ============================================================
# Medidas_SAE.xlsx: LOS CUATRO PASOS DE UN VIAJE
#
# La logica vive en Script/Medidas/ (un modulo por cada uno de los
# scripts sueltos que habia antes), incluida la clave de las dos APIs
# (Script/config.py, seccion "claves_api" de config.json). Aca queda
# lo que es del caso:
# resolver rutas, pegar las dos fuentes -las dos hojas del Excel de
# homologacion- y escribir el Excel.
#
# Los intermedios (lotes descargados, marca de reanudacion) van a
# <CARPETA_BASE>/Medidas/_trabajo/, que la ventana no muestra: no son
# entradas ni salidas del caso. Los diagnosticos que el script
# original exportaba a Excel (conteos por punto de medida,
# incompletos, generacion total por clave) se resumen en el log.
# ============================================================

def _resumir_diagnostico_medidas(diagnostico, registrar):
    """Lo que antes iba a 'reporte_medidas_consolidadas.xlsx'."""

    incompletos = diagnostico["incompletos"]

    if not incompletos.empty:
        registrar(
            f"  puntos de medida descartados por incompletos "
            f"({len(incompletos):,}), esperados "
            f"{diagnostico['cuartos_esperados']:,} cuartos de hora:"
        )
        for _, fila in incompletos.head(15).iterrows():
            registrar(
                f"    {fila['idPuntoMedida']}: "
                f"canalVal1={int(fila['canalVal1']):,} "
                f"canalVal3={int(fila['canalVal3']):,}"
            )
        if len(incompletos) > 15:
            registrar(f"    ... y {len(incompletos) - 15:,} mas")

    total = diagnostico["generacion_total"]

    registrar(f"  generacion total por clave ({len(total)} clave(s)):")
    for _, fila in total.head(25).iterrows():
        registrar(f"    {fila['clave']}: {fila['Gen_Unidad']:,.3f}")
    if len(total) > 25:
        registrar(f"    ... y {len(total) - 25:,} mas")


def generar_medidas_sae(
    carpeta_base, aamm, registrar=print, progreso=None
):
    """
    Genera/actualiza <CARPETA_BASE>/Medidas/Medidas_SAE.xlsx corriendo
    los cuatro pasos seguidos (boton "Actualizar" de esa fila):

      1. lee el Excel de homologacion de Auxiliares/;
      2. descarga las medidas por punto de medida (API de medidas),
         reanudable por lotes;
      3. arma el calendario de cuartos de hora y agrupa por clave;
      4. agrega las centrales de la hoja "Gen real" del MISMO Excel
         de homologacion, desde la API de operacion real.

    El paso 4 es opcional: si la hoja no existe o esta vacia, se
    escribe solo lo que viene del paso 3.

    Las dos claves de las APIs (son distintas) salen de la seccion
    "claves_api" de config.json, que no se versiona -- ver
    Script/config.py y config.ejemplo.json.
    """

    def avanzar(valor):
        if progreso:
            progreso(valor)

    aamm = validar_aamm(aamm)
    anio, mes = periodo_desde_aamm(aamm)
    periodo = f"{anio}{mes:02d}"
    ultimo_dia = calendar.monthrange(anio, mes)[1]

    rutas = resolver_rutas(carpeta_base)

    if not rutas["base"].is_dir():
        raise ErrorEntrada(f"No se encontro la carpeta base {rutas['base']}")

    archivo_homol = Homologacion.buscar_archivo_homologacion(
        rutas["auxiliares_dir"]
    )

    if not archivo_homol:
        raise ErrorEntrada(
            f"No se encontro el archivo de homologacion en "
            f"{rutas['auxiliares_dir']}: ningun Excel de esa carpeta "
            f"tiene 'homologacion' en el nombre (el real se llama "
            f"'Homologacion ClavesTF y PRMTE.xlsx')."
        )

    registrar(f"Periodo: {anio}-{mes:02d} ({aamm})")

    try:
        registrar(f"Leyendo {archivo_homol.name}...")
        df_homol = Homologacion.leer_homologacion(archivo_homol)
        puntos = Homologacion.puntos_de_medida(df_homol)
        registrar(
            f"  puntos de medida a consultar: {len(puntos):,} "
            f"({df_homol['clave'].nunique()} clave(s))"
        )
        avanzar(5)

        registrar("Descargando medidas por punto de medida...")
        df_crudo, _ = Descarga_PRMTE.descargar(
            puntos, periodo, rutas["trabajo_medidas"],
            registrar=registrar, progreso=progreso, desde=5, hasta=55,
        )

        registrar("Armando calendario y agrupando por clave...")
        df_sae, calendario, diagnostico = Claves_Balance.construir_por_clave(
            df_crudo, df_homol, registrar=registrar
        )
        _resumir_diagnostico_medidas(diagnostico, registrar)
        avanzar(60)

        centrales_api = Homologacion.leer_gen_real(archivo_homol)

        if centrales_api:

            registrar(
                f"Centrales de la hoja "
                f"'{Homologacion.HOJA_GEN_REAL}': {len(centrales_api)}"
            )
            for central in centrales_api:
                registrar(
                    f"    {central['topologyName']} -> "
                    f"{central['clave']}"
                    + ("" if central["factor"] == 1 else
                       f"  (factor {central['factor']:g})")
                )

            df_opreal = Generacion_Real.descargar_mes(
                anio, mes, ultimo_dia, registrar=registrar,
                progreso=progreso, desde=60, hasta=85,
            )
            df_opreal = Generacion_Real.filtrar_y_desempatar(
                df_opreal, centrales_api, registrar=registrar
            )
            df_cuartos = Generacion_Real.expandir_a_cuartos(
                df_opreal, centrales_api, registrar=registrar
            )
            df_cuartos = Generacion_Real.pegar_calendario(
                df_cuartos, calendario, registrar=registrar
            )

            df_sae = pd.concat(
                [df_sae, df_cuartos[Claves_Balance.COLUMNAS_SAE]],
                ignore_index=True,
            )

        else:
            registrar(
                f"La hoja '{Homologacion.HOJA_GEN_REAL}' de "
                f"{archivo_homol.name} no existe o esta vacia: no se "
                f"agrega ninguna central desde la API de operacion real."
            )

    except ErrorMedidas as error:
        raise ErrorEntrada(str(error)) from error

    avanzar(90)

    df_sae = (
        df_sae
        .sort_values(["Cuarto de Hora", "clave"])
        .reset_index(drop=True)
    )

    faltantes = [c for c in COLUMNAS_AI if c not in df_sae.columns]
    if faltantes:
        raise ErrorEntrada(
            f"El resultado no trae las columnas {faltantes} que "
            f"{ARCHIVO_MEDIDAS_SAE} necesita para alimentar Medidores."
        )

    rutas["medidas_dir"].mkdir(parents=True, exist_ok=True)

    registrar(f"Escribiendo {rutas['medidas_sae']}...")
    df_sae[COLUMNAS_AI].to_excel(
        rutas["medidas_sae"], sheet_name=HOJA_MEDIDAS_SAE, index=False
    )

    registrar(f"  filas: {len(df_sae):,}")
    registrar(f"  claves: {df_sae['clave'].nunique()}")
    registrar(
        f"  cuartos de hora: 1 a {int(df_sae['Cuarto de Hora'].max())}"
    )

    avanzar(100)
    registrar(f"Listo: {rutas['medidas_sae']}")

    return rutas["medidas_sae"]
