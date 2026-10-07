# -*- coding: utf-8 -*-
"""
Descarga_PRMTE — baja las mediciones cuarto-horarias de la API de
medidas del Coordinador, punto de medida por punto de medida.

Viene de "1_generacion_prmte.py". Cambios respecto del original:

  - la carpeta de trabajo (lotes descargados + marca de reanudacion)
    se recibe por parametro y vive dentro del caso, no en el
    directorio actual;
  - los lotes y la marca de reanudacion llevan el periodo en el
    nombre: antes, correr dos meses en la misma carpeta mezclaba los
    lotes de los dos ("medidas_batch_*.parquet" los levantaba todos) y
    daba por procesados puntos de otro mes;
  - el user_key sale de config.json (comun.leer_clave_api) en vez de
    estar repetido en cada script;
  - el avance se informa por callback a la ventana en vez de tqdm.

Es la parte lenta del proceso: miles de llamadas HTTP. Por eso se
guarda de a lotes y se puede reanudar: si se corta a la mitad, la
corrida siguiente arranca donde quedo.
"""

import time
from pathlib import Path

import pandas as pd

from .comun import ErrorMedidas, leer_clave_api, CLAVE_PRMTE


URL_MEDIDAS = "https://medidas.api.coordinador.cl/medidas/api/medidas/{periodo}/"

# API de medidas v2 (ruta que el responsable de la API le dio al
# usuario el 2026-10-07; la v1 dejo de traer septiembre 2026 despues del
# 13). Misma informacion, otros nombres: extraer_datos_api_v2() la
# traduce a las columnas de la v1, asi que el resto del proceso no se
# entera. Con USAR_API_V2 = False se vuelve a la v1 sin tocar nada mas.
USAR_API_V2 = True
URL_MEDIDAS_V2 = "https://medidas.api.coordinador.cl/medidas-v2/measurement"

# Como pide el mes la v2: 'period' de 12 digitos AAAAMMDDhhmm (el
# ejemplo que dieron es 202609150000). SIN CONFIRMAR que significa
# exactamente: se pide desde el dia 1 a las 00:00, que sirve tanto si
# devuelve "el mes de esa fecha" como si devuelve "desde esa fecha".
FORMATO_PERIODO_V2 = "{periodo}010000"

CANALES = (1, 3)
TAMANO_LOTE = 100
REINTENTOS = 10
ESPERA_REINTENTO = 1.0
TIMEOUT = 30

# Columnas de cabecera que el original copiaba del registro a cada
# fila de 'mediciones'.
CAMPOS_CABECERA = (
    ("idCoordinado", "idCoordinado"),
    ("idPuntoMedida", "idPuntoMedida"),
    ("periodo", "periodo"),
    ("subEstacion", "subEstacion"),
    ("fechaUltimaLectura", "fechaUltimaLectura"),
)


def _nombre_lote(periodo, numero):
    return f"medidas_batch_{periodo}_{numero:03d}.parquet"


def _archivo_procesados(carpeta_trabajo, periodo):
    return Path(carpeta_trabajo) / f"puntos_procesados_{periodo}.txt"


def _leer_procesados(carpeta_trabajo, periodo):
    archivo = _archivo_procesados(carpeta_trabajo, periodo)

    if not archivo.is_file():
        return set()

    with open(archivo, "r", encoding="utf-8") as f:
        return {linea.strip() for linea in f if linea.strip()}


def _anotar_procesados(carpeta_trabajo, periodo, puntos):
    with open(
        _archivo_procesados(carpeta_trabajo, periodo), "a", encoding="utf-8"
    ) as f:
        for punto in puntos:
            f.write(f"{punto}\n")


# El ultimo error de la API que se vio en esta descarga (HTTP y texto),
# para que el log diga POR QUE un punto vino vacio en vez de solo
# "sin datos". Lo vacia descargar() al empezar.
ULTIMOS_ERRORES = []


def _anotar_error(texto):
    if len(ULTIMOS_ERRORES) < 5:
        ULTIMOS_ERRORES.append(texto)


def _a_numero(serie):
    """Los canales de la v2 vienen como texto: a numero (vacio = NaN)."""

    texto = serie.astype(str).str.strip()
    texto = texto.mask(texto.str.lower().isin(["", "none", "null", "nan"]))
    return pd.to_numeric(
        texto.str.replace(",", ".", regex=False), errors="coerce"
    )


def traducir_v2(registro, id_canal):
    """
    Un registro de la v2 -> el DataFrame de 'mediciones' con las mismas
    columnas que armaba la v1 para ese canal:

        measurement[].dateRange   -> intervalo
        measurement[].utcRange    -> intervaloUtc
        measurement[].principal   -> principal
        measurement[].channel<N>  -> canalVal<N>   (solo el canal pedido)
        coordinatorId / measurePointId / period / subStation /
        lastReadingDate           -> idCoordinado / idPuntoMedida /
                                     periodo / subEstacion /
                                     fechaUltimaLectura
        measurer[0].name          -> nombreMedidor
        channel[] (el de ese channelId) -> slugCanal / descripcionCanal

    Solo se toma el canal pedido y la otra columna queda vacia, igual
    que con la v1 (una llamada por canal): si la v2 devolviera todos los
    canales en cada llamada, sumarlos dos veces duplicaria la energia.
    """

    df = pd.DataFrame(registro.get("measurement") or [])

    if df.empty:
        return df

    df = df.rename(columns={
        "dateRange": "intervalo",
        "utcRange": "intervaloUtc",
    })

    for canal in CANALES:
        df[f"canalVal{canal}"] = (
            _a_numero(df[f"channel{canal}"])
            if canal == id_canal and f"channel{canal}" in df
            else float("nan")
        )

    df = df.drop(
        columns=[c for c in df.columns if str(c).startswith("channel")]
    )

    for campo, destino in (
        ("coordinatorId", "idCoordinado"),
        ("measurePointId", "idPuntoMedida"),
        ("period", "periodo"),
        ("subStation", "subEstacion"),
        ("lastReadingDate", "fechaUltimaLectura"),
    ):
        df[destino] = registro.get(campo, "")

    medidores = registro.get("measurer") or [{}]
    canales = registro.get("channel") or [{}]
    del_canal = next(
        (c for c in canales if str(c.get("channelId")) == str(id_canal)),
        canales[0],
    )

    df["nombreMedidor"] = medidores[0].get("name", "")
    df["descripcionCanal"] = del_canal.get("description", "")
    df["slugCanal"] = del_canal.get("slug", "")

    return df


def extraer_datos_api_v2(sesion, id_punto_medida, id_canal, periodo,
                         user_key):
    """
    Como extraer_datos_api(), contra la v2. Un error 4xx (pedido mal
    armado, sin permiso, punto inexistente) no se reintenta: repetirlo
    da lo mismo y con REINTENTOS x puntos x canales la descarga se
    eternizaba. Se reintenta solo un 5xx, un corte o una respuesta
    vacia.
    """

    params = {
        "channelId": id_canal,
        "measurePointId": id_punto_medida,
        "period": FORMATO_PERIODO_V2.format(periodo=periodo),
    }
    if user_key:
        params["user_key"] = user_key

    for _ in range(REINTENTOS):

        try:
            respuesta = sesion.get(
                URL_MEDIDAS_V2, params=params, timeout=TIMEOUT,
                headers={"accept": "application/json"},
            )

            if respuesta.status_code == 200:
                datos = respuesta.json()
                if datos and isinstance(datos, list):
                    df = traducir_v2(datos[0], id_canal)
                    if not df.empty:
                        return df
                continue

            _anotar_error(
                f"{id_punto_medida} canal {id_canal}: HTTP "
                f"{respuesta.status_code} {respuesta.text[:300]}"
            )
            if 400 <= respuesta.status_code < 500:
                return pd.DataFrame()
            time.sleep(ESPERA_REINTENTO)

        except Exception as error:
            _anotar_error(f"{id_punto_medida} canal {id_canal}: {error}")
            time.sleep(ESPERA_REINTENTO)

    return pd.DataFrame()


def extraer_datos_api(sesion, id_punto_medida, id_canal, periodo, user_key):
    """
    Un punto de medida + un canal. Devuelve el DataFrame de
    'mediciones' con los campos de cabecera pegados, o vacio si no hay
    datos despues de REINTENTOS intentos.
    """

    url = URL_MEDIDAS.format(periodo=periodo)
    params = {
        "idCanal": id_canal,
        "idPuntoMedida": id_punto_medida,
        "user_key": user_key,
    }

    if USAR_API_V2:
        return extraer_datos_api_v2(
            sesion, id_punto_medida, id_canal, periodo, user_key
        )

    for _ in range(REINTENTOS):

        try:
            respuesta = sesion.get(url, params=params, timeout=TIMEOUT)

            if respuesta.status_code == 200:

                datos = respuesta.json()

                if datos and isinstance(datos, list):

                    registro = datos[0]
                    df = pd.DataFrame(registro.get("mediciones", []))

                    if df.empty:
                        continue

                    for campo, destino in CAMPOS_CABECERA:
                        df[destino] = registro.get(campo, "")

                    medidores = registro.get("medidores") or [{}]
                    canales = registro.get("canales") or [{}]

                    df["nombreMedidor"] = medidores[0].get("nombre", "")
                    df["descripcionCanal"] = canales[0].get("descripcion", "")
                    df["slugCanal"] = canales[0].get("slug", "")

                    return df

            else:
                time.sleep(ESPERA_REINTENTO)

        except Exception:
            time.sleep(ESPERA_REINTENTO)

    return pd.DataFrame()


def descargar(
    puntos, periodo, carpeta_trabajo, user_key=None,
    registrar=print, progreso=None, desde=0, hasta=100,
):
    """
    Descarga todos los puntos y devuelve
    (df_consolidado, puntos_fallidos).

    Reanudable: si una corrida anterior se corto a la mitad, los
    puntos ya anotados en puntos_procesados_<periodo>.txt no se vuelven
    a pedir, y sus lotes ya guardados se releen del disco. Si la
    corrida anterior habia terminado (todos los puntos anotados), NO se
    reanuda: se borran sus lotes y se descarga todo de nuevo. Antes se
    reusaba para siempre, y una descarga hecha cuando la API tenia el
    mes a medio publicar quedaba pegada (los puntos salian incompletos
    en cada corrida sin que se volviera a consultar).

    desde/hasta: rango de la barra de progreso de la ventana que le
    toca a esta etapa (la descarga es lo que se lleva casi todo el
    tiempo del proceso).
    """

    try:
        import requests
    except ImportError as error:
        raise ErrorMedidas(
            "Falta la libreria 'requests' (pip install -r "
            "requirements.txt): sin ella no se puede consultar la API "
            "de medidas."
        ) from error

    user_key = user_key or leer_clave_api(CLAVE_PRMTE)

    carpeta_trabajo = Path(carpeta_trabajo)
    carpeta_trabajo.mkdir(parents=True, exist_ok=True)

    ULTIMOS_ERRORES.clear()
    registrar(
        f"  API: {URL_MEDIDAS_V2 if USAR_API_V2 else URL_MEDIDAS.format(periodo=periodo)}"
    )

    procesados = _leer_procesados(carpeta_trabajo, periodo)
    pendientes = [p for p in puntos if str(p) not in procesados]

    if procesados and not pendientes:
        registrar(
            "  la descarga anterior de este periodo estaba completa: se "
            "vuelve a descargar para traer los datos al dia"
        )
        for archivo in carpeta_trabajo.glob(
            f"medidas_batch_{periodo}_*.parquet"
        ):
            archivo.unlink()
        _archivo_procesados(carpeta_trabajo, periodo).unlink()
        procesados = set()
        pendientes = list(puntos)

    if procesados:
        registrar(
            f"  reanudando: {len(procesados):,} punto(s) ya descargados "
            f"en una corrida anterior, quedan {len(pendientes):,}"
        )

    lotes_existentes = sorted(
        carpeta_trabajo.glob(f"medidas_batch_{periodo}_*.parquet")
    )
    numero_lote = len(lotes_existentes) + 1

    fallidos = []
    acumulado = []
    total = max(len(pendientes), 1)

    def avanzar(indice):
        if progreso:
            progreso(desde + (hasta - desde) * indice / total)

    sesion = requests.Session()

    def guardar_lote(resultados):
        nonlocal numero_lote
        archivo = carpeta_trabajo / _nombre_lote(periodo, numero_lote)
        pd.concat(resultados, ignore_index=True).to_parquet(archivo, index=False)
        _anotar_procesados(
            carpeta_trabajo, periodo,
            [p for df in resultados for p in df["idPuntoMedida"].unique()],
        )
        registrar(f"  lote guardado: {archivo.name}")
        numero_lote += 1

    for indice, punto in enumerate(pendientes, start=1):

        partes = [
            extraer_datos_api(sesion, punto, canal, periodo, user_key)
            for canal in CANALES
        ]
        partes = [parte for parte in partes if not parte.empty]

        if not partes:
            fallidos.append(punto)
        else:
            acumulado.append(pd.concat(partes, ignore_index=True))

        if len(acumulado) >= TAMANO_LOTE:
            guardar_lote(acumulado)
            acumulado = []

        if indice % 25 == 0 or indice == len(pendientes):
            registrar(
                f"  puntos consultados: {indice:,}/{len(pendientes):,} "
                f"(fallidos: {len(fallidos):,})"
            )

        avanzar(indice)

    if acumulado:
        guardar_lote(acumulado)

    if fallidos:
        registrar(
            f"  AVISO: {len(fallidos):,} punto(s) sin datos tras "
            f"{REINTENTOS} intentos: {', '.join(map(str, fallidos[:10]))}"
            + (" ..." if len(fallidos) > 10 else "")
        )
        for error in ULTIMOS_ERRORES:
            registrar(f"    respuesta de la API: {error}")

    archivos = sorted(
        carpeta_trabajo.glob(f"medidas_batch_{periodo}_*.parquet")
    )

    if not archivos:
        raise ErrorMedidas(
            f"No se obtuvo ningun dato de la API para el periodo "
            f"{periodo}. Revisa la clave (user_key), la conexion y que "
            f"el periodo ya este publicado."
            + (
                "\n\nLo que respondio la API:\n"
                + "\n".join(f"  {e}" for e in ULTIMOS_ERRORES)
                if ULTIMOS_ERRORES else ""
            )
        )

    df = pd.concat(
        [pd.read_parquet(archivo) for archivo in archivos], ignore_index=True
    )

    registrar(
        f"  registros descargados: {len(df):,} "
        f"({len(archivos)} lote(s))"
    )

    return df, fallidos
