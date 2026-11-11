# -*- coding: utf-8 -*-
"""
El SoC del periodo: se trae desde PI AF (generar_soc, boton "Traer" de
la fila del SoC) y se lee (extraer_soc).

El archivo SOC_AAMM.xlsx tiene ahora TRES columnas (central, fecha,
soc), una fila por central y cuarto de hora. El formato de antes (un
bloque de columnas por central, como lo dejaba PI DataLink) se sigue
leyendo, para los casos viejos.
"""

import calendar

import pandas as pd

from .utiles import ErrorEntrada, normalizar


# ============================================================
# EL ARCHIVO NUEVO: central | fecha | soc
# ============================================================

HOJA_SOC = "SoC"
HOJA_RUTAS_SOC = "Rutas PI"
COLUMNAS_SOC = ["central", "fecha", "soc"]
COLUMNAS_RUTAS_SOC = [
    "central", "estado", "ruta_af", "origen", "registros", "sin_dato",
]


def nombre_archivo_soc(aamm):
    return f"SOC_{aamm}.xlsx"


# ============================================================
# EXTRACCION DEL SoC POR BLOQUES
# ============================================================

def detectar_fila_nombres(df_crudo, maximo_filas=30):
    """
    Busca la fila que contiene los nombres de centrales.

    Criterio: se sube desde la fila de encabezados 'Time Stamp'/
    'Value' hasta encontrar una fila con contenido util, y desde ahi
    se sigue subiendo mientras las filas sigan teniendo contenido
    util (sin saltar un hueco en blanco) -- se devuelve la MAS
    ARRIBA de ese bloque contiguo, no la primera que se encuentra.

    Hace falta este segundo paso porque un archivo real (visto con
    datos reales) puede traer DOS filas de metadata pegadas justo
    arriba del hueco en blanco que precede a los encabezados: la fila
    con el nombre limpio de la central, y debajo (mas cerca de
    'Time Stamp'/'Value') una fila con informacion adicional (ej. la
    ruta SCADA completa del punto). Quedarse con "la primera fila util
    subiendo" agarra la fila equivocada (la de mas informacion, no la
    del nombre limpio) en ese caso -- subir hasta el tope del bloque
    contiguo trae la de mas arriba, que es la que tiene el nombre.
    """

    fila_encabezados = None

    for indice in range(min(maximo_filas, len(df_crudo))):

        textos = {
            normalizar(v)
            for v in df_crudo.iloc[indice].tolist()
        }

        if "time stamp" in textos and "value" in textos:
            fila_encabezados = indice
            break

    if fila_encabezados is None:
        raise ErrorEntrada(
            "No se encontro ninguna fila con los encabezados "
            "'Time Stamp' y 'Value' en el archivo SOC."
        )

    encabezados_bloque = {
        "status",
        "questionable",
        "time stamp",
        "value",
        "",
    }

    def _fila_tiene_contenido_util(indice):
        textos = [
            normalizar(v)
            for v in df_crudo.iloc[indice].tolist()
        ]
        return any(t not in encabezados_bloque for t in textos)

    fila_nombres = None

    for indice in range(fila_encabezados - 1, -1, -1):

        if _fila_tiene_contenido_util(indice):
            # Sigue siendo candidata mientras haya contenido util;
            # se actualiza en cada vuelta para quedarse con la MAS
            # ARRIBA del bloque contiguo, no la primera encontrada.
            fila_nombres = indice
            continue

        # Fila en blanco: si ya se encontro una candidata, el bloque
        # contiguo termino aca -- se corta la busqueda.
        if fila_nombres is not None:
            break

    if fila_nombres is None:
        raise ErrorEntrada(
            "Se encontraron los encabezados 'Time Stamp'/'Value' "
            "pero no una fila de nombres de centrales arriba."
        )

    return fila_nombres, fila_encabezados


def detectar_bloques(df_crudo, fila_nombres, fila_encabezados):
    """
    Para cada nombre de la fila de nombres, define su bloque
    horizontal y ubica dentro sus columnas Time Stamp y Value.

    No se usan offsets fijos: los encabezados se buscan por
    texto dentro del rango de cada central.
    """

    encabezados_bloque = {
        "status",
        "questionable",
        "time stamp",
        "value",
        "",
    }

    fila_n = df_crudo.iloc[fila_nombres].tolist()

    posiciones = []

    for indice, valor in enumerate(fila_n):
        texto = normalizar(valor)
        if texto and texto not in encabezados_bloque:
            posiciones.append((indice, str(valor).strip()))

    if not posiciones:
        raise ErrorEntrada(
            "La fila de nombres no contiene ninguna central."
        )

    fila_e = [
        normalizar(v)
        for v in df_crudo.iloc[fila_encabezados].tolist()
    ]

    bloques = []
    incidencias = []

    for orden, (columna_inicio, nombre) in enumerate(posiciones):

        if orden + 1 < len(posiciones):
            columna_fin = posiciones[orden + 1][0] - 1
        else:
            columna_fin = len(fila_e) - 1

        columnas_ts = [
            c
            for c in range(columna_inicio, columna_fin + 1)
            if c < len(fila_e) and fila_e[c] == "time stamp"
        ]

        columnas_val = [
            c
            for c in range(columna_inicio, columna_fin + 1)
            if c < len(fila_e) and fila_e[c] == "value"
        ]

        if len(columnas_ts) != 1 or len(columnas_val) != 1:
            incidencias.append(
                f"{nombre}: se esperaba exactamente un "
                f"'Time Stamp' y un 'Value' entre las columnas "
                f"{columna_inicio + 1} y {columna_fin + 1}; "
                f"se encontraron {len(columnas_ts)} y "
                f"{len(columnas_val)}. Bloque excluido."
            )
            continue

        bloques.append(
            {
                "nombre_bess_origen": nombre,
                "columna_inicio_bloque": columna_inicio,
                "columna_fin_bloque": columna_fin,
                "columna_timestamp": columnas_ts[0],
                "columna_value": columnas_val[0],
            }
        )

    return bloques, incidencias


def _extraer_nombre_desde_ruta_scada(texto):
    r"""
    Si el nombre de un bloque de SoC viene como una ruta SCADA (visto
    con datos reales: exportaciones tipo PI traen el nombre de la
    central como
        \\SERVIDOR\SEN\Generación\SEN\<region>\<central>|<sufijo>
    en vez de solo "<central>"), devuelve unicamente "<central>": el
    ultimo tramo de la ruta (separado por barras invertidas), sin el
    sufijo despues de "|".

    Este docstring es CRUDO (lleva una r adelante) justamente por esas
    barras invertidas: si no, la ruta del ejemplo no seria texto sino
    escapes, y Python avisa "SyntaxWarning: invalid escape sequence"
    al importar el modulo.

    No es una reinterpretacion de datos: es separar una ESTRUCTURA
    conocida (ruta + sufijo) que ya viene asi en el archivo, no una
    suposicion sobre a que central corresponde. Si el texto no tiene
    ese formato (ninguna barra invertida), se devuelve tal cual -- no
    se inventa nada quitando texto de un nombre que no es una ruta.
    """

    texto = str(texto).strip()

    if "\\" not in texto:
        return texto

    return texto.split("\\")[-1].split("|")[0].strip()


def columnas_tres_columnas(df_crudo):
    """
    {"central": i, "fecha": j, "soc": k} si la primera fila con algo
    del archivo es el encabezado del formato nuevo; None si es el
    formato de bloques de antes.
    """

    for indice in range(min(5, len(df_crudo))):
        textos = [normalizar(v) for v in df_crudo.iloc[indice].tolist()]
        if not any(textos):
            continue
        columnas = {
            nombre: textos.index(nombre)
            for nombre in COLUMNAS_SOC if nombre in textos
        }
        if len(columnas) == len(COLUMNAS_SOC):
            columnas["_fila"] = indice
            return columnas
        return None

    return None


def _extraer_soc_tres_columnas(df_crudo, columnas, mapa_homologacion):
    """extraer_soc() para el formato nuevo (central | fecha | soc)."""

    sub = df_crudo.iloc[
        columnas["_fila"] + 1:,
        [columnas["central"], columnas["fecha"], columnas["soc"]],
    ].copy()
    sub.columns = ["nombre_scada_original", "timestamp", "soc"]

    sub["nombre_scada_original"] = (
        sub["nombre_scada_original"].map(
            lambda v: "" if normalizar(v) == "" else str(v).strip()
        )
    )
    sub["timestamp"] = pd.to_datetime(sub["timestamp"], errors="coerce")
    sub["soc"] = pd.to_numeric(sub["soc"], errors="coerce")
    sub = sub[
        (sub["nombre_scada_original"] != "") & sub["timestamp"].notna()
    ]

    if sub.empty:
        raise ErrorEntrada(
            "El archivo de SoC tiene el encabezado central | fecha | soc "
            "pero ninguna fila con datos."
        )

    sub["central"] = sub["nombre_scada_original"].map(
        lambda v: mapa_homologacion.get(normalizar(v), v)
    )

    incidencias = []

    for nombre, grupo in sub.groupby("nombre_scada_original", sort=False):

        vacios = int(grupo["soc"].isna().sum())
        if vacios:
            incidencias.append(
                f"{nombre}: {vacios} valores de SoC vacios o no numericos."
            )

        duplicados = int(grupo["timestamp"].duplicated().sum())
        if duplicados:
            incidencias.append(
                f"{nombre}: {duplicados} timestamps duplicados (en abril, "
                f"los de la hora que se repite por el cambio de hora son "
                f"normales)."
            )

    df_soc = sub[
        ["central", "timestamp", "soc", "nombre_scada_original"]
    ].reset_index(drop=True)

    return df_soc, incidencias


def extraer_soc(ruta_soc, mapa_homologacion=None):
    """
    Lee el archivo SOC y devuelve (df_soc, incidencias). Acepta el
    formato nuevo (central | fecha | soc, el que deja generar_soc) y el
    de bloques de antes.

    df_soc: central | timestamp | soc | nombre_scada_original
    """

    # Motor rapido (lector.py): con el archivo real de SoC
    # (docs/SOC_real_2607.xlsx) entrega exactamente el mismo DataFrame,
    # fechas incluidas, 6 veces mas rapido.
    from .lector import leer_excel

    df_crudo = leer_excel(
        ruta_soc,
        sheet_name=0,
        header=None,
    )

    columnas = columnas_tres_columnas(df_crudo)
    if columnas is not None:
        return _extraer_soc_tres_columnas(
            df_crudo, columnas, mapa_homologacion or {}
        )

    fila_nombres, fila_encabezados = detectar_fila_nombres(
        df_crudo
    )

    bloques, incidencias = detectar_bloques(
        df_crudo,
        fila_nombres,
        fila_encabezados,
    )

    mapa_homologacion = mapa_homologacion or {}

    partes = []

    for bloque in bloques:

        nombre_origen = bloque["nombre_bess_origen"]

        # Primero se prueba el texto literal (compatibilidad con
        # cualquier archivo de SoC "limpio", sin ruta SCADA). Si no
        # hay match, se prueba de nuevo con el nombre extraido de la
        # ruta (ver _extraer_nombre_desde_ruta_scada) -- el
        # Diccionario puede tener registrada cualquiera de las dos
        # formas. Si ninguna tiene match, se usa igual el nombre
        # extraido (no la ruta completa) como "canonico": aunque no
        # homologue, es mucho mas legible en avisos/incidencias que
        # la ruta cruda, y no cambia el comportamiento (sigue sin
        # cruzar contra Medidores).
        clave_directa = normalizar(nombre_origen)
        nombre_limpio = _extraer_nombre_desde_ruta_scada(nombre_origen)

        if clave_directa in mapa_homologacion:
            canonico = mapa_homologacion[clave_directa]
        else:
            canonico = mapa_homologacion.get(
                normalizar(nombre_limpio),
                nombre_limpio,
            )

        sub = df_crudo.iloc[
            fila_encabezados + 1:,
            [
                bloque["columna_timestamp"],
                bloque["columna_value"],
            ],
        ].copy()

        sub.columns = ["timestamp", "soc"]

        sub["timestamp"] = pd.to_datetime(
            sub["timestamp"],
            errors="coerce",
        )

        sub["soc"] = pd.to_numeric(
            sub["soc"],
            errors="coerce",
        )

        sub = sub.dropna(subset=["timestamp"])

        if sub.empty:
            incidencias.append(
                f"{nombre_origen}: bloque sin registros validos."
            )
            continue

        if sub["soc"].isna().any():
            incidencias.append(
                f"{nombre_origen}: "
                f"{int(sub['soc'].isna().sum())} valores de SoC "
                f"no numericos."
            )

        duplicados = sub["timestamp"].duplicated().sum()

        if duplicados:
            incidencias.append(
                f"{nombre_origen}: {duplicados} timestamps "
                f"duplicados (en abril, los de la hora que se repite "
                f"por el cambio de hora son normales)."
            )

        sub["central"] = canonico
        sub["nombre_scada_original"] = nombre_origen

        partes.append(sub)

    if not partes:
        raise ErrorEntrada(
            "No se pudo extraer SoC de ningun bloque.\n"
            + "\n".join(incidencias)
        )

    df_soc = pd.concat(partes, ignore_index=True)

    df_soc = df_soc[
        [
            "central",
            "timestamp",
            "soc",
            "nombre_scada_original",
        ]
    ]

    return df_soc, incidencias


# ============================================================
# TRAER EL SoC DESDE PI AF (boton "Traer" de la fila del SoC)
# ============================================================

def _columna_nombre(resumen):
    return next(
        (c for c in resumen.columns if "nombre" in normalizar(c)),
        resumen.columns[0],
    )


def centrales_para_soc(resumen, diccionario):
    """
    Lo que hay que pedirle a PI: una entrada por central de 'Resumen
    BESS', en ese orden, con la ruta de la columna SoC_PI del
    Diccionario (si esta) y todos sus nombres (para la busqueda por
    Nombre PCP/PID cuando no hay ruta).
    """

    from .lectura import (
        ROL_BALANCE_BESS, ROL_SOC_PI, ROLES_NO_NOMBRE, _tiene_valor,
        filas_diccionario,
    )

    fila_de = {}
    for fila in filas_diccionario(diccionario):
        fila_de.setdefault(normalizar(fila.get(ROL_BALANCE_BESS)), fila)

    columna = _columna_nombre(resumen)
    centrales = []
    vistas = set()

    for valor in resumen[columna].tolist():

        if not _tiene_valor(valor):
            continue
        nombre = str(valor).strip()
        if normalizar(nombre) in vistas:
            continue
        vistas.add(normalizar(nombre))

        fila = fila_de.get(normalizar(nombre), {})
        nombres = [nombre] + [
            str(v).strip() for rol, v in fila.items()
            if rol not in ROLES_NO_NOMBRE and _tiene_valor(v)
            and normalizar(v) != normalizar(nombre)
        ]

        centrales.append({
            "central": nombre,
            "ruta": fila.get(ROL_SOC_PI),
            "nombres": nombres,
        })

    return centrales


def rango_del_periodo(aamm):
    """('AAAA-MM-01 00:00:00', 'primero del mes siguiente 00:00:00')."""

    from .rutas import periodo_desde_aamm

    anio, mes = periodo_desde_aamm(aamm)
    ultimo = calendar.monthrange(anio, mes)[1]
    fin = pd.Timestamp(anio, mes, ultimo) + pd.Timedelta(days=1)

    return (
        f"{anio}-{mes:02d}-01 00:00:00",
        fin.strftime("%Y-%m-%d %H:%M:%S"),
    )


def armar_tabla_soc(filas):
    """DataFrame central | fecha | soc desde las filas de Soc_PI.descargar."""

    df = pd.DataFrame(filas, columns=COLUMNAS_SOC)
    df["fecha"] = pd.to_datetime(df["fecha"])
    df["soc"] = pd.to_numeric(df["soc"], errors="coerce")
    return df


def escribir_soc(ruta, df_soc, df_rutas):
    """SOC_AAMM.xlsx: hoja 'SoC' (las tres columnas) y 'Rutas PI'."""

    from .formato import formatear_hoja

    try:
        with pd.ExcelWriter(ruta, engine="openpyxl") as writer:
            df_soc.to_excel(writer, sheet_name=HOJA_SOC, index=False)
            df_rutas.to_excel(writer, sheet_name=HOJA_RUTAS_SOC, index=False)
            for hoja in (HOJA_SOC, HOJA_RUTAS_SOC):
                formatear_hoja(writer.book[hoja])
    except PermissionError as error:
        raise ErrorEntrada(
            f"No se pudo escribir {ruta}: cerralo en Excel y volve a "
            f"apretar el boton."
        ) from error


def generar_soc(carpeta_base, aamm, registrar=print, progreso=None,
                conexion=None):
    """
    Trae desde PI AF el SoC de todas las centrales de 'Resumen BESS'
    (Centrales.xlsx) y lo deja en <CARPETA_BASE>/Medidas/SOC_AAMM.xlsx,
    en tres columnas: central | fecha | soc.

    La ruta de cada central sale de la columna SoC_PI del Diccionario
    (las de AF1 se buscan en AF2); si no hay, se busca el elemento por
    su atributo "Nombre PCP/PID" (ver Script/Medidas/Soc_PI.py). La
    hoja 'Rutas PI' del archivo dice de donde salio cada una.

    conexion: solo para las pruebas (una ConexionAF falsa).
    """

    from .externos import ErrorMedidas, Soc_PI
    from .lectura import leer_centrales
    from .rutas import buscar_soc, resolver_rutas, validar_aamm

    def avanzar(valor):
        if progreso:
            progreso(valor)

    aamm = validar_aamm(aamm)
    rutas = resolver_rutas(carpeta_base, aamm)

    if not rutas["base"].is_dir():
        raise ErrorEntrada(f"No se encontro la carpeta base {rutas['base']}")

    if not rutas["centrales"].is_file():
        raise ErrorEntrada(f"No se encontro {rutas['centrales']}")

    # Se lee Centrales.xlsx y no la foto de la planilla: es donde se
    # pegan las rutas, y el SoC se trae antes de armar la planilla.
    registrar(f"Leyendo {rutas['centrales'].name}...")
    resumen, diccionario = leer_centrales(rutas["centrales"])
    centrales = centrales_para_soc(resumen, diccionario)

    if not centrales:
        raise ErrorEntrada(
            "La hoja 'Resumen BESS' de Centrales.xlsx no tiene centrales."
        )

    con_ruta = sum(1 for c in centrales if Soc_PI.limpiar_ruta(c["ruta"])[0])
    registrar(
        f"  centrales: {len(centrales)} ({con_ruta} con ruta en la columna "
        f"SoC_PI del Diccionario)"
    )

    inicio, fin = rango_del_periodo(aamm)
    registrar(
        f"Trayendo '{Soc_PI.ATRIBUTO_SOC}' cada {Soc_PI.INTERVALO} desde "
        f"{Soc_PI.SERVIDOR_AF}, {inicio} a {fin}..."
    )
    avanzar(5)

    try:
        filas, resumen_rutas = Soc_PI.descargar(
            centrales, inicio, fin, registrar=registrar, progreso=progreso,
            desde=10, hasta=85, conexion=conexion,
        )
    except ErrorMedidas as error:
        raise ErrorEntrada(str(error)) from error

    if not filas:
        raise ErrorEntrada(
            "No se pudo traer el SoC de ninguna central desde PI. "
            "Revisa las rutas de la columna SoC_PI del Diccionario."
        )

    df_soc = armar_tabla_soc(filas)
    df_rutas = pd.DataFrame(resumen_rutas, columns=COLUMNAS_RUTAS_SOC)
    avanzar(90)

    rutas["medidas_dir"].mkdir(parents=True, exist_ok=True)
    destino = rutas["medidas_dir"] / nombre_archivo_soc(aamm)

    registrar(f"Escribiendo {destino}...")
    escribir_soc(destino, df_soc, df_rutas)
    registrar(
        f"  filas: {len(df_soc):,}   centrales: {df_soc['central'].nunique()}"
    )

    faltan = df_rutas[df_rutas["estado"] != "OK"]
    if not faltan.empty:
        registrar(
            f"AVISO: {len(faltan)} central(es) sin SoC: "
            f"{', '.join(faltan['central'])}. Pega su ruta de AF en la "
            f"columna SoC_PI del Diccionario y volve a traerlo."
        )

    buscadas = df_rutas[df_rutas["origen"] == Soc_PI.ORIGEN_BUSQUEDA]
    if not buscadas.empty:
        registrar(
            "Centrales encontradas por Nombre PCP/PID (se puede pegar la "
            "ruta en la columna SoC_PI del Diccionario):"
        )
        for _, fila in buscadas.iterrows():
            registrar(f"    {fila['central']}: {fila['ruta_af']}")

    # El programa usa EL archivo de SoC del periodo: si quedo otro al
    # lado (el de DataLink de antes, con otro nombre), hay que sacarlo.
    try:
        buscar_soc(rutas["medidas_dir"], aamm)
    except ErrorEntrada as error:
        registrar(f"AVISO: {str(error).splitlines()[0]}")

    avanzar(100)
    registrar(f"Listo: {destino}")

    return destino
