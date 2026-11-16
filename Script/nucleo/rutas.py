# -*- coding: utf-8 -*-
"""
Rutas del caso y busqueda de los archivos de entrada.
"""

import re
from pathlib import Path

from .carpetas import carpeta_fd, raiz_herramienta
from .parametros import (
    ARCHIVO_CENTRALES, CARPETA_BESS, ARCHIVO_CMG, ARCHIVO_CONTROL, ARCHIVO_MEDIDAS_SAE,
    ARCHIVO_PUNTOS_FALLIDOS,
    ARCHIVO_REGISTRO, ARCHIVO_SALIDA_ANTIGUO, CARPETA_AUXILIARES,
    CARPETA_CMG, CARPETA_LOG, PREFIJO_CSV_SAE, PREFIJO_SALIDA,
    CARPETA_DB_SUBASTAS, CARPETA_FD_FMA,
    CARPETA_MEDIDAS, CARPETA_OFERTAS, CARPETA_PRORRATA_RETIROS,
    CARPETA_SUBASTAS,
    CARPETA_TRABAJO_MEDIDAS, EXTENSIONES_EXCEL, PATRON_AAMM,
    PATRON_NOMBRE_OFERTAS, PATRON_NOMBRE_SSCC_DESEMPENO,
)
from .utiles import ErrorEntrada, normalizar


# ============================================================
# RESOLUCION DE RUTAS
# ============================================================

def nombre_salida(aamm):
    """'2609' -> '11_PAGOS_BESS_2609.xlsx'."""

    return f"{PREFIJO_SALIDA}{aamm}.xlsx"


def nombre_csv_sae(aamm):
    """'2609' -> 'SAE_2609.csv'."""

    return f"{PREFIJO_CSV_SAE}{aamm}.csv"


def _aamm_o_none(aamm):
    texto = str(aamm).strip() if aamm is not None else ""
    return texto if re.fullmatch(PATRON_AAMM, texto) else None


def _salida_sin_periodo(base):
    """
    Sin un AAMM valido: la 11_PAGOS_BESS_AAMM.xlsx que haya en la
    carpeta (cada caso es de un solo periodo), o el nombre de antes si
    solo esta ese, o un nombre generico.
    """

    try:
        encontradas = sorted(
            p for p in base.glob(f"{PREFIJO_SALIDA}*.xlsx")
            if re.fullmatch(rf"{re.escape(PREFIJO_SALIDA)}\d{{4}}\.xlsx", p.name)
        )
    except OSError:
        encontradas = []
    if len(encontradas) == 1:
        return encontradas[0]
    if (base / ARCHIVO_SALIDA_ANTIGUO).is_file():
        return base / ARCHIVO_SALIDA_ANTIGUO
    return base / f"{PREFIJO_SALIDA}AAMM.xlsx"


def _mover(origen, destino):
    """Mueve origen a destino si destino no esta; True si quedo en destino."""

    if destino.exists():
        return True
    if not origen.exists():
        return False
    try:
        destino.parent.mkdir(parents=True, exist_ok=True)
        origen.replace(destino)
        return True
    except OSError:
        # Abierto en Excel, sin permiso...: queda donde estaba y se
        # vuelve a intentar la proxima vez.
        return False


def _migrar_archivos_antiguos(base, salida):
    """
    Un caso de antes del cambio de nombres (pedido del usuario
    2026-11-04): Balance_BESS.xlsx pasa a 11_PAGOS_BESS_AAMM.xlsx y el
    control y el registro, a la carpeta Log/, y los sellos (json), a
    __config__ (ver sellos.ruta_sellos). Nunca pisa un
    archivo que ya este. Devuelve la ruta de la planilla a usar: la
    nueva, o la vieja si no se pudo renombrar (estaba abierta).
    """

    log = base / CARPETA_LOG
    from .sellos import mover_sellos_antiguos

    vieja = base / ARCHIVO_SALIDA_ANTIGUO
    if vieja.is_file() and not salida.exists():
        if not _mover(vieja, salida):
            salida = vieja
    # Los sellos (json) van a __config__, nunca a la carpeta del caso.
    mover_sellos_antiguos(salida, vieja)
    for nombre in (ARCHIVO_CONTROL, ARCHIVO_REGISTRO):
        if (base / nombre).is_file():
            _mover(base / nombre, log / nombre)
    return salida


def raiz_bess(carpeta_base):
    """
    Donde van las carpetas propias del BESS y su planilla:
    <carpeta>/BESS/, o la carpeta misma en un caso de antes de la
    division BESS/CRA (ver carpetas.py).
    """

    return raiz_herramienta(
        carpeta_base, CARPETA_BESS, CARPETAS_DISTINTIVAS,
        archivos=(f"{PREFIJO_SALIDA}*.xlsx", ARCHIVO_SALIDA_ANTIGUO),
    )


def resolver_rutas(carpeta_base, aamm=None):
    """
    Deriva todas las rutas desde la carpeta base elegida por el
    usuario. Nada depende de donde vive el .py.

    La carpeta base es la del MES (la misma que usa Balance_CRA.py):
    lo propio del BESS va en <base>/BESS/ y lo comun con el CRA (FD y
    FMA, Prorrata retiros) en la raiz. Un caso de antes, con todo
    suelto en la raiz, se sigue usando tal cual (ver carpetas.py).

    aamm: el periodo del caso. La planilla se llama
    11_PAGOS_BESS_AAMM.xlsx; sin un AAMM valido se usa la que haya en la
    carpeta (ver _salida_sin_periodo).
    """

    base = Path(carpeta_base)
    raiz = raiz_bess(base)

    medidas_dir = raiz / CARPETA_MEDIDAS
    auxiliares_dir = raiz / CARPETA_AUXILIARES
    ofertas_dir = raiz / CARPETA_OFERTAS
    subastas_dir = raiz / CARPETA_SUBASTAS
    # Comunes con el CRA, en la raiz del mes (Cmg, FD y FMA, Prorrata
    # retiros). La del FD se llama "FD y
    # FMA"; si un caso viejo todavia tiene la de nombre anterior y no
    # la nueva, se usa esa.
    fd_fma_dir = carpeta_fd(base)
    cmg_dir = base / CARPETA_CMG
    prorrata_retiros_dir = base / CARPETA_PRORRATA_RETIROS

    periodo = _aamm_o_none(aamm)
    salida = (
        raiz / nombre_salida(periodo) if periodo else _salida_sin_periodo(raiz)
    )
    if raiz.is_dir():
        salida = _migrar_archivos_antiguos(raiz, salida)

    return {
        "base": base,
        "raiz_bess": raiz,
        "medidas_dir": medidas_dir,
        "auxiliares_dir": auxiliares_dir,
        "ofertas_dir": ofertas_dir,
        "cmg_dir": cmg_dir,
        "sscc_desempeno_dir": fd_fma_dir,
        "subastas_dir": subastas_dir,
        "prorrata_retiros_dir": prorrata_retiros_dir,
        "db_subastas_dir": subastas_dir / CARPETA_DB_SUBASTAS,
        "medidas_sae": medidas_dir / ARCHIVO_MEDIDAS_SAE,
        "puntos_fallidos": medidas_dir / ARCHIVO_PUNTOS_FALLIDOS,
        "trabajo_medidas": medidas_dir / CARPETA_TRABAJO_MEDIDAS,
        "centrales": auxiliares_dir / ARCHIVO_CENTRALES,
        "cmg": cmg_dir / ARCHIVO_CMG,
        # Una sola planilla (entradas + calculo), el CSV de pagos al
        # lado y lo de cada corrida en Log/.
        "salida": salida,
        "csv_sae": raiz / nombre_csv_sae(periodo or "AAMM"),
        "log_dir": raiz / CARPETA_LOG,
        "control": raiz / CARPETA_LOG / ARCHIVO_CONTROL,
        "registro": raiz / CARPETA_LOG / ARCHIVO_REGISTRO,
    }


# ============================================================
# CREAR EL CASO DE UN PERIODO NUEVO
# ============================================================

# Las carpetas propias del BESS (van en <mes>/BESS/), en el orden en que
# se usan. Sirven tambien para reconocer un caso de antes de la division
# BESS/CRA: si alguna de estas esta suelta en la raiz y no hay BESS/, el
# caso es de antes y se usa tal cual (ver carpetas.py).
#
# "Subastas/DB subastas" va con su carpeta adentro porque la ventana ya
# la creaba sola al revisar (ofertas_adj.asegurar_carpeta_db); crearla
# aca deja el caso entero armado de una vez.
#
# NO se crea Medidas/_trabajo: son los andamios de la descarga, los
# arma generar_medidas_sae cuando corre.
CARPETAS_PROPIAS = (
    CARPETA_MEDIDAS,
    CARPETA_AUXILIARES,
    CARPETA_OFERTAS,
    CARPETA_SUBASTAS,
    f"{CARPETA_SUBASTAS}/{CARPETA_DB_SUBASTAS}",
)

# Las que delatan un caso de antes: las propias que el CRA no tiene
# (Auxiliares/ suelta puede ser de un caso CRA de antes).
CARPETAS_DISTINTIVAS = tuple(
    nombre for nombre in CARPETAS_PROPIAS
    if nombre != CARPETA_AUXILIARES and "/" not in nombre
)

# Las carpetas que tiene que tener un caso NUEVO, relativas a la carpeta
# del mes: las propias adentro de BESS/ y las comunes con el CRA en la
# raiz. Es la misma estructura que revisa revisar_estructura() y que
# dibuja la ventana: si aca falta una, el caso nace incompleto. Para un
# caso que ya existe, ver subcarpetas_del_caso().
SUBCARPETAS_CASO = tuple(
    f"{CARPETA_BESS}/{nombre}" for nombre in CARPETAS_PROPIAS
) + (CARPETA_CMG, CARPETA_FD_FMA, CARPETA_PRORRATA_RETIROS)


def subcarpetas_del_caso(carpeta_base):
    """
    Las subcarpetas que tiene que tener ESTE caso, relativas a su
    carpeta: SUBCARPETAS_CASO para uno nuevo o ya dividido; las de
    antes (todo suelto en la raiz) para uno de antes de la division, asi
    completarlo no le crea un BESS/ vacio al lado (que ademas lo haria
    pasar a buscar todo ahi). Lo mismo con "FD y FMA": si el caso usa
    la carpeta de nombre anterior, no se le crea la nueva.
    """

    base = Path(carpeta_base)
    raiz = raiz_bess(base)
    fd = carpeta_fd(base)

    propias = [
        Path(nombre) if raiz == base else Path(CARPETA_BESS) / nombre
        for nombre in CARPETAS_PROPIAS
    ]
    comunes = [Path(CARPETA_CMG), Path(fd.name), Path(CARPETA_PRORRATA_RETIROS)]

    return tuple(ruta.as_posix() for ruta in propias + comunes)


def faltan_subcarpetas(carpeta_base):
    """Las de subcarpetas_del_caso() que todavia no estan."""

    base = Path(carpeta_base)
    return [
        sub for sub in subcarpetas_del_caso(base) if not (base / sub).is_dir()
    ]


def crear_estructura_caso(carpeta_base, registrar=print):
    """
    Crea la carpeta base de un caso y todas sus subcarpetas
    (subcarpetas_del_caso). Devuelve la lista de las que se crearon (las
    que ya estaban no se tocan ni se cuentan).

    Pedido del usuario: *"cuando yo elija un mes que no exista, me
    permita elegir y crear la carpeta con las carpetas dentro"*.

    Es idempotente a proposito: sirve tanto para un periodo nuevo
    (crea todo) como para completar un caso al que le falta una
    carpeta (crea solo esa). Nunca borra ni mueve nada.

    Un caso VIEJO no recibe carpetas vacias de la estructura nueva: si
    tiene todo suelto en la raiz no se le crea BESS/, y si tiene la
    carpeta "SSCC_Desempeño" no se le crea una "FD y FMA" al lado (ver
    subcarpetas_del_caso).

    Si la carpeta ya tiene CRA/ (el mes lo empezo Balance_CRA.py), se
    agrega BESS/ al lado: es la misma carpeta del mes.
    """

    base = Path(carpeta_base)

    if base.exists() and not base.is_dir():
        raise ErrorEntrada(
            f"{base} existe y no es una carpeta: elegi otro nombre."
        )

    creadas = []

    try:
        if not base.is_dir():
            base.mkdir(parents=True)
            creadas.append(base)

        for relativa in subcarpetas_del_caso(base):

            carpeta = base / relativa

            if not carpeta.is_dir():
                carpeta.mkdir(parents=True)
                creadas.append(carpeta)

    except OSError as error:
        raise ErrorEntrada(
            f"No se pudo crear la estructura del caso en {base}: {error}"
        ) from error

    for carpeta in creadas:
        registrar(f"  creada: {carpeta}")

    if not creadas:
        registrar(f"  {base} ya tenia todas sus carpetas.")

    return creadas


def nombre_caso_sugerido(aamm, carpeta_modelo=None):
    """
    Como llamar a la carpeta del periodo 'aamm', a partir de como se
    llama la del periodo que el usuario tenia abierto.

    Si esa carpeta trae un AAMM en el nombre, se cambia SOLO ese
    pedazo y se conserva el resto ("Balance BESS 2607" -> "Balance
    BESS 2608"): cada equipo nombra sus casos a su manera y no hay una
    convencion que imponer. Si no trae ninguno (o no hay carpeta
    previa), se propone el AAMM pelado.

    Es solo una SUGERENCIA: la ventana la muestra en un campo que el
    usuario puede cambiar antes de crear nada.
    """

    aamm = validar_aamm(aamm)

    if carpeta_modelo:
        nombre = Path(carpeta_modelo).name
        cambiado, cuantos = re.subn(r"(?<!\d)\d{4}(?!\d)", aamm, nombre, count=1)
        if cuantos and cambiado != nombre:
            return cambiado

    return aamm


# Que relacion tiene una carpeta con un periodo (ver
# carpeta_corresponde_al_periodo). Son tres respuestas, no dos: "no se
# sabe" NO es lo mismo que "no", y confundirlos es justo lo que hacia
# que la ventana se quedara en la carpeta del mes anterior sin decir
# nada.
CARPETA_DEL_PERIODO = "si"
CARPETA_DE_OTRO_PERIODO = "no"
CARPETA_SIN_PERIODO = "no_se_sabe"


def periodos_en_el_nombre(nombre):
    """
    Los AAMM que trae el nombre de una carpeta: grupos de 4 digitos
    sueltos (ni 3 ni 5). "Balance BESS 2607" -> ["2607"].
    """

    return re.findall(r"(?<!\d)\d{4}(?!\d)", str(nombre))


def carpeta_corresponde_al_periodo(carpeta, aamm, carpetas_recordadas=None):
    """
    Si 'carpeta' es la del periodo 'aamm'. Devuelve una de las tres
    constantes CARPETA_*: si, no, o no se sabe.

    Regla del usuario: *"dos meses no pueden tener la misma carpeta"*.
    Por eso una carpeta que YA es la de otro periodo se responde que
    NO, aunque su nombre no diga nada: lo que manda es
    carpetas_recordadas ({aamm: carpeta}, lo que la ventana fue
    guardando en config.json), y recien despues el nombre.

    Sin registro y sin ningun AAMM en el nombre no hay con que
    decidir: eso es CARPETA_SIN_PERIODO, y quien llama tiene que
    PREGUNTAR, no suponer que sirve.
    """

    if not carpeta:
        return CARPETA_DE_OTRO_PERIODO

    carpeta = Path(carpeta)

    if not carpeta.is_dir():
        return CARPETA_DE_OTRO_PERIODO

    recordadas = {
        str(periodo): Path(ruta)
        for periodo, ruta in (carpetas_recordadas or {}).items()
    }

    def misma(a, b):
        try:
            return a.resolve() == b.resolve()
        except OSError:
            return str(a) == str(b)

    if aamm in recordadas and misma(recordadas[aamm], carpeta):
        return CARPETA_DEL_PERIODO

    for periodo, ruta in recordadas.items():
        if periodo != aamm and misma(ruta, carpeta):
            return CARPETA_DE_OTRO_PERIODO

    periodos = periodos_en_el_nombre(carpeta.name)

    if not periodos:
        return CARPETA_SIN_PERIODO

    return (
        CARPETA_DEL_PERIODO if aamm in periodos else CARPETA_DE_OTRO_PERIODO
    )


def validar_aamm(aamm):
    """Levanta ErrorEntrada si aamm no son 4 digitos (ej. '2607')."""

    if not aamm or not PATRON_AAMM.match(str(aamm).strip()):
        raise ErrorEntrada(
            "Ingresa el periodo AAMM en la ventana (4 digitos, "
            "por ejemplo 2607 para julio de 2026)."
        )
    return str(aamm).strip()


def _es_archivo_de_soc(nombre_archivo, aamm):
    """
    El nombre del archivo de SoC no sigue un patron fijo (no es
    literalmente 'SOC_AAMM.xlsx'): basta con que contenga 'SOC' y el
    AAMM del periodo, en cualquier posicion y con cualquier separador
    (plan de migracion, seccion 19.1). El archivo en si siempre es
    .xlsx (ver buscar_soc) - no confundir con otros archivos del caso
    que puedan compartir "SOC"+AAMM en el nombre sin serlo.
    """

    nombre = normalizar(Path(nombre_archivo).stem)
    return "soc" in nombre and aamm in nombre


def buscar_soc(medidas_dir, aamm):
    """
    Busca dentro de Medidas/ el archivo de SoC del periodo AAMM
    indicado por el usuario.

    Ninguno   -> error
    Uno       -> se usa
    Varios    -> error, NO se elige por fecha de modificacion
    """

    medidas_dir = Path(medidas_dir)
    aamm = validar_aamm(aamm)

    if not medidas_dir.is_dir():
        raise ErrorEntrada(
            f"No existe la carpeta {medidas_dir}"
        )

    candidatos = [
        archivo
        for archivo in medidas_dir.iterdir()
        if archivo.is_file()
        and not archivo.name.startswith("~$")
        and archivo.suffix.lower() == ".xlsx"
        and _es_archivo_de_soc(archivo.name, aamm)
    ]

    if not candidatos:
        raise ErrorEntrada(
            f"No se encontro ningun archivo de SoC del periodo {aamm} "
            f"en {medidas_dir}\n"
            f"El nombre debe contener 'SOC' y '{aamm}', por ejemplo "
            f"SOC_{aamm}.xlsx"
        )

    if len(candidatos) > 1:
        nombres = ", ".join(sorted(c.name for c in candidatos))
        raise ErrorEntrada(
            f"Hay {len(candidatos)} archivos de SoC del periodo {aamm} "
            f"en {medidas_dir}:\n"
            f"  {nombres}\n"
            f"Deja solo el del periodo que vas a procesar. "
            f"No se elige automaticamente para no tomar en "
            f"silencio el archivo equivocado."
        )

    return candidatos[0]


def _buscar_archivo_excel_mas_reciente(carpeta, patron, desde_inicio=False):
    """
    Busca en 'carpeta' el archivo Excel (EXTENSIONES_EXCEL) mas
    reciente por fecha de modificacion cuyo nombre (en minusculas)
    contenga 'patron' (o empiece con el, si desde_inicio=True).
    Ignora archivos temporales (~$). None si la carpeta no existe o no
    hay ningun candidato.

    Comun a *OfertasSSCC*, SSCC_Desempeño_* y 3_REMUNERACIÓN_SUBASTAS_
    E_ID_*: las tres macros originales, ante varios candidatos, SI
    eligen automaticamente el mas reciente (a diferencia del SoC).
    """

    carpeta = Path(carpeta)

    if not carpeta.is_dir():
        return None

    patron = patron.lower()

    def coincide(nombre):
        nombre = nombre.lower()
        return nombre.startswith(patron) if desde_inicio else patron in nombre

    candidatos = [
        archivo
        for archivo in carpeta.iterdir()
        if archivo.is_file()
        and not archivo.name.startswith("~$")
        and archivo.suffix.lower() in EXTENSIONES_EXCEL
        and coincide(archivo.stem)
    ]

    if not candidatos:
        return None

    return max(candidatos, key=lambda a: a.stat().st_mtime)


def buscar_archivo_ofertas(ofertas_dir):
    """
    Replica OSSCC_BuscarArchivoOfertas: busca dentro de Ofertas/
    cualquier archivo Excel cuyo nombre contenga "OfertasSSCC".
    """

    return _buscar_archivo_excel_mas_reciente(
        ofertas_dir, PATRON_NOMBRE_OFERTAS
    )


def buscar_archivo_sscc_desempeno(carpeta):
    """
    Replica BuscarArchivoSSCCMasReciente: busca dentro de
    SSCC_Desempeño/ el archivo Excel mas reciente cuyo nombre empiece
    con "SSCC_Desempeño_".
    """

    return _buscar_archivo_excel_mas_reciente(
        carpeta, PATRON_NOMBRE_SSCC_DESEMPENO, desde_inicio=True
    )


def periodo_desde_aamm(aamm):
    """'2607' -> (2026, 7)"""

    anio = 2000 + int(aamm[:2])
    mes = int(aamm[2:])

    if not 1 <= mes <= 12:
        raise ErrorEntrada(
            f"El periodo '{aamm}' no tiene un mes valido."
        )

    return anio, mes
