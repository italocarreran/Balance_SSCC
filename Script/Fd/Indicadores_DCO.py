# -*- coding: utf-8 -*-
"""
Indicadores_DCO — trae el FD del periodo desde el arbol del DCO.

Las dos piezas salen de entradas_sscc.py y su
archivo_de_configuracion.yaml, con la raiz actualizada a la que usa hoy
el usuario (unidad F:, ver RAIZ_DCO_INDICADORES):

  - `ruta_fma_dco`, la raiz:
        F:\\11 SSCC\\05 Verificación SSCC\\02 Cálculo indicadores
  - el nombre del archivo de FD:
        SSCC_Disponibilidad_CSF_<Mes>_<AAAA>_<version>.zip
    con `version_fd` = V1 para el Preliminar y V2 para el Definitivo.

y la ruta completa la muestra el propio script en el comentario de la
rutina de FMA CPF:

    ...\\02 Cálculo indicadores\\2021\\12. Diciembre
        \\Indicadores Publicar\\V1\\01 Respuesta\\01 Indices CPF\\...

o sea, la carpeta del periodo es

    <RAIZ>\\<AAAA>\\<MM>. <Mes>\\Indicadores Publicar\\<version>

y adentro de esa, el FD (los "factores de desempeño") sale de
'03 Desempeño para publicar' -ruta confirmada por el usuario- y de
NINGUNA OTRA: ni de la carpeta donde vivia antes
('04 Desempeño para transferencias') ni de otras ramas del arbol de la
version. El boton "Traer FD" copia a <CARPETA_BASE>/FD y FMA/ los
archivos del periodo que haya SUELTOS ahi; si lo que encuentra es el
.zip, lo descomprime ahi mismo, porque lo que la etapa FD lee despues
es el Excel SSCC_Desempeño_* que viene adentro.

Nada de esto se adivina en cuanto a NOMBRES, pero si hay una decision
tomada por nosotros: **cual version usar**. La ventana no tiene un
selector Pre/Def, asi que por omision se toma la version MAS ALTA que
TENGA el archivo en esa carpeta (V2 y, si ahi no esta, V1) y se deja
dicho en el log. Se puede forzar pasando `version`. La ventana muestra
como "Origen: DCO" esa misma carpeta -la que se uso de verdad-, ver
ruta_origen_fd().

No importa nada de nucleo (solo la biblioteca estandar), igual que
los demas modulos de Script/. Los errores previsibles salen como
ErrorFd.
"""

import re
import shutil
import unicodedata
import zipfile
from pathlib import Path


# ============================================================
# DONDE PUBLICA EL DCO
# ============================================================

# `ruta_fma_dco` del archivo_de_configuracion.yaml, con la raiz que usa
# hoy el usuario: la unidad F: (antes se escribia el UNC
# \\nas-cen1\DCO\..., que es el mismo arbol montado de otra forma). Es
# la tercera (y ultima) ruta del programa que apunta fuera de la carpeta
# base del caso: si el servidor cambia, se cambia aca y nada mas.
RAIZ_DCO_INDICADORES = (
    r"F:\11 SSCC\05 Verificación SSCC\02 Cálculo indicadores"
)

CARPETA_PUBLICACION = "Indicadores Publicar"

# Dentro de la carpeta de version, el FD (los "factores de desempeño")
# sale de ESTA subcarpeta y de ninguna otra -ruta confirmada por el
# usuario-:
#
#   <version>/03 Desempeño para publicar/
#
# Se mira SOLO ahi y SIN entrar en sus subcarpetas. Antes, si no
# encontraba, se caia a una busqueda recursiva desde la carpeta de
# version: eso es lo que hacia que "Traer FD" copiara un monton de
# archivos de otras ramas del arbol del DCO. La regla ahora es la que
# pidio el usuario: esa carpeta en V2 y, si ahi no esta, la misma
# carpeta en V1. Nada mas.
SUBCARPETAS_FD = ("03 Desempeño para publicar",)

# Donde vivia el FD antes. YA NO SE BUSCA AHI: solo se nombra en el
# mensaje de error, para que se entienda que pasa si se abre un mes
# viejo que todavia lo tiene en la carpeta anterior.
SUBCARPETAS_FD_ANTIGUA = ("04 Desempeño para transferencias",)

# La version que se muestra en la ventana cuando todavia no se puede
# leer el servidor (no hay como saber cual esta publicada): la ruta que
# se arma ahi es orientativa, para poder abrir el arbol y mirar.
VERSION_POR_OMISION = "V1"

# Los dos nombres con los que puede aparecer el FD: el zip que publica
# el DCO y el Excel que viene adentro (que es el que lee la etapa FD).
PREFIJOS_FD = ("SSCC_Disponibilidad_CSF", "SSCC_Desempeño")

EXTENSIONES_FD = (".zip", ".xlsx", ".xlsm", ".xlsb", ".xls")

MESES = (
    "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
    "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre",
)

_PATRON_AAMM = re.compile(r"^\d{4}$")
_PATRON_VERSION = re.compile(r"^v(\d+)$", re.IGNORECASE)


class ErrorFd(Exception):
    """Error previsible al traer el FD."""


def _normalizar(texto):
    if texto is None:
        return ""

    texto = unicodedata.normalize("NFKD", str(texto))
    texto = "".join(c for c in texto if not unicodedata.combining(c))

    return re.sub(r"\s+", " ", texto).strip().lower()


def _validar_aamm(aamm):
    aamm = str(aamm or "").strip()

    if not _PATRON_AAMM.match(aamm):
        raise ErrorFd(
            f"Periodo invalido: '{aamm}'. Tienen que ser 4 digitos "
            f"(por ejemplo 2603 para marzo de 2026)."
        )

    return aamm


def periodo_desde_aamm(aamm):
    """'2603' -> (2026, 3)."""

    aamm = _validar_aamm(aamm)
    anio = 2000 + int(aamm[:2])
    mes = int(aamm[2:])

    if not 1 <= mes <= 12:
        raise ErrorFd(f"El periodo '{aamm}' no tiene un mes valido.")

    return anio, mes


def nombre_mes(mes):
    """3 -> 'Marzo'."""

    return MESES[int(mes) - 1]


def _subcarpeta(padre, *alternativas):
    """
    Busca una subcarpeta comparando por nombre normalizado (sin tildes,
    sin importar mayusculas ni espacios de mas). Devuelve None si no
    esta. Se hace asi y no con una ruta literal porque estas carpetas
    las escribe una persona todos los meses: "03. Marzo" y "3. Marzo"
    son la misma para nosotros.
    """

    if not padre.is_dir():
        return None

    buscadas = [_normalizar(a) for a in alternativas]

    try:
        for hijo in padre.iterdir():
            if hijo.is_dir() and _normalizar(hijo.name) in buscadas:
                return hijo
    except OSError:
        return None

    return None


def bajar_por_subcarpetas(carpeta, subcarpetas):
    """
    Baja por los nombres dados uno tras otro, comparando por nombre
    normalizado. Devuelve None si en algun escalon no esta.
    """

    actual = Path(carpeta)

    for nombre in subcarpetas:
        actual = _subcarpeta(actual, nombre)
        if actual is None:
            return None

    return actual


def carpeta_del_periodo(aamm, raiz=None):
    """
    <RAIZ>/<AAAA>/<MM>. <Mes>/Indicadores Publicar

    Levanta ErrorFd diciendo exactamente que tramo de la ruta falta:
    es el error que de verdad se va a ver (una carpeta del mes todavia
    sin publicar, o el servidor sin conectar).
    """

    anio, mes = periodo_desde_aamm(aamm)
    raiz = Path(raiz or RAIZ_DCO_INDICADORES)

    if not raiz.is_dir():
        raise ErrorFd(
            f"No se puede leer la carpeta de indicadores del DCO:\n"
            f"  {raiz}\n\n"
            f"Revisa que tengas conexion y permiso sobre ese servidor. "
            f"Si la ruta cambio, se cambia en RAIZ_DCO_INDICADORES "
            f"(Script/Fd/Indicadores_DCO.py)."
        )

    carpeta_anio = _subcarpeta(raiz, str(anio))

    if carpeta_anio is None:
        raise ErrorFd(f"No existe la carpeta del año {anio} en {raiz}")

    carpeta_mes = _subcarpeta(
        carpeta_anio,
        f"{mes:02d}. {nombre_mes(mes)}",
        f"{mes}. {nombre_mes(mes)}",
        nombre_mes(mes),
    )

    if carpeta_mes is None:
        raise ErrorFd(
            f"No existe la carpeta de {nombre_mes(mes)} ({mes:02d}) en "
            f"{carpeta_anio}"
        )

    carpeta_publicacion = _subcarpeta(carpeta_mes, CARPETA_PUBLICACION)

    if carpeta_publicacion is None:
        raise ErrorFd(
            f"No existe '{CARPETA_PUBLICACION}' en {carpeta_mes}: el DCO "
            f"todavia no publico los indicadores de ese mes."
        )

    return carpeta_publicacion


def versiones_publicadas(carpeta_publicacion):
    """
    Las carpetas de version que hay publicadas (V1, V2, ...), ordenadas
    de menor a mayor por su numero.
    """

    versiones = []

    try:
        for hijo in Path(carpeta_publicacion).iterdir():
            if hijo.is_dir() and _PATRON_VERSION.match(hijo.name.strip()):
                versiones.append(hijo)
    except OSError:
        return []

    return sorted(
        versiones,
        key=lambda h: int(_PATRON_VERSION.match(h.name.strip()).group(1)),
    )


def elegir_version(carpeta_publicacion, version=None):
    """
    Devuelve la carpeta de version a usar. Si no se pide una en
    particular, la MAS ALTA publicada: V2 (Definitivo) le gana a V1
    (Preliminar), que es lo que se quiere cuando ya salio la
    definitiva.
    """

    versiones = versiones_publicadas(carpeta_publicacion)

    if not versiones:
        raise ErrorFd(
            f"No hay ninguna carpeta de version (V1, V2, ...) en "
            f"{carpeta_publicacion}"
        )

    if version is None:
        return versiones[-1]

    buscada = _normalizar(version)

    for carpeta in versiones:
        if _normalizar(carpeta.name) == buscada:
            return carpeta

    raise ErrorFd(
        f"No existe la version '{version}' en {carpeta_publicacion}. "
        f"Publicadas: {', '.join(c.name for c in versiones)}"
    )


def buscar_en_versiones(
    carpeta_publicacion, buscar, version=None, registrar=print
):
    """
    Recorre las versiones publicadas **de mayor a menor** y devuelve la
    primera en la que `buscar(carpeta_version)` encuentre algo.

    No alcanza con quedarse con la carpeta de version mas alta: puede
    existir la carpeta V2 y NO tener adentro el archivo que se busca
    (recien publicada, a medio subir, o esa version no incluye esa
    entrega). En ese caso hay que caer a V1. Esto es exactamente lo que
    pidio el usuario: "debe ser el mayor que tenga disponible el
    archivo".

    buscar: funcion carpeta_version -> resultado (lo que sea, falsy si
    no encontro).
    version: para forzar una en particular; entonces se prueba solo esa.

    Devuelve (carpeta_version, resultado, revisadas), con
    carpeta_version y resultado en None si ninguna la tenia.
    `revisadas` son los nombres de las versiones que se miraron, para
    poder decirlo en el error.
    """

    versiones = versiones_publicadas(carpeta_publicacion)

    if not versiones:
        raise ErrorFd(
            f"No hay ninguna carpeta de version (V1, V2, ...) en "
            f"{carpeta_publicacion}"
        )

    if version is not None:
        candidatas = [elegir_version(carpeta_publicacion, version)]
    else:
        candidatas = list(reversed(versiones))

    revisadas = []

    for carpeta in candidatas:

        resultado = buscar(carpeta)

        if resultado:

            if revisadas:
                registrar(
                    f"  se usa {carpeta.name}: en "
                    f"{', '.join(revisadas)} no estaba el archivo"
                )
            elif version is None:
                registrar(
                    f"  version mas alta con el archivo: {carpeta.name} "
                    f"(V1 = Preliminar, V2 = Definitivo)"
                )

            return carpeta, resultado, revisadas

        revisadas.append(carpeta.name)

    return None, None, revisadas


# ============================================================
# LA RUTA DE ORIGEN, PARA MOSTRARLA EN LA VENTANA
#
# La ventana pone "Origen: DCO" en el detalle de la fila del FD, y ese
# "DCO" es un link a la carpeta exacta de la que sale. Por eso estas
# funciones NO levantan ErrorFd: cuando el servidor no esta conectado o
# el mes todavia no se publico igual hay que poder mostrar (y abrir) la
# ruta que le corresponderia al periodo.
# ============================================================

def carpeta_publicacion_o_literal(aamm, raiz=None):
    """
    <RAIZ>/<AAAA>/<MM>. <Mes>/Indicadores Publicar, resuelta contra el
    servidor si se puede leer y armada a mano si no.
    """

    raiz = Path(raiz or RAIZ_DCO_INDICADORES)

    try:
        return carpeta_del_periodo(aamm, raiz=raiz)
    except ErrorFd:
        pass

    try:
        anio, mes = periodo_desde_aamm(aamm)
    except ErrorFd:
        return raiz

    return (
        raiz / str(anio) / f"{mes:02d}. {nombre_mes(mes)}"
        / CARPETA_PUBLICACION
    )


def carpeta_version_usada(aamm, buscar=None, version=None, raiz=None):
    """
    La carpeta de version (V1, V2, ...) que de verdad se va a usar.

    buscar: la misma funcion que usa el boton para decidir
    (carpeta_version -> lo que encontro, falsy si nada). Con eso, la
    version que devuelve es la que TIENE el archivo, no la mas alta
    que exista: si V2 esta publicada pero vacia, el boton baja a V1 y
    el link de la ventana tiene que apuntar a V1 tambien.

    Sin `buscar`, o si ninguna version tiene nada, la mas alta
    publicada. Si no se puede leer el servidor, la pedida o
    VERSION_POR_OMISION, para poder mostrar igual una ruta.
    """

    publicacion = carpeta_publicacion_o_literal(aamm, raiz=raiz)

    if buscar is not None:
        try:
            carpeta, _, _ = buscar_en_versiones(
                publicacion, buscar, version=version,
                registrar=lambda *_: None,
            )
        except ErrorFd:
            carpeta = None

        if carpeta is not None:
            return carpeta

    try:
        return elegir_version(publicacion, version)
    except ErrorFd:
        return publicacion / (version or VERSION_POR_OMISION)


def ruta_origen(aamm, subcarpetas=None, version=None, raiz=None,
                buscar=None):
    """
    La carpeta exacta del arbol del DCO de la que sale algo del
    periodo, para mostrarla y abrirla desde la ventana.

    Es la MISMA que usa el boton: la version la elige `buscar` (ver
    carpeta_version_usada) y, dentro de ella, la cadena `subcarpetas`.
    Si esa carpeta no existe todavia, se devuelve armada a mano: la
    ventana tiene que poder abrir el arbol igual.
    """

    try:
        periodo_desde_aamm(aamm)
    except ErrorFd:
        # Sin un periodo valido no hay carpeta de año, mes ni version
        # que armar: lo unico cierto es la raiz.
        return Path(raiz or RAIZ_DCO_INDICADORES)

    base = carpeta_version_usada(
        aamm, buscar=buscar, version=version, raiz=raiz
    )
    subcarpetas = tuple(subcarpetas or SUBCARPETAS_FD)

    carpeta = bajar_por_subcarpetas(base, subcarpetas)

    if carpeta is not None:
        return carpeta

    for nombre in subcarpetas:
        base = base / nombre

    return base


def ruta_origen_fd(aamm, version=None, raiz=None):
    """
    De donde sale el FD del periodo:

        <RAIZ>/<AAAA>/<MM>. <Mes>/Indicadores Publicar/<Vn>
            /03 Desempeño para publicar

    con el <Vn> que de verdad tiene el archivo (V2 y, si ahi no esta,
    V1). Es la ruta que la ventana muestra como "Origen: DCO" en la
    fila del SSCC_Desempeño_* y la que abre al hacerle click.
    """

    try:
        anio, _ = periodo_desde_aamm(aamm)
    except ErrorFd:
        return Path(raiz or RAIZ_DCO_INDICADORES)

    return ruta_origen(
        aamm, SUBCARPETAS_FD, version=version, raiz=raiz,
        buscar=lambda carpeta: buscar_archivos_fd(carpeta, anio),
    )


def _es_archivo_fd(ruta, anio):
    """
    Un archivo sirve si su nombre empieza con alguno de los prefijos de
    FD y menciona el año del periodo (asi no se cuelan los de otro mes
    que a veces quedan sueltos en la misma carpeta).
    """

    if ruta.suffix.lower() not in EXTENSIONES_FD:
        return False

    nombre = _normalizar(ruta.stem)

    if not any(nombre.startswith(_normalizar(p)) for p in PREFIJOS_FD):
        return False

    return str(anio) in ruta.stem


def carpeta_fd_de_la_version(carpeta_version):
    """
    <version>/03 Desempeño para publicar, o None si esa version no la
    tiene. Es la UNICA carpeta de la que sale el FD.
    """

    return bajar_por_subcarpetas(Path(carpeta_version), SUBCARPETAS_FD)


def buscar_archivos_fd(carpeta_version, anio):
    """
    Los archivos de FD del periodo que hay en
    '<version>/03 Desempeño para publicar'. Solo ahi y solo los de esa
    carpeta (no se entra en sus subcarpetas): lo que no esta ahi, no es
    el FD de este periodo. Devuelve la lista ordenada por nombre, vacia
    si la carpeta no existe o no tiene ninguno.
    """

    carpeta_fd = carpeta_fd_de_la_version(carpeta_version)

    if carpeta_fd is None:
        return []

    return _archivos_fd_en(carpeta_fd, anio)


def _archivos_fd_en(carpeta, anio, recursivo=False):
    """Los archivos de FD del año que haya en una carpeta."""

    carpeta = Path(carpeta)
    encontrados = []

    try:
        rutas = carpeta.rglob("*") if recursivo else carpeta.iterdir()
        for ruta in rutas:
            if ruta.is_file() and _es_archivo_fd(ruta, anio):
                encontrados.append(ruta)
    except OSError as error:
        raise ErrorFd(f"No se pudo recorrer {carpeta}: {error}") from error

    return sorted(encontrados, key=lambda r: r.name)


def _contenido_de_la_carpeta_fd(carpeta_publicacion):
    """
    Que hay en '03 Desempeño para publicar' de cada version. Se usa
    SOLO para el mensaje de error: si no se encontro el archivo, lo
    util es ver que SI hay en la carpeta de la que se saca.

    Si una version no tiene esa carpeta pero si la de antes
    ('04 Desempeño para transferencias', donde vivia el FD hasta hace
    poco), se dice: ahi no se busca mas, y es la explicacion de por que
    un mes viejo no trae nada.
    """

    lineas = []

    for carpeta_version in reversed(versiones_publicadas(carpeta_publicacion)):

        carpeta_fd = carpeta_fd_de_la_version(carpeta_version)

        if carpeta_fd is None:

            antigua = bajar_por_subcarpetas(
                carpeta_version, SUBCARPETAS_FD_ANTIGUA
            )

            lineas.append(
                f"  {carpeta_version.name}: no tiene "
                f"'{SUBCARPETAS_FD[0]}'"
                + (
                    f" (si tiene '{SUBCARPETAS_FD_ANTIGUA[0]}', la de "
                    f"antes, donde ya no se busca)"
                    if antigua is not None else ""
                )
            )
            continue

        try:
            nombres = sorted(r.name for r in carpeta_fd.iterdir() if r.is_file())
        except OSError:
            nombres = []

        lineas.append(
            f"  {carpeta_version.name}: "
            + (", ".join(nombres) if nombres else "(vacia)")
        )

    return lineas


def _extraer_zip(ruta_zip, destino, registrar=print):
    """
    Descomprime el zip del FD en la carpeta destino. Solo saca los
    Excel que trae adentro y los deja SUELTOS en la carpeta (sin
    recrear el arbol del zip), que es donde la etapa FD los busca.

    Ignora cualquier entrada con una ruta rara (absoluta o con ".."):
    un zip no tiene por que poder escribir fuera de la carpeta a la que
    se lo descomprime.
    """

    extraidos = []

    try:
        with zipfile.ZipFile(ruta_zip) as zip_fd:

            for miembro in zip_fd.infolist():

                if miembro.is_dir():
                    continue

                nombre = Path(miembro.filename).name

                if not nombre or nombre.startswith("."):
                    continue

                if Path(miembro.filename).is_absolute() or ".." in Path(
                    miembro.filename
                ).parts:
                    registrar(f"  [AVISO] entrada ignorada del zip: {miembro.filename}")
                    continue

                if Path(nombre).suffix.lower() not in EXTENSIONES_FD:
                    continue

                destino_archivo = Path(destino) / nombre

                with zip_fd.open(miembro) as origen, open(
                    destino_archivo, "wb"
                ) as salida:
                    shutil.copyfileobj(origen, salida)

                extraidos.append(nombre)

    except (zipfile.BadZipFile, OSError) as error:
        raise ErrorFd(
            f"No se pudo descomprimir {Path(ruta_zip).name}: {error}"
        ) from error

    return extraidos


def traer_fd(carpeta_destino, aamm, version=None, raiz=None, registrar=print):
    """
    Copia a <CARPETA_BASE>/FD y FMA/ el FD del periodo publicado por el
    DCO, y descomprime el zip si lo que vino es un zip.

    Devuelve (copiados, extraidos, carpeta_version_usada).
    """

    aamm = _validar_aamm(aamm)
    anio, _ = periodo_desde_aamm(aamm)

    carpeta_destino = Path(carpeta_destino)

    if not carpeta_destino.is_dir():
        raise ErrorFd(f"No se encontro la carpeta {carpeta_destino}")

    carpeta_publicacion = carpeta_del_periodo(aamm, raiz=raiz)

    carpeta_version, archivos, revisadas = buscar_en_versiones(
        carpeta_publicacion,
        lambda carpeta: buscar_archivos_fd(carpeta, anio),
        version=version,
        registrar=registrar,
    )

    if not archivos:
        detalle = _contenido_de_la_carpeta_fd(carpeta_publicacion)

        raise ErrorFd(
            f"No se encontro ningun archivo de FD "
            f"({' / '.join(p + '*' for p in PREFIJOS_FD)}) del año {anio} "
            f"en {carpeta_publicacion}.\n\n"
            f"Version(es) revisada(s): {', '.join(revisadas)}\n\n"
            f"Lo que hay en la carpeta de desempeño de cada version:\n"
            + "\n".join(detalle)
        )

    copiados, extraidos = [], []

    for ruta in archivos:

        destino_archivo = carpeta_destino / ruta.name

        registrar(f"  copiando {ruta.name}...")

        try:
            shutil.copy2(ruta, destino_archivo)
        except OSError as error:
            raise ErrorFd(f"No se pudo copiar {ruta.name}: {error}") from error

        copiados.append(ruta.name)

        if destino_archivo.suffix.lower() == ".zip":
            nuevos = _extraer_zip(destino_archivo, carpeta_destino, registrar)
            extraidos.extend(nuevos)
            registrar(
                f"  descomprimido {ruta.name}: "
                f"{', '.join(nuevos) if nuevos else 'sin Excel adentro'}"
            )

    return copiados, extraidos, carpeta_version
