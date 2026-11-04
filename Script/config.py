# -*- coding: utf-8 -*-
"""
La configuracion del programa vive en __config__/, nunca en las
carpetas de trabajo.

Regla del usuario (2026-11-04): "ese tipo de documentos json deben
quedar en la carpeta __config__ > AAAA > MM mes, si es especifico de un
mes, no en las carpetas de trabajo". __config__ es la carpeta hermana de
la del programa, la MISMA que usan el Revisor y los comparadores (repo
abby, __comun__/salidas.py):

    <raiz>/
        Balance_SSCC/            <- este programa
        Revisor_Relq/ ...
        __config__/
            config.json          <- compartido con los otros programas
            tks.json             <- los tokens de las APIs
            2026/09 Septiembre/Balance BESS/...   <- lo de un mes

config.json es COMPARTIDO: lo de este programa va en su propia seccion,
"modo_bess" (y "modo_cra" para Balance_CRA.py), con un bloque por
equipo y usuario adentro, igual que "modo_15min" del Revisor:

    {"modo_bess": {"<equipo>_<usuario>": {"carpeta_base": ..., "aamm": ...}}}

Mismas reglas que __comun__/config.py de abby, porque el archivo es de
todos: solo se agregan o actualizan claves (nunca se borra lo ajeno), si
el archivo existe pero no se puede leer NO se escribe (mejor perder un
ajuste que el archivo entero), y se escribe a un .tmp y despues se
reemplaza (un corte no lo deja a medias).

tks.json guarda los tokens de las APIs del Coordinador, que son los
mismos para todo el equipo:

    {"coordinador": {"prmte": "...", "generacion_real": "..."}}

Antes todo estaba en un config.json al lado del programa (seccion por
usuario y "claves_api"): la primera vez se copia de ahi, y los tokens se
siguen leyendo de ahi si tks.json todavia no los tiene.

BESS_CARPETA_CONFIG (variable de entorno) cambia la carpeta: la usan las
pruebas para no tocar la configuracion real.
"""

import json
import os
import re
import unicodedata
from pathlib import Path

_RAIZ_PROGRAMA = Path(__file__).resolve().parent.parent

CARPETA_CONFIG = Path(
    os.environ.get("BESS_CARPETA_CONFIG") or (_RAIZ_PROGRAMA.parent / "__config__")
)
RUTA_CONFIG = CARPETA_CONFIG / "config.json"
RUTA_TOKENS = CARPETA_CONFIG / "tks.json"
# El config.json de antes, al lado del programa (solo se lee, para
# migrar). Con BESS_CARPETA_CONFIG no se mira (las pruebas).
RUTA_CONFIG_ANTIGUA = (
    CARPETA_CONFIG / "_sin_config_antiguo.json"
    if os.environ.get("BESS_CARPETA_CONFIG")
    else _RAIZ_PROGRAMA / "config.json"
)

MODO_BESS = "modo_bess"
MODO_CRA = "modo_cra"

# Dentro de __config__/AAAA/MM Mes/, lo de este programa.
CARPETA_DEL_PROGRAMA = "Balance BESS"

SECCION_TOKENS = "coordinador"
SECCION_CLAVES = "claves_api"   # donde estaban antes (config.json viejo)

# Los nombres de los dos tokens. Son distintos entre si: el de
# medidas.coordinador.cl no sirve para operacion.coordinador.cl.
CLAVE_PRMTE = "prmte"
CLAVE_GENERACION_REAL = "generacion_real"

# Lo que trae el ejemplo: si no se reemplaza, se trata como si no hubiera.
MARCAS_SIN_COMPLETAR = ("", "PEGAR_AQUI_LA_CLAVE")

MESES = (
    "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
    "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre",
)


class ErrorConfig(Exception):
    """Falta algo en la configuracion, o esta mal escrito."""


# ------------------------------------------------------------
# Leer y escribir un json compartido
# ------------------------------------------------------------

def _leer_json(ruta):
    """(contenido, se_puede_escribir). {} si no existe."""

    ruta = Path(ruta)
    if not ruta.exists():
        return {}, True
    try:
        with open(ruta, "r", encoding="utf-8") as archivo:
            contenido = json.load(archivo)
    except (json.JSONDecodeError, OSError, UnicodeDecodeError):
        return {}, False
    if not isinstance(contenido, dict):
        return {}, False
    return contenido, True


def _escribir_json(ruta, contenido):
    """Escritura atomica (.tmp + os.replace). False si no se pudo."""

    ruta = Path(ruta)
    try:
        ruta.parent.mkdir(parents=True, exist_ok=True)
        temporal = ruta.with_suffix(ruta.suffix + ".tmp")
        with open(temporal, "w", encoding="utf-8") as archivo:
            json.dump(contenido, archivo, ensure_ascii=False, indent=2)
            archivo.flush()
            os.fsync(archivo.fileno())
        os.replace(temporal, ruta)
    except OSError:
        return False
    return True


def _modificar(ruta, cambio):
    """Lee, aplica cambio(contenido) y escribe, sin pisar un archivo roto."""

    contenido, se_puede = _leer_json(ruta)
    if not se_puede:
        return False
    cambio(contenido)
    return _escribir_json(ruta, contenido)


def leer_todo():
    """El config.json compartido entero. {} si no existe o no se puede leer."""

    return _leer_json(RUTA_CONFIG)[0]


def escribir_todo(contenido):
    """Guarda el config.json entero (no se usa si hay otro camino)."""

    def cambio(todo):
        todo.clear()
        todo.update(contenido)

    return _modificar(RUTA_CONFIG, cambio)


def seccion(nombre):
    """Una seccion de primer nivel del config.json. {} si no esta."""

    valor = leer_todo().get(nombre, {})
    return valor if isinstance(valor, dict) else {}


def actualizar_seccion(nombre, datos):
    """Mezcla datos en una seccion de primer nivel, sin tocar el resto."""

    def cambio(todo):
        if not isinstance(todo.get(nombre), dict):
            todo[nombre] = {}
        todo[nombre].update(datos)

    return _modificar(RUTA_CONFIG, cambio)


# ------------------------------------------------------------
# El bloque de un equipo/usuario dentro de un modo
# ------------------------------------------------------------

def _seccion_antigua(clave):
    if not clave:
        return {}
    valor = _leer_json(RUTA_CONFIG_ANTIGUA)[0].get(clave, {})
    return valor if isinstance(valor, dict) else {}


def seccion_usuario(modo, usuario, clave_antigua=None):
    """
    config.json -> modo -> usuario, como dict. Si todavia no esta y el
    config.json de antes tenia el bloque de este usuario (clave_antigua),
    se copia la primera vez.
    """

    bloque = seccion(modo).get(usuario)
    if isinstance(bloque, dict):
        return bloque
    antiguo = _seccion_antigua(clave_antigua)
    if antiguo:
        actualizar_seccion_usuario(modo, usuario, antiguo)
    return dict(antiguo)


def actualizar_seccion_usuario(modo, usuario, datos):
    """Mezcla datos en config.json -> modo -> usuario, sin tocar lo demas."""

    def cambio(todo):
        if not isinstance(todo.get(modo), dict):
            todo[modo] = {}
        if not isinstance(todo[modo].get(usuario), dict):
            todo[modo][usuario] = {}
        todo[modo][usuario].update(datos)

    return _modificar(RUTA_CONFIG, cambio)


# ------------------------------------------------------------
# Lo de un mes: __config__/AAAA/MM Mes/Balance BESS/
# ------------------------------------------------------------

def _normalizar(texto):
    texto = unicodedata.normalize("NFKD", str(texto or ""))
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    return " ".join(texto.split()).upper()


def carpeta_del_mes(aamm):
    """
    __config__/AAAA/MM Mes/Balance BESS, como arma las carpetas de mes el
    repo abby (__comun__/salidas.carpeta_mes): si el mes ya esta escrito
    de otra forma ("9 Septiembre") se usa esa, para no partir el mes en
    dos. Sin un AAMM valido: __config__/sin_mes/Balance BESS. No crea nada.
    """

    texto = str(aamm or "").strip()
    if not re.fullmatch(r"\d{4}", texto) or not 1 <= int(texto[2:]) <= 12:
        return CARPETA_CONFIG / "sin_mes" / CARPETA_DEL_PROGRAMA
    anio, mes = "20" + texto[:2], int(texto[2:])
    canonica = CARPETA_CONFIG / anio / f"{mes:02d} {MESES[mes - 1]}"
    if not canonica.exists() and (CARPETA_CONFIG / anio).is_dir():
        for otra in sorted((CARPETA_CONFIG / anio).iterdir()):
            partes = re.fullmatch(r"(\d{1,2})\s+(.+)", otra.name.strip())
            if (
                otra.is_dir() and partes and int(partes.group(1)) == mes
                and _normalizar(partes.group(2)) == _normalizar(MESES[mes - 1])
            ):
                canonica = otra
                break
    return canonica / CARPETA_DEL_PROGRAMA


# ------------------------------------------------------------
# Los tokens de las APIs: __config__/tks.json
# ------------------------------------------------------------

def _limpio(valor):
    valor = valor.strip() if isinstance(valor, str) else ""
    return "" if valor in MARCAS_SIN_COMPLETAR else valor


def clave_api(cual):
    """
    El token de una de las dos APIs del Coordinador: de tks.json, o si
    todavia no esta ahi, de "claves_api" del config.json de antes.

    Levanta ErrorConfig -con el formato exacto para pegar- si no esta.
    """

    tokens = _leer_json(RUTA_TOKENS)[0].get(SECCION_TOKENS, {})
    valor = _limpio(tokens.get(cual) if isinstance(tokens, dict) else "")
    if not valor:
        for ruta in (RUTA_CONFIG, RUTA_CONFIG_ANTIGUA):
            viejas = _leer_json(ruta)[0].get(SECCION_CLAVES, {})
            valor = _limpio(viejas.get(cual) if isinstance(viejas, dict) else "")
            if valor:
                break

    if not valor:
        raise ErrorConfig(
            f"Falta el token '{cual}' de la API del Coordinador.\n"
            f"\n"
            f"Abri (o crea) este archivo:\n"
            f"    {RUTA_TOKENS}\n"
            f"\n"
            f"y deja adentro esto, con los tokens reales en vez del texto de "
            f"ejemplo (si el archivo ya existe, agrega solo lo que falta):\n"
            f"\n"
            f'    {{\n'
            f'      "{SECCION_TOKENS}": {{\n'
            f'        "{CLAVE_PRMTE}": "PEGAR_AQUI_LA_CLAVE",\n'
            f'        "{CLAVE_GENERACION_REAL}": "PEGAR_AQUI_LA_CLAVE"\n'
            f'      }}\n'
            f'    }}\n'
            f"\n"
            f"Los dos son DISTINTOS: '{CLAVE_PRMTE}' es el de "
            f"medidas.coordinador.cl (PRMTE) y '{CLAVE_GENERACION_REAL}' el "
            f"de operacion.coordinador.cl (Gen real). Son los mismos para "
            f"todo el equipo."
        )

    return valor


def migrar_tokens():
    """
    Copia a tks.json los tokens que todavia esten solo en el config.json
    de antes ("claves_api"). No pisa ninguno que ya este. True si copio.
    """

    viejas = {}
    for ruta in (RUTA_CONFIG_ANTIGUA, RUTA_CONFIG):
        seccion_vieja = _leer_json(ruta)[0].get(SECCION_CLAVES, {})
        if isinstance(seccion_vieja, dict):
            for nombre, valor in seccion_vieja.items():
                if _limpio(valor) and nombre not in viejas:
                    viejas[nombre] = _limpio(valor)
    actuales = _leer_json(RUTA_TOKENS)[0].get(SECCION_TOKENS, {})
    actuales = actuales if isinstance(actuales, dict) else {}
    faltan = {k: v for k, v in viejas.items() if not _limpio(actuales.get(k))}
    if not faltan:
        return False

    def cambio(todo):
        if not isinstance(todo.get(SECCION_TOKENS), dict):
            todo[SECCION_TOKENS] = {}
        todo[SECCION_TOKENS].update(faltan)

    return _modificar(RUTA_TOKENS, cambio)
