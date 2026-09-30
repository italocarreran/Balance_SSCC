# -*- coding: utf-8 -*-
"""
config.json: lo que no va en el codigo.

El archivo vive al lado de Balance_BESS.py y NO se versiona (esta en
.gitignore). Tiene dos clases de contenido, y la diferencia importa:

  - Secciones POR PC/USUARIO ("<host>_<usuario>"): la carpeta base y el
    periodo AAMM que dejo abierto cada uno, mas "carpetas_por_periodo"
    ({aamm: carpeta}, la carpeta que uso cada mes: de ahi sale que
    cambiar de mes cambie de carpeta sola, y que dos meses no puedan
    compartir una). Las escribe la ventana sola.
  - La seccion COMPARTIDA "claves_api": las claves de las dos APIs del
    Coordinador. El valor es el MISMO para todo el equipo -no depende de
    quien corra el programa-, pero como el archivo no se versiona, cada
    uno tiene que pegarlas una vez en su copia (ver config.ejemplo.json).

"claves_api" es un nombre reservado: una seccion de usuario nunca se
llama asi (las de usuario son "<host>_<usuario>").
"""

import json
from pathlib import Path


# El archivo esta en la raiz del repositorio, un nivel arriba de Script/.
RUTA_CONFIG = Path(__file__).resolve().parent.parent / "config.json"

SECCION_CLAVES = "claves_api"

# Los nombres de las dos claves dentro de "claves_api". Son distintas
# entre si: la de medidas.coordinador.cl no sirve para
# operacion.coordinador.cl ni al reves.
CLAVE_PRMTE = "prmte"
CLAVE_GENERACION_REAL = "generacion_real"

# Lo que trae config.ejemplo.json: si alguien copia el ejemplo y no
# reemplaza el valor, se trata como si no hubiera clave.
MARCAS_SIN_COMPLETAR = ("", "PEGAR_AQUI_LA_CLAVE")


class ErrorConfig(Exception):
    """Falta algo en config.json, o esta mal escrito."""


def leer_todo():
    """El archivo entero como dict. {} si no existe o no se puede leer."""

    if not RUTA_CONFIG.exists():
        return {}

    try:
        with open(RUTA_CONFIG, "r", encoding="utf-8") as archivo:
            contenido = json.load(archivo)
    except (json.JSONDecodeError, OSError):
        return {}

    return contenido if isinstance(contenido, dict) else {}


def escribir_todo(contenido):
    """Guarda el archivo entero. Si no se puede escribir, no explota."""

    try:
        with open(RUTA_CONFIG, "w", encoding="utf-8") as archivo:
            json.dump(contenido, archivo, ensure_ascii=False, indent=2)
    except OSError:
        return False

    return True


def seccion(nombre):
    """Una seccion del config como dict. {} si no esta."""

    valor = leer_todo().get(nombre, {})

    return valor if isinstance(valor, dict) else {}


def actualizar_seccion(nombre, datos):
    """Mezcla `datos` en una seccion, conservando el resto del archivo."""

    contenido = leer_todo()
    contenido.setdefault(nombre, {}).update(datos)

    return escribir_todo(contenido)


def clave_api(cual):
    """
    La clave de una de las dos APIs del Coordinador, de la seccion
    "claves_api" de config.json.

    Levanta ErrorConfig -con el formato exacto para pegar- si el
    archivo no esta, si no tiene la seccion, o si la clave quedo con el
    texto de ejemplo sin reemplazar.
    """

    valor = seccion(SECCION_CLAVES).get(cual, "")
    valor = valor.strip() if isinstance(valor, str) else ""

    if valor in MARCAS_SIN_COMPLETAR:
        raise ErrorConfig(
            f"Falta la clave '{cual}' de la API del Coordinador.\n"
            f"\n"
            f"Abri (o crea) este archivo:\n"
            f"    {RUTA_CONFIG}\n"
            f"\n"
            f"y deja adentro esto, con la clave real en vez del texto de "
            f"ejemplo (si el archivo ya existe, agrega solo el bloque "
            f"\"{SECCION_CLAVES}\" sin tocar el resto):\n"
            f"\n"
            f'    {{\n'
            f'      "{SECCION_CLAVES}": {{\n'
            f'        "{CLAVE_PRMTE}": "PEGAR_AQUI_LA_CLAVE",\n'
            f'        "{CLAVE_GENERACION_REAL}": "PEGAR_AQUI_LA_CLAVE"\n'
            f'      }}\n'
            f'    }}\n'
            f"\n"
            f"Las dos claves son DISTINTAS: '{CLAVE_PRMTE}' es la de "
            f"medidas.coordinador.cl (PRMTE) y '{CLAVE_GENERACION_REAL}' "
            f"la de operacion.coordinador.cl (Gen real). El valor es el "
            f"mismo para todo el equipo; config.json no se versiona, asi "
            f"que cada uno lo pega una vez en su copia."
        )

    return valor
