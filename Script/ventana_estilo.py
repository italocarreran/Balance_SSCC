# -*- coding: utf-8 -*-
"""
El aspecto comun de las ventanas Balance_BESS.py y Balance_CRA.py:
colores, simbolos de estado, anchos del diagrama, rueda del mouse,
colores del registro y como se abre una carpeta desde una fila.

Pedido del usuario (2026-11-16): "la ventana del CRA dejala en el
formato del BESS". Esto salio tal cual de Balance_BESS.py para que las
dos lo tomen del mismo lugar: si se cambia aca, cambian las dos.

Liviano a proposito (solo stdlib): Balance_BESS.py lo importa antes de
mostrar la ventana (ver desarrollo/tests/test_arranque_liviano.py).
"""

import os
import re
import subprocess
import sys
from pathlib import Path


COLOR_OK = "#1a6b1a"
COLOR_FALTA = "#b00020"
COLOR_PENDIENTE = "#a06000"
COLOR_NEUTRO = "#555555"
COLOR_BASE_OK = "#1a4fb0"
COLOR_LINK = "#1a4fb0"

SIMBOLO = {
    "ok": "OK",
    "falta": "FALTA",
    "pendiente": "PENDIENTE",
}

# El simbolo que la fila de Medidas_SAE.xlsx muestra cuando el archivo
# se genero con informacion incompleta (ver nucleo.aviso_puntos_fallidos).
# Va sin el selector de emoji (U+FE0F): Tk en Windows lo dibuja como un
# cuadrito.
SIMBOLO_AVISO = "\u26a0"
COLOR_AVISO = "#d97706"

# Colores del Registro (pedido del usuario): verde lo que termino bien,
# amarillo los avisos, rojo los errores. El amarillo es oscuro (ambar)
# porque un amarillo puro no se lee sobre el fondo blanco.
COLOR_REGISTRO = {
    "ok": "#1a7f1a",
    "aviso": "#b8860b",
    "error": "#c00020",
}

# Donde queda el registro de la ultima corrida, en la carpeta Log/ del
# caso (para poder mandarlo sin copiar a mano): nucleo.ARCHIVO_REGISTRO,
# que tambien se lee como Balance_BESS.ARCHIVO_REGISTRO (ver __getattr__).

_PATRON_ERROR = re.compile(
    r"\bERROR\b|Traceback|\bNO se corre\b|: error\b|termino con problemas",
    re.IGNORECASE,
)
_PATRON_AVISO = re.compile(
    r"\[AVISO\]|\bAVISO\b|\bOJO\b|advertencia|desactualizad|omitido|"
    r"queda fuera|no se encontro|sin datos|incomplet|cambio de hora",
    re.IGNORECASE,
)
_PATRON_OK = re.compile(
    r"^(\[[^\]]+\]\s*)?(Listo|Terminado|Hecho|OK\b)|: ok\b|correctamente",
    re.IGNORECASE,
)


def nivel_del_mensaje(texto):
    """'error' / 'aviso' / 'ok' / None: que color lleva en el Registro."""

    texto = str(texto).strip()
    if not texto:
        return None
    if _PATRON_ERROR.search(texto):
        return "error"
    if _PATRON_AVISO.search(texto):
        return "aviso"
    if _PATRON_OK.search(texto):
        return "ok"
    return None


def formatear_reloj(segundos):
    """83.4 -> '01:23'; 3725 -> '1:02:05'."""

    entero = int(max(0, segundos))
    horas, resto = divmod(entero, 3600)
    minutos, seg = divmod(resto, 60)
    if horas:
        return f"{horas}:{minutos:02d}:{seg:02d}"
    return f"{minutos:02d}:{seg:02d}"

COLOR_ESTADO = {
    "ok": COLOR_OK,
    "falta": COLOR_FALTA,
    "pendiente": COLOR_PENDIENTE,
}

# Anchos de las columnas del diagrama, TODOS en pixeles y cada celda
# en un Frame de ancho fijo (pack_propagate(False)).
#
# Antes estaban en "caracteres" (width=N de un Label) y las columnas
# salian torcidas: una fila en negrita mide mas que la misma fila en
# redonda aunque las dos digan 52 caracteres, asi que el estado y el
# boton de las filas de nivel 0 quedaban corridos a la derecha
# respecto del resto. En pixeles no hay forma de que se corran: la
# celda mide lo mismo diga lo que diga adentro.
ANCHO_ESTRUCTURA = 420     # pixeles
ANCHO_ESTADO = 86          # pixeles
ANCHO_ACCION = 104         # pixeles
ALTO_ACCION = 26           # pixeles (alto de una fila CON boton)
ALTO_FILA = 18             # pixeles (alto de una fila sin boton)

# El boton ocupa su celda entera: todos los botones del diagrama miden
# exactamente lo mismo y quedan en una columna recta (pedido del
# usuario). Por eso los textos son de una palabra ("Traer",
# "Resumir", "Generar", "Actualizar", "Calcular").
RELLENO_ACCION = 4         # pixeles de aire alrededor del boton

# Rueda del mouse. La ventana entera se desplaza de a PIXELES, no de a
# "unidades" del Canvas: una unidad de Canvas sin yscrollincrement es
# un decimo del alto visible, y ese salto es lo que se sentia pegado.
# El registro, en cambio, se desplaza de a lineas, que es lo natural en
# un widget de texto.
PIXELES_POR_MUESCA = 45    # cuanto baja la ventana por cada muesca
LINEAS_POR_MUESCA = 3      # cuantas lineas baja el registro por muesca
# Windows/macOS mandan <MouseWheel>; X11 manda los botones 4 y 5.
EVENTOS_RUEDA = ("<MouseWheel>", "<Button-4>", "<Button-5>")

# Cuantos mensajes del hilo de trabajo se escriben en el registro por
# pasada (ver vaciar_cola_log): una rafaga grande no congela la ventana.
MENSAJES_POR_TANDA = 400


def carpeta_a_abrir(ruta, es_archivo=False):
    """
    Que carpeta se abre al hacer click en una fila del diagrama: la
    propia si la fila es una carpeta, la que CONTIENE al archivo si es
    un archivo -- nunca se abre el archivo, para no arrancar Excel sin
    que se lo pidan.

    Si todavia no existe (una carpeta que falta, un archivo que aun no
    se genero, el mes que el DCO no publico), se sube hasta el primer
    ancestro que si exista: asi el click siempre lleva a algun lado.
    Devuelve None si no existe ni la raiz (unidad desconectada).
    """

    carpeta = Path(ruta)

    if es_archivo:
        carpeta = carpeta.parent

    while True:

        try:
            if carpeta.is_dir():
                return carpeta
        except OSError:
            return None

        padre = carpeta.parent

        if padre == carpeta:
            return None

        carpeta = padre


def abrir_en_explorador(ruta, es_archivo=False):
    """
    Abre en el explorador la carpeta de esa ruta. Devuelve la carpeta
    que abrio, o None si no habia ninguna que abrir.
    """

    if not ruta:
        return None

    carpeta = carpeta_a_abrir(ruta, es_archivo)

    if carpeta is None:
        return None

    if sys.platform == "win32":
        subprocess.Popen(["explorer", str(carpeta)])
    elif sys.platform == "darwin":
        subprocess.Popen(["open", str(carpeta)])
    else:
        subprocess.Popen(["xdg-open", str(carpeta)])

    return carpeta


def abrir_archivo(ruta):
    """
    Abre ESE archivo con su programa (Excel para un .xlsx). Es la
    excepcion a "nunca se abre el archivo": solo se usa cuando el
    usuario lo pide con un boton (el "Ver detalle" del aviso de
    Medidas_SAE.xlsx). Devuelve False si el archivo no esta.
    """

    ruta = Path(ruta)

    if not ruta.is_file():
        return False

    if sys.platform == "win32":
        os.startfile(str(ruta))
    elif sys.platform == "darwin":
        subprocess.Popen(["open", str(ruta)])
    else:
        subprocess.Popen(["xdg-open", str(ruta)])

    return True


def formato_tiempo(segundos):
    minutos, seg = divmod(int(segundos), 60)
    horas, minutos = divmod(minutos, 60)
    return f"{horas:02d}:{minutos:02d}:{seg:02d}"
