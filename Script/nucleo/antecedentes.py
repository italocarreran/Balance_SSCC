# -*- coding: utf-8 -*-
"""
Antecedentes SAE: la carpeta "03 Antecedentes SAE" que va junto al
11_PAGOS_BESS_AAMM.xlsx entregado, con el codigo para recalcularlo.

Pedido del usuario (2026-11-10): "que se adjunte junto con el archivo 11_
y que al darle play se recalculen las hojas de calculo"; un solo codigo,
generado automatico, que abra una ventana "desde el punto de los
calculos en adelante". Si se corre desde esa carpeta toma solo el 11_ de
la carpeta de arriba; si no, se elige con Examinar.

escribir_recalculador() arma ese codigo (Calcular_11_PAGOS_BESS_AAMM.py): UN archivo .py con los modulos
del calculo adentro, tal cual estan en el programa (como texto que se
puede leer), y un cargador que los importa desde ahi. Va recortado: solo
lleva lo que usa el calculo desde las hojas de entrada del 11_ en
adelante (sin APIs, descargas, Access ni reportes del DCO). Como sale del
mismo codigo del programa al momento de generarlo, recalcula igual que el
calculo con que se hizo el 11_.

generar_antecedentes() crea la carpeta al lado del 11_ con ese codigo,
el SAE_AAMM.csv (por defecto) y los demas archivos del caso que se elijan.
"""

import ast
import datetime
import shutil
import subprocess
from pathlib import Path

from .rutas import nombre_csv_sae, resolver_rutas, validar_aamm
from .utiles import ErrorEntrada

CARPETA_ANTECEDENTES = "03 Antecedentes SAE"

_SCRIPT = Path(__file__).resolve().parent.parent      # .../Script
_NUCLEO = _SCRIPT / "nucleo"

# Desde donde se sigue el rastro de imports dentro de nucleo/.
_ENTRADAS_NUCLEO = ("recalculo", "proceso", "csv_sae")

# Lo de fuera de nucleo/ que el calculo usa (comprobado corriendo el
# calculo con un externos recortado: no toca nada mas).
_OTROS = (
    "__init__.py", "config.py", "dialogos.py", "ventana_recalculo.py",
    "Subastas/__init__.py", "Subastas/Fma.py",
    "Subastas/Ofertas_Adjudicadas.py",
    "Fd/__init__.py", "Fd/Desempeno_Horario.py",
)

# Primero lo que un revisor busca: el calculo de las hojas.
_PRIMERO = (
    "ecostos", "ecostos_columnas", "ecostos_prorratas", "ecostos_valorizacion",
    "columnas_compartidas", "re545", "re545_componentes", "re545_reservas",
    "re545_resumen", "compensacion", "prorrata_retiros", "csv_sae",
    "recalculo",
)

_INIT_NUCLEO = '''# -*- coding: utf-8 -*-
"""
El calculo del caso, un modulo por etapa. (En el programa este archivo
re-exporta todo; aca va vacio: cada modulo importa lo suyo.)
"""
'''

# externos recortado: solo los tres modulos de fuera que usa el calculo;
# lo demas (APIs de medidas, CMg, reportes del DCO) no viene incluido.
_EXTERNOS = '''# -*- coding: utf-8 -*-
"""
Los paquetes hermanos de nucleo/ que usa el calculo. Version recortada
para el recalculo: las descargas (APIs de medidas, CMg, reportes del
DCO) no vienen incluidas, porque el recalculo parte de las hojas de
entrada del 11_.
"""

from ..Subastas import Ofertas_Adjudicadas as ofertas_adj
from ..Subastas import Fma as fma_subastas
from ..Fd import Desempeno_Horario as desempeno_fd


class _NoIncluido:
    """Un modulo de descarga que no viene en el recalculo."""

    def __init__(self, nombre):
        self._nombre = nombre

    def __getattr__(self, atributo):
        raise RuntimeError(
            f"{self._nombre} no viene en el recalculo: es de la descarga de "
            f"datos, no del calculo."
        )


extrae_cmg = _NoIncluido("Extrae_CMG_barras")
Homologacion = _NoIncluido("Homologacion")
Descarga_PRMTE = _NoIncluido("Descarga_PRMTE")
Claves_Balance = _NoIncluido("Claves_Balance")
Generacion_Real = _NoIncluido("Generacion_Real")
indicadores_dco = _NoIncluido("Indicadores_DCO")
indices_fma = _NoIncluido("Indices_FMA")


class ErrorMedidas(Exception):
    """Error de la descarga de medidas (no viene en el recalculo)."""
'''

_CARGADOR = '''
# ============================================================
# Cargador: importa los modulos de FUENTES (mas abajo) como el paquete
# "calculo_sae". Es lo unico de este archivo que no es del programa.
# ============================================================

import importlib
import importlib.abc
import importlib.util
import linecache
import os
import sys
import tempfile
from pathlib import Path

PAQUETE = "calculo_sae"


def _instalar():
    modulos = {}
    for ruta, fuente in FUENTES.items():
        partes = ruta[:-3].split("/")
        es_paquete = partes[-1] == "__init__"
        if es_paquete:
            partes = partes[:-1]
        modulos[".".join([PAQUETE] + partes)] = (ruta, fuente, es_paquete)

    class Cargador(importlib.abc.MetaPathFinder, importlib.abc.Loader):
        def find_spec(self, nombre, path=None, target=None):
            if nombre in modulos:
                return importlib.util.spec_from_loader(
                    nombre, self, is_package=modulos[nombre][2]
                )
            return None

        def create_module(self, spec):
            return None

        def exec_module(self, modulo):
            ruta, fuente, _ = modulos[modulo.__name__]
            archivo = f"{Path(__file__).name}/{ruta}"
            # Para que un error diga la linea del modulo que fallo.
            linecache.cache[archivo] = (
                len(fuente), None, fuente.splitlines(True), archivo
            )
            modulo.__file__ = archivo
            exec(compile(fuente, archivo, "exec"), modulo.__dict__)

    sys.meta_path.insert(0, Cargador())


def main():
    # Lo que el calculo guarda aparte (sellos de las hojas) va a una
    # carpeta temporal: el recalculo no deja nada fuera de sus resultados.
    os.environ.setdefault(
        "BESS_CARPETA_CONFIG", tempfile.mkdtemp(prefix="recalculo_config_")
    )
    _instalar()
    ventana = importlib.import_module(PAQUETE + ".ventana_recalculo")
    ventana.ejecutar_desde(Path(__file__).resolve().parent)
'''


def _imports_nucleo(modulo):
    """Los modulos de nucleo/ que importa modulo (relativos, nivel 1)."""

    arbol = ast.parse((_NUCLEO / f"{modulo}.py").read_text(encoding="utf-8"))
    salida = set()
    for nodo in ast.walk(arbol):
        if isinstance(nodo, ast.ImportFrom) and nodo.level == 1:
            if nodo.module:
                salida.add(nodo.module.split(".")[0])
            else:
                salida.update(a.name for a in nodo.names)
    return {m for m in salida if (_NUCLEO / f"{m}.py").is_file()}


def modulos_del_calculo():
    """Los modulos de nucleo/ que necesita el recalculo (por sus imports)."""

    vistos, pendientes = set(), list(_ENTRADAS_NUCLEO)
    while pendientes:
        modulo = pendientes.pop()
        if modulo in vistos or modulo in ("__init__", "externos"):
            continue
        vistos.add(modulo)
        pendientes.extend(_imports_nucleo(modulo) - vistos)
    primero = [m for m in _PRIMERO if m in vistos]
    return primero + sorted(vistos - set(primero))


def _version():
    try:
        commit = subprocess.run(
            ["git", "-C", str(_SCRIPT.parent), "rev-parse", "--short", "HEAD"],
            capture_output=True, text=True, timeout=10,
        ).stdout.strip()
    except (OSError, subprocess.SubprocessError):
        commit = ""
    return commit or "sin version de git"


def _primera_linea(fuente):
    try:
        doc = ast.get_docstring(ast.parse(fuente)) or ""
    except SyntaxError:
        doc = ""
    return doc.strip().splitlines()[0] if doc.strip() else ""


def texto_recalculador(aamm=None):
    """El contenido del archivo Calcular_11_PAGOS_BESS_AAMM.py."""

    fuentes = {}
    for modulo in modulos_del_calculo():
        fuentes[f"nucleo/{modulo}.py"] = (_NUCLEO / f"{modulo}.py").read_text(encoding="utf-8")
    fuentes["nucleo/__init__.py"] = _INIT_NUCLEO
    fuentes["nucleo/externos.py"] = _EXTERNOS
    for ruta in _OTROS:
        fuentes[ruta] = (_SCRIPT / ruta).read_text(encoding="utf-8")

    for ruta, fuente in fuentes.items():
        if "'''" in fuente or fuente.endswith("\\"):
            raise ErrorEntrada(f"{ruta} no se puede incluir como texto (trae ''').")

    periodo = f" del periodo {aamm}" if aamm else ""
    indice = "\n".join(
        f"    {ruta:<34} {_primera_linea(fuente)}" for ruta, fuente in fuentes.items()
    )
    lineas = [
        "# -*- coding: utf-8 -*-",
        '"""',
        f"Calcular 11_PAGOS_BESS_AAMM.xlsx{periodo} (antecedentes SAE).",
        "",
        "Calcula las hojas de calculo (Calculo E Costos, Calculo RE545,",
        "COMPENSACION_CENTRAL, PRORRATA_RETIROS, Resumen) con los datos del",
        "propio archivo 11_, en una copia (..._calculado.xlsx). El archivo",
        "original no se modifica.",
        "",
        "Uso: doble clic, o 'python Calcular_11_PAGOS_BESS_AAMM.py'. Si este",
        "archivo esta en la carpeta '03 Antecedentes SAE', toma solo el 11_ de",
        "la carpeta de arriba; si no, se elige con 'Examinar...'.",
        "Sin ventana: python Calcular_11_PAGOS_BESS_AAMM.py ruta_del_11.xlsx --sin-ventana",
        "",
        "Requisitos: Python 3.10 o mas nuevo, con pandas y openpyxl",
        "(pip install pandas openpyxl). python-calamine es opcional (lee mas",
        "rapido).",
        "",
        f"Generado el {datetime.date.today():%d-%m-%Y} desde el programa Balance BESS "
        f"(version {_version()}).",
        "",
        "Los modulos del calculo van mas abajo, en FUENTES, tal cual estan en",
        "el programa (un texto por modulo):",
        indice,
        '"""',
        _CARGADOR,
        "",
        "FUENTES = {}",
        "",
    ]
    for ruta, fuente in fuentes.items():
        lineas += [
            "",
            "# " + "=" * 68,
            f"# {ruta}",
            "# " + "=" * 68,
            f"FUENTES[{ruta!r}] = r'''" + fuente + "'''",
        ]
    lineas += ["", "", 'if __name__ == "__main__":', "    main()", ""]
    return "\n".join(lineas)


def nombre_recalculador(aamm):
    # "Calcular", no "Recalcular" (pedido del usuario 2026-11-10).
    return f"Calcular_11_PAGOS_BESS_{aamm}.py"


def escribir_recalculador(destino, aamm=None):
    """Escribe el recalculador en destino (carpeta o archivo .py)."""

    destino = Path(destino)
    if destino.suffix.lower() != ".py":
        destino = destino / nombre_recalculador(aamm or "AAMM")
    destino.parent.mkdir(parents=True, exist_ok=True)
    destino.write_text(texto_recalculador(aamm), encoding="utf-8")
    return destino


def generar_antecedentes(carpeta_base, aamm, incluir_codigo=True,
                         incluir_csv=True, otros=(), registrar=print,
                         progreso=None):
    """
    Crea '03 Antecedentes SAE' al lado del 11_ con el recalculador, el
    SAE_AAMM.csv y los archivos de 'otros' (rutas del caso; se copian
    con su carpeta relativa al caso, para que no choquen nombres).
    Devuelve la carpeta.
    """

    aamm = validar_aamm(aamm)
    rutas = resolver_rutas(carpeta_base, aamm)
    if not rutas["salida"].is_file():
        raise ErrorEntrada(f"No esta {rutas['salida'].name}: primero calcula la planilla.")
    carpeta = rutas["salida"].parent / CARPETA_ANTECEDENTES
    carpeta.mkdir(exist_ok=True)
    registrar(f"Carpeta: {carpeta}")

    if incluir_codigo:
        ruta = escribir_recalculador(carpeta / nombre_recalculador(aamm), aamm)
        registrar(f"  codigo de recalculo: {ruta.name} ({ruta.stat().st_size // 1024:,} KB)")

    if incluir_csv:
        csv = rutas["base"] / nombre_csv_sae(aamm)
        if csv.is_file():
            shutil.copy2(csv, carpeta / csv.name)
            registrar(f"  {csv.name}")
        else:
            registrar(f"  [AVISO] no esta {csv.name}: generalo con su boton y vuelve a generar los antecedentes.")

    base = Path(rutas["base"]).resolve()
    for otro in otros:
        otro = Path(otro)
        if not otro.is_file():
            registrar(f"  [AVISO] no esta {otro}: no se copia.")
            continue
        try:
            relativa = otro.resolve().relative_to(base)
        except ValueError:
            relativa = Path(otro.name)
        destino = carpeta / relativa
        destino.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(otro, destino)
        registrar(f"  {relativa}")

    if progreso:
        progreso(100)
    registrar(f"Listo: {carpeta}")
    return carpeta
