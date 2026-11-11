# -*- coding: utf-8 -*-
r"""
El SoC de cada BESS, directo desde PI AF (antes: el Excel que se armaba
a mano con PI DataLink, un bloque de columnas por central).

Reemplaza al script de prueba PruebaPI.py (que comparaba el Excel de
DataLink contra AF): la logica de conexion y lectura es la misma, pero
ahora el resultado ES el archivo de SoC del periodo.

De donde sale cada central
--------------------------
1. La ruta del elemento en AF, si la hoja Diccionario de
   Centrales.xlsx la tiene (columna "SoC_PI"). Se acepta pegada tal
   cual sale de DataLink, con el atributo al final:

       \\SRV-SCADA-AF1\SEN\Generación\SEN\03 Región II\SAE-Tocopilla|Nombre PCP/PID

   Lo que va despues de "|" se ignora, y si la ruta apunta a AF1 se
   busca el MISMO camino en AF2 (el acceso que hay es a AF2).
2. Si no hay ruta, o la ruta no existe en AF2, se buscan todos los
   elementos "SAE*" de la base y se empareja por su atributo
   "Nombre PCP/PID" contra cualquiera de los nombres de la central en
   el Diccionario (es lo que hacia PruebaPI.py).

El valor es el atributo "Porc Carga Disponible" (viene en %, se divide
por 100 para dejarlo como fraccion, igual que el Excel de antes),
interpolado cada 15 minutos desde el primer dia del mes a las 00:00
hasta el primero del mes siguiente a las 00:00 (incluido: el Excel de
antes tambien lo traia). La hora es la local del PC (America/Santiago),
asi que los dias de cambio de hora salen como salian de DataLink: en
septiembre falta la hora que el reloj se salta y en abril la que se
repite viene dos veces.

Solo corre en un PC con el AF SDK de OSIsoft instalado (Windows) y
pythonnet (`pip install pythonnet`). Se importa recien al usarse: el
resto del programa no lo necesita.

Como los demas modulos de Medidas/, no importa nada de nucleo y levanta
ErrorMedidas.
"""

import sys

from .comun import ErrorMedidas, normalizar


# ===================== CONFIGURACION =====================
SERVIDOR_AF = "SRV-SCADA-AF2"
# Servidores que se cambian por SERVIDOR_AF en las rutas del
# Diccionario (el acceso que hay es a AF2).
SERVIDORES_REEMPLAZADOS = ("SRV-SCADA-AF1",)
BASE_AF = "SEN"
ATRIBUTO_SOC = "Porc Carga Disponible"       # viene en %, se divide por 100
ATRIBUTO_NOMBRE = "Nombre PCP/PID"
FILTRO_BUSQUEDA = "Name:SAE*"
INTERVALO = "15m"
RUTA_AF_SDK = r"C:\Program Files (x86)\PIPC\AF\PublicAssemblies\4.0"
# =========================================================

ORIGEN_RUTA = "ruta del Diccionario"
ORIGEN_RUTA_AF1 = "ruta del Diccionario (AF1 -> AF2)"
ORIGEN_BUSQUEDA = "busqueda por Nombre PCP/PID"


# ============================================================
# RUTAS (sin AF: se pueden probar en cualquier maquina)
# ============================================================

def limpiar_ruta(ruta):
    r"""
    (ruta_en_AF2, venia_de_otro_servidor) a partir de lo que se pego en
    el Diccionario. Devuelve ("", False) si no hay ruta.

        \\SRV-SCADA-AF1\SEN\X\Y|Nombre PCP/PID  ->  (\\SRV-SCADA-AF2\SEN\X\Y, True)
    """

    if ruta is None:
        return "", False

    texto = str(ruta).strip()
    if not texto or texto.lower() == "nan":
        return "", False

    texto = texto.split("|")[0].strip().rstrip("\\")

    if not texto.startswith("\\\\"):
        # Una ruta sin servidor (ej. "SEN\Generación\...") se toma como
        # relativa a SERVIDOR_AF.
        texto = f"\\\\{SERVIDOR_AF}\\{texto.lstrip(chr(92))}"

    partes = texto[2:].split("\\")
    cambio = False

    if partes[0].upper() in {s.upper() for s in SERVIDORES_REEMPLAZADOS}:
        partes[0] = SERVIDOR_AF
        cambio = True

    return "\\\\" + "\\".join(partes), cambio


def partes_de_ruta(ruta):
    r"""
    (servidor, base, [elementos...]) de una ruta \\SERVIDOR\BASE\A\B\C.
    ErrorMedidas si no tiene al menos servidor, base y un elemento.
    """

    partes = [p for p in ruta[2:].split("\\") if p.strip()]

    if len(partes) < 3:
        raise ErrorMedidas(
            f"La ruta '{ruta}' no tiene la forma "
            f"\\\\SERVIDOR\\BASE\\...\\ELEMENTO."
        )

    return partes[0], partes[1], partes[2:]


# ============================================================
# AF (solo en el PC con el AF SDK)
# ============================================================

def _cargar_af_sdk():
    """Importa el AF SDK. ErrorMedidas con el motivo si no se puede."""

    try:
        import clr  # pythonnet
    except ImportError as error:
        raise ErrorMedidas(
            "Para traer el SoC desde PI hace falta pythonnet: "
            "pip install pythonnet"
        ) from error

    if RUTA_AF_SDK not in sys.path:
        sys.path.append(RUTA_AF_SDK)

    try:
        clr.AddReference("OSIsoft.AFSDK")
        from OSIsoft.AF import PISystems
        from OSIsoft.AF.Search import AFElementSearch
        from OSIsoft.AF.Time import AFTimeRange, AFTimeSpan
    except Exception as error:  # el error real es de .NET
        raise ErrorMedidas(
            f"No se pudo cargar el AF SDK de OSIsoft (se busca en "
            f"{RUTA_AF_SDK}). Este boton solo funciona en un PC con PI "
            f"AF instalado.\nDetalle: {error}"
        ) from error

    return {
        "PISystems": PISystems,
        "AFElementSearch": AFElementSearch,
        "AFTimeRange": AFTimeRange,
        "AFTimeSpan": AFTimeSpan,
    }


class ConexionAF:
    """La base SEN de AF2, con lo necesario para encontrar cada central."""

    def __init__(self, registrar=print):
        self.sdk = _cargar_af_sdk()
        self.registrar = registrar

        try:
            sistema = self.sdk["PISystems"]()[SERVIDOR_AF]
            if sistema is None:
                raise ErrorMedidas(
                    f"El servidor {SERVIDOR_AF} no esta registrado en "
                    f"este PC (PI System Explorer -> Servidores)."
                )
            sistema.Connect()
            self.base = sistema.Databases[BASE_AF]
        except ErrorMedidas:
            raise
        except Exception as error:
            raise ErrorMedidas(
                f"No se pudo conectar a {SERVIDOR_AF}: {error}"
            ) from error

        if self.base is None:
            raise ErrorMedidas(
                f"{SERVIDOR_AF} no tiene la base '{BASE_AF}'."
            )

        self._por_nombre = None

    # ---- buscar el elemento ----------------------------------------

    def elemento_por_ruta(self, ruta):
        """El elemento de esa ruta en AF2, o None si no esta."""

        servidor, base, elementos = partes_de_ruta(ruta)

        if normalizar(servidor) != normalizar(SERVIDOR_AF):
            return None
        if normalizar(base) != normalizar(BASE_AF):
            return None

        actual = self.base.Elements[elementos[0]]
        for nombre in elementos[1:]:
            if actual is None:
                return None
            actual = actual.Elements[nombre]

        return actual

    def elementos_por_nombre_pcp(self):
        """{Nombre PCP/PID normalizado: elemento} de los SAE* (una vez)."""

        if self._por_nombre is not None:
            return self._por_nombre

        self.registrar(
            f"  buscando elementos '{FILTRO_BUSQUEDA}' en "
            f"{SERVIDOR_AF}\\{BASE_AF} (para las centrales sin ruta)..."
        )

        busqueda = self.sdk["AFElementSearch"](
            self.base, "BuscarSAE", FILTRO_BUSQUEDA
        )

        encontrados = {}
        for elem in busqueda.FindObjects(0, True, 1000):
            atributo = elem.Attributes[ATRIBUTO_NOMBRE]
            if atributo is None or elem.Attributes[ATRIBUTO_SOC] is None:
                continue
            try:
                nombre = atributo.GetValue().Value
            except Exception:
                continue
            if nombre:
                encontrados.setdefault(normalizar(str(nombre)), elem)

        self.registrar(
            f"  elementos SAE con '{ATRIBUTO_SOC}': {len(encontrados)}"
        )
        self._por_nombre = encontrados
        return encontrados

    def ubicar(self, ruta_diccionario, nombres):
        """
        (elemento, ruta_usada, origen) de una central, o (None, motivo,
        None) si no se encontro.

        nombres: todos los nombres de la central en el Diccionario (el
        de Balance_BESS primero), para la busqueda por Nombre PCP/PID.
        """

        ruta, era_af1 = limpiar_ruta(ruta_diccionario)
        motivo = "sin ruta en el Diccionario"

        if ruta:
            try:
                elemento = self.elemento_por_ruta(ruta)
            except ErrorMedidas as error:
                elemento, motivo = None, str(error)
            else:
                if elemento is not None:
                    return (
                        elemento, ruta,
                        ORIGEN_RUTA_AF1 if era_af1 else ORIGEN_RUTA,
                    )
                motivo = f"la ruta {ruta} no existe en {SERVIDOR_AF}"

        por_nombre = self.elementos_por_nombre_pcp()
        for nombre in nombres:
            elemento = por_nombre.get(normalizar(nombre))
            if elemento is not None:
                return elemento, str(elemento.GetPath()), ORIGEN_BUSQUEDA

        return None, motivo, None

    # ---- leer el SoC -----------------------------------------------

    def leer_soc(self, elemento, inicio, fin):
        """
        [(fecha_local 'AAAA-MM-DD HH:MM:SS', soc_fraccion o None), ...]
        cada INTERVALO entre inicio y fin (textos 'AAAA-MM-DD HH:MM:SS',
        hora local).
        """

        rango = self.sdk["AFTimeRange"](inicio, fin)
        intervalo = self.sdk["AFTimeSpan"].Parse(INTERVALO)

        valores = elemento.Attributes[ATRIBUTO_SOC].Data.InterpolatedValues(
            rango, intervalo, None, None, False
        )

        datos = []
        for valor in valores:
            fecha = valor.Timestamp.LocalTime.ToString("yyyy-MM-dd HH:mm:ss")
            soc = None
            if valor.IsGood:
                try:
                    soc = float(valor.Value) / 100
                except (TypeError, ValueError):
                    soc = None
            datos.append((str(fecha), soc))

        return datos


def descargar(centrales, inicio, fin, registrar=print, progreso=None,
              desde=0, hasta=100, conexion=None):
    """
    centrales: lista de dicts {"central", "ruta", "nombres"} (ruta puede
    venir vacia; nombres = todos los del Diccionario).

    Devuelve (filas, resumen):
      filas   lista de (central, fecha_texto, soc)
      resumen lista de dicts {central, estado, ruta_af, origen,
              registros, sin_dato}
    """

    conexion = conexion or ConexionAF(registrar=registrar)

    filas = []
    resumen = []
    total = max(len(centrales), 1)

    for numero, central in enumerate(centrales, start=1):

        nombre = central["central"]
        elemento, ruta, origen = conexion.ubicar(
            central.get("ruta"), central.get("nombres") or [nombre]
        )

        if elemento is None:
            registrar(f"  [NO ENCONTRADA EN AF] {nombre}: {ruta}")
            resumen.append({
                "central": nombre, "estado": "No encontrada en AF",
                "ruta_af": None, "origen": ruta,
                "registros": 0, "sin_dato": 0,
            })
            continue

        registrar(f"  {nombre} <- {ruta}  ({origen})")

        try:
            datos = conexion.leer_soc(elemento, inicio, fin)
        except Exception as error:  # el error real es de .NET
            registrar(f"  [ERROR] {nombre}: {error}")
            resumen.append({
                "central": nombre, "estado": f"Error al leer: {error}",
                "ruta_af": ruta, "origen": origen,
                "registros": 0, "sin_dato": 0,
            })
            continue

        sin_dato = sum(1 for _, soc in datos if soc is None)
        filas.extend((nombre, fecha, soc) for fecha, soc in datos)

        resumen.append({
            "central": nombre, "estado": "OK",
            "ruta_af": ruta, "origen": origen,
            "registros": len(datos), "sin_dato": sin_dato,
        })

        if sin_dato:
            registrar(
                f"    {sin_dato:,} de {len(datos):,} valores con mala "
                f"calidad en PI (quedan vacios)"
            )

        if progreso:
            progreso(desde + (hasta - desde) * numero / total)

    return filas, resumen
