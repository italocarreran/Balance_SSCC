# -*- coding: utf-8 -*-
"""
Las filas de DATOS de las hojas grandes, escritas directo en XML.

Pedido del usuario (2026-11-03): "hacerlo mucho más rápido [...] que no
cague la planilla, mismos resultados". En un mes real, escribir el libro
de calculo tardaba 195 s: openpyxl crea un objeto por celda (unos 5
millones), le da formato celda por celda y despues serializa todo.

Ahora cada hoja se arma en dos partes:

  1. el ESQUELETO, con openpyxl y el codigo de siempre: encabezados,
     titulos de cuadro, comentarios, anchos, panel, totales... pero con
     cada tabla cortada a sus primeras FILAS_MUESTRA filas. Es la misma
     muestra con la que formato.formatear_hoja() ya decidia anchos y
     formatos de numero, asi que salen iguales;
  2. el RELLENO (este modulo): todas las filas de datos, generadas como
     texto XML con las MISMAS reglas que aplicaban pandas, openpyxl,
     formatear_hoja() y los _terminar_* de escritura.py, celda por celda:

       - valor: como lo escribe pandas (NaN -> texto vacio, numpy ->
         Python, fechas con su formato por defecto) y como lo serializa
         openpyxl (numeros con "%.16g", textos en linea, "=..." como
         formula);
       - formato de numero de la columna (el que decidio la muestra), en
         toda celda con valor (el texto vacio cuenta como valor);
       - el contorno de cada tabla (borde izquierdo, derecho y abajo en
         su ultima fila);
       - lo que agregan los _terminar_*: formulas por fila, guion en las
         curvas sin puesto, relleno verde del Monto, la alineacion y el
         "0.00" de las columnas de respuesta de FD.

xlsx_rapido.pegar_hojas() reemplaza las filas de datos del esqueleto por
las de aca, ya con los estilos del libro final. La prueba de que da lo
mismo es desarrollo/tests/test_relleno_rapido.py: el libro sale celda por
celda igual (valor y estilo) que escribiendolo entero con openpyxl.
"""

import math
import re
from datetime import date, datetime, timedelta
from xml.sax.saxutils import escape

import numpy as np
import pandas as pd
from openpyxl.cell import Cell
from openpyxl.styles import Alignment, Border, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.utils.datetime import to_excel

# Formatos que pone pandas a las fechas (ExcelWriter por defecto).
FORMATO_FECHA_HORA_PANDAS = "YYYY-MM-DD HH:MM:SS"
FORMATO_FECHA_PANDAS = "YYYY-MM-DD"

# Marca de la fila en las plantillas de formula (ver plantilla_por_fila).
MARCA_FILA = "\x00"


def plantilla_por_fila(plantilla, letras):
    """'={AG}*{AM}' -> '=P\\x00*T\\x00' (el \\x00 es el numero de fila)."""

    return re.sub(
        r"\{(\w+)\}", lambda m: f"{letras[m.group(1)]}{MARCA_FILA}", plantilla
    )


# ------------------------------------------------------------
# Valores: lo mismo que pandas (_format_value + _value_with_fmt) y
# openpyxl (_set_attributes + safe_string).
# ------------------------------------------------------------

def _valor_pandas(valor):
    """(valor, formato de pandas o None), como lo manda pandas a openpyxl."""

    if pd.api.types.is_scalar(valor) and pd.isna(valor):
        return "", None
    if isinstance(valor, (bool, np.bool_)):
        return bool(valor), None
    if isinstance(valor, (int, np.integer)):
        return int(valor), None
    if isinstance(valor, (float, np.floating)):
        valor = float(valor)
        if math.isinf(valor):
            return ("inf" if valor > 0 else "-inf"), None
        return valor, None
    if isinstance(valor, datetime):
        return valor, FORMATO_FECHA_HORA_PANDAS
    if isinstance(valor, date):
        return valor, FORMATO_FECHA_PANDAS
    if isinstance(valor, timedelta):
        return valor.total_seconds() / 86400, "0"
    return str(valor), None


def valor_guardado_xml(valor):
    """
    (atributo t, texto de <v>) del valor calculado que se guarda junto a
    una formula (ver escritura._formulas_de_prueba_en_fila). Vacio si no
    hay valor: Excel lo calcula al abrir.
    """

    if isinstance(valor, (bool, np.bool_)):
        return ' t="b"', "1" if valor else "0"
    if isinstance(valor, (int, float, np.integer, np.floating)):
        if not math.isfinite(float(valor)):
            return "", ""
        return "", "%.16g" % float(valor)
    if isinstance(valor, str) and valor != "":
        return ' t="str"', escape(valor)
    return "", ""


def celda_formula_xml(ref, estilo, formula, matricial, valor):
    """<c> con formula (y su valor calculado, si hay). estilo: ' s="n"'."""

    t, v = valor_guardado_xml(valor)
    f = f'<f t="array" ref="{ref}">' if matricial else "<f>"
    return (
        f'<c r="{ref}"{estilo}{t}>{f}{escape(formula[1:])}</f>'
        f"<v>{v}</v></c>"
    )


def _texto_xml(texto):
    t = escape(texto)
    preservar = texto.strip() and texto != texto.strip()
    if preservar:
        return f'<is><t xml:space="preserve">{t}</t></is>'
    return f"<is><t>{t}</t></is>"


def _celda(ref, estilo, valor):
    """El XML de una celda (estilo ya es el id del libro final o None)."""

    s = f' s="{estilo}"' if estilo is not None else ""
    if isinstance(valor, bool):
        return f'<c r="{ref}"{s} t="b"><v>{int(valor)}</v></c>'
    if isinstance(valor, (int, float)):
        if isinstance(valor, float) and (math.isnan(valor) or math.isinf(valor)):
            return f'<c r="{ref}"{s} t="n"><v></v></c>'
        return f'<c r="{ref}"{s} t="n"><v>{"%.16g" % valor}</v></c>'
    if isinstance(valor, (datetime, date)):
        return f'<c r="{ref}"{s} t="n"><v>{"%.16g" % to_excel(valor)}</v></c>'
    if valor == "":
        # Sin estilo openpyxl ni la escribe (pandas deja el NaN vacio).
        return f'<c r="{ref}"{s} t="inlineStr"/>' if s else ""
    if valor.startswith("=") and len(valor) > 1:
        return f'<c r="{ref}"{s}><f>{escape(valor[1:])}</f><v></v></c>'
    return f'<c r="{ref}"{s} t="inlineStr">{_texto_xml(valor)}</c>'


# ------------------------------------------------------------
# El relleno de una hoja
# ------------------------------------------------------------

class RellenoHoja:
    """
    Las filas de datos de una hoja (todas sus tablas), listas para
    generarlas en XML.

    ws: la hoja del esqueleto, ya formateada (formato.formatear_hoja deja
    en ws._balance_formato lo que decidio) y terminada (los _terminar_*
    dejan en ws._balance_extras lo que agregan por fila).
    bloques: [(fila de la primera fila de datos, columna de la primera
    columna, DataFrame completo)] -- lo que se escribio cortado.
    """

    def __init__(self, ws, bloques):
        self.ws = ws
        self.bloques = bloques
        info = getattr(ws, "_balance_formato", None) or {}
        extras = getattr(ws, "_balance_extras", None) or {}

        self.fila_encabezado = info.get("fila_encabezado", 1)
        self.primera = self.fila_encabezado + 1
        self.con_formato = info.get("con_formato", False)
        formatos = info.get("formatos", {}) if self.con_formato else {}
        tabla_de = info.get("tabla_de", {}) if self.con_formato else {}

        formulas = extras.get("formulas", {})          # letra -> plantilla
        guion = extras.get("guion", {})                # letra -> texto
        rellenos = extras.get("relleno", {})           # letra -> (PatternFill, ultima)
        izquierda = extras.get("izquierda_000", set())  # letras de FD
        # letra -> {fila: (formula, matricial)}: una sola celda con
        # formula (la de prueba) en una columna de valores.
        self.formula_celda = extras.get("formula_celda", {})
        self.guardados = getattr(ws, "_balance_guardados", None) or {}

        # Ultima fila de cada tabla: la ultima fila con dato (los vacios
        # que deja pandas son texto vacio: tambien cuentan).
        ultima_de_tabla = {}
        for fila0, col0, df in bloques:
            for j in range(df.shape[1]):
                tabla = tabla_de.get(col0 + j)
                if tabla is not None and len(df):
                    ultima_de_tabla[tabla] = max(
                        ultima_de_tabla.get(tabla, 0), fila0 + len(df) - 1
                    )
        bordes_de = {}
        for tabla, ultima in ultima_de_tabla.items():
            columnas = [c for c, t in tabla_de.items() if t == tabla]
            bordes_de[tabla] = (min(columnas), max(columnas), ultima)

        self._estilos = {}
        self.columnas = []   # una entrada por columna de cada bloque
        for fila0, col0, df in bloques:
            for j in range(df.shape[1]):
                columna = col0 + j
                letra = get_column_letter(columna)
                tabla = tabla_de.get(columna)
                borde = bordes_de.get(tabla)
                self.columnas.append({
                    "fila0": fila0,
                    "letra": letra,
                    "columna": columna,
                    "valores": df.iloc[:, j],
                    "formato": formatos.get(columna),
                    "borde": borde,
                    "formula": formulas.get(letra),
                    "guion": guion.get(letra),
                    "relleno": rellenos.get(letra),
                    "izquierda_000": letra in izquierda,
                })

        self.ultima_fila = max(
            (fila0 + len(df) - 1 for fila0, _, df in bloques), default=0
        )
        self.ultima_columna = max(
            (col0 + df.shape[1] - 1 for _, col0, df in bloques), default=0
        )

    # --- estilos ---------------------------------------------------

    def _estilo(self, formato, borde, relleno, alineacion):
        """
        Id (del libro temporal) del estilo con esos atributos. Se arma con
        una celda suelta de la misma hoja, igual que formatear_hoja: el
        mismo estilo da el mismo id que en el esqueleto.
        """

        clave = (formato, borde, relleno, alineacion)
        if clave not in self._estilos:
            if getattr(self, "_cerrado", False):
                # Despues de guardar el temporal un estilo nuevo no
                # llegaria a su styles.xml: mejor fallar que escribirlo
                # con otro formato.
                raise RuntimeError(f"estilo no preparado: {clave}")
            celda = Cell(self.ws)
            if formato is not None:
                celda.number_format = formato
            if borde is not None:
                izquierda, derecha, abajo = borde
                from .formato import _LINEA
                celda.border = Border(
                    left=_LINEA if izquierda else None,
                    right=_LINEA if derecha else None,
                    bottom=_LINEA if abajo else None,
                )
            if relleno is not None:
                celda.fill = relleno
            if alineacion is not None:
                celda.alignment = Alignment(horizontal=alineacion)
            self._estilos[clave] = celda.style_id if celda.has_style else None
        return self._estilos[clave]

    def preparar_estilos(self):
        """
        Registra en el libro temporal todos los estilos que van a hacer
        falta (antes de guardarlo: despues ya no se pueden agregar).
        """

        for columna in self.columnas:
            for fila_es_ultima in (False, True):
                for con_valor in (True, False):
                    for tipo in ("normal", "fecha", "fecha_dia", "guion"):
                        self._clave_celda(
                            columna, fila_es_ultima, con_valor, tipo
                        )
        self._cerrado = True

    def _clave_celda(self, columna, es_ultima, con_valor, tipo):
        """Los atributos de estilo de una celda de la columna."""

        formato = None
        if con_valor:
            if tipo == "fecha":
                formato = FORMATO_FECHA_HORA_PANDAS
            elif tipo == "fecha_dia":
                formato = FORMATO_FECHA_PANDAS
            if columna["formato"] is not None:
                formato = columna["formato"]
            if columna["izquierda_000"]:
                formato = "0.00"
        borde = None
        if columna["borde"] is not None:
            izq, der, ultima = columna["borde"]
            c = columna["columna"]
            abajo = es_ultima and True
            if c in (izq, der):
                borde = (c == izq, c == der, abajo)
            elif abajo:
                borde = (False, False, True)
        relleno = columna["relleno"][0] if columna["relleno"] else None
        alineacion = None
        if columna["izquierda_000"] and con_valor:
            alineacion = "left"
        if tipo == "guion":
            alineacion = "right"
        return self._estilo(formato, borde, relleno, alineacion)

    # --- XML ---------------------------------------------------------

    def _celdas_de_columna(self, columna, mapa):
        """Lista (fila, xml) de la columna, con estilos del libro final."""

        def final(estilo):
            return None if estilo is None else mapa.get(estilo, 0)

        letra = columna["letra"]
        fila0 = columna["fila0"]
        ultima_tabla = columna["borde"][2] if columna["borde"] else None
        cache = {}

        def estilo(es_ultima, tipo):
            clave = (es_ultima, tipo)
            if clave not in cache:
                cache[clave] = final(
                    self._clave_celda(columna, es_ultima, True, tipo)
                )
            return cache[clave]

        plantilla = columna["formula"]
        guion = columna["guion"]
        valores = columna["valores"]
        filas = range(fila0, fila0 + len(valores))
        indice_ultima = (
            ultima_tabla - fila0
            if ultima_tabla is not None and 0 <= ultima_tabla - fila0 < len(valores)
            else None
        )

        # Columna numerica sin formula ni guion (la gran mayoria): todo de
        # una vez con numpy. Mismo texto que openpyxl ("%.16g"); el NaN
        # es el texto vacio que deja pandas.
        rapida = (
            plantilla is None and guion is None
            and (
                pd.api.types.is_float_dtype(valores.dtype)
                or pd.api.types.is_integer_dtype(valores.dtype)
            )
            and not pd.api.types.is_bool_dtype(valores.dtype)
        )
        if rapida:
            numeros = valores.to_numpy(dtype=float, na_value=np.nan)
            vacio = np.isnan(numeros) | np.isinf(numeros)
            textos = np.char.mod("%.16g", np.where(vacio, 0.0, numeros))
            s_normal = estilo(False, "normal")
            s_ultima = estilo(True, "normal")

            def a_s(valor):
                return f' s="{valor}"' if valor is not None else ""

            sn, su = a_s(s_normal), a_s(s_ultima)
            celdas = [
                (f'<c r="{letra}{fila}"{sn} t="inlineStr"/>' if sn else "")
                if es_vacio
                else f'<c r="{letra}{fila}"{sn} t="n"><v>{texto}</v></c>'
                for fila, es_vacio, texto in zip(filas, vacio.tolist(), textos.tolist())
            ]
            if indice_ultima is not None:
                fila = filas[indice_ultima]
                if vacio[indice_ultima]:
                    celdas[indice_ultima] = (
                        f'<c r="{letra}{fila}"{su} t="inlineStr"/>' if su else ""
                    )
                else:
                    celdas[indice_ultima] = (
                        f'<c r="{letra}{fila}"{su} t="n"><v>{textos[indice_ultima]}</v></c>'
                    )
            return celdas

        def a_s(valor):
            return f' s="{valor}"' if valor is not None else ""

        # Formula por fila: la plantilla se escapa una vez y solo cambia
        # el numero de fila (el valor de la celda se reemplaza siempre).
        if plantilla is not None:
            sn = a_s(estilo(False, "normal"))
            su = a_s(estilo(True, "normal"))
            cuerpo = escape(plantilla[1:])
            celdas = [
                f'<c r="{letra}{fila}"{sn}><f>'
                + cuerpo.replace(MARCA_FILA, str(fila)) + "</f><v></v></c>"
                for fila in filas
            ]
            if indice_ultima is not None:
                fila = filas[indice_ultima]
                celdas[indice_ultima] = (
                    f'<c r="{letra}{fila}"{su}><f>'
                    + cuerpo.replace(MARCA_FILA, str(fila)) + "</f><v></v></c>"
                )
            return celdas

        # Lo demas, celda por celda; cada valor distinto se convierte una
        # vez (las centrales, barras, etc. se repiten miles de veces).
        convertidos = {}
        celdas = []
        for i, original in enumerate(valores):
            fila = fila0 + i
            try:
                clave_valor = (type(original), original)
                hash(clave_valor)
            except TypeError:
                clave_valor = None
            if (
                clave_valor is not None and guion is None
                and isinstance(original, str) and original != ""
            ):
                parte = convertidos.get(clave_valor)
                if parte is None:
                    valor, _ = _valor_pandas(original)
                    parte = convertidos[clave_valor] = _celda("\x01", None, valor)
                es_ultima = i == indice_ultima
                ref = f'r="{letra}{fila}"' + a_s(estilo(es_ultima, "normal"))
                celdas.append(parte.replace('r="\x01"', ref, 1))
                continue
            valor, formato_pandas = _valor_pandas(original)
            tipo = "normal"
            if formato_pandas == FORMATO_FECHA_HORA_PANDAS:
                tipo = "fecha"
            elif formato_pandas == FORMATO_FECHA_PANDAS:
                tipo = "fecha_dia"
            if plantilla is not None:
                valor = plantilla.replace(MARCA_FILA, str(fila))
            elif guion is not None and valor == "":
                valor = guion
                tipo = "guion"
            celdas.append(_celda(
                f"{letra}{fila}", estilo(i == indice_ultima, tipo), valor
            ))
        return celdas

    def _poner_formula_celda(self, columna, celdas):
        """
        Cambia en 'celdas' (lista de la columna, desde su fila0) las que
        llevan formula de prueba, con el mismo estilo que tenia el valor
        y el valor como valor calculado.
        """

        letra = columna["letra"]
        for fila, (formula, matricial) in self.formula_celda.get(letra, {}).items():
            i = fila - columna["fila0"]
            if not 0 <= i < len(celdas):
                continue
            ref = f"{letra}{fila}"
            estilo = re.search(r' s="\d+"', celdas[i] or "")
            celdas[i] = celda_formula_xml(
                ref, estilo.group(0) if estilo else "", formula, matricial,
                self.guardados.get(ref),
            )

    def filas_xml(self, mapa):
        """
        Las filas de datos (<row>...</row>) como bytes, en orden, con los
        estilos ya pasados al libro final (mapa: id temporal -> final).
        """

        if not self.columnas:
            return
        primera = min(c["fila0"] for c in self.columnas)
        alto = self.ultima_fila - primera + 1
        # Cada columna rellenada con "" donde no tiene celda, para juntar
        # las filas con zip (en C) en vez de celda por celda.
        columnas = []
        for columna in sorted(self.columnas, key=lambda c: c["columna"]):
            celdas = self._celdas_de_columna(columna, mapa)
            self._poner_formula_celda(columna, celdas)
            antes = columna["fila0"] - primera
            despues = alto - antes - len(celdas)
            columnas.append([""] * antes + celdas + [""] * despues)

        for desplazamiento, celdas in enumerate(zip(*columnas)):
            yield (
                f'<row r="{primera + desplazamiento}">' + "".join(celdas)
                + "</row>"
            ).encode("utf-8")
