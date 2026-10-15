# -*- coding: utf-8 -*-
"""
Formato de lectura de las hojas de salida (ancho, negrita, panel fijo).

No toca ni un valor: solo el aspecto con el que se abren los dos
libros. El calculo ya esta hecho cuando esto corre -- si alguna vez
alguna de estas funciones cambia un dato, es un bug, no una opcion.

Pedido del usuario: "necesito que ordenes las planillas". Antes las
hojas salian tal cual las deja pandas: encabezados sin negrita,
columnas de ancho 8 (todos los numeros como ####), sin panel fijo, asi
que revisar 100.000 filas obligaba a formatear a mano cada vez que se
regeneraba el libro.
"""

from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


# Ancho de columna en "caracteres" de Excel. El maximo existe para que
# una Barra o un nombre de central largo no empuje el resto de la hoja
# fuera de la pantalla.
ANCHO_MINIMO = 9
ANCHO_MAXIMO = 38

# Cuantas filas de datos se miran para decidir ancho y formato de
# numero. Con 200 alcanza para ver la forma de la columna y no cuesta
# nada en una hoja de 100.000 filas.
FILAS_MUESTRA = 200

# Tope de celdas a las que se les escribe el formato de numero una por
# una (openpyxl no tiene formato por columna que Excel respete). El
# tope esta puesto MUY arriba a proposito: un mes 15-minutal de una
# decena de centrales son ~30.000 filas, o sea ~1,5 millones de celdas
# en "Calculo E Costos" (medido: ~1 segundo cada millon y medio). Es
# un freno para un caso disparatado, no algo que un periodo normal
# tenga que tocar -- si lo toca, el libro sale igual, solo que sin
# separador de miles.
MAXIMO_CELDAS_FORMATO = 8_000_000

# Encabezados: "Azul, Enfasis 1, Oscuro 50%" de Excel, letra blanca en
# negrita (pedido del usuario, para todas las hojas). Cada tabla va
# con su contorno enmarcado (borde fino), asi dos tablas de una misma hoja (CSF y CPF
# en FD, los cuadros de PRORRATA_RETIROS...) se distinguen solas.
COLOR_ENCABEZADO = "244062"
RELLENO_ENCABEZADO = PatternFill(
    fill_type="solid", start_color=COLOR_ENCABEZADO, end_color=COLOR_ENCABEZADO
)
LETRA_ENCABEZADO = Font(bold=True, color="FFFFFF")
_LINEA = Side(style="thin", color="808080")
BORDE = Border(left=_LINEA, right=_LINEA, top=_LINEA, bottom=_LINEA)

FORMATO_ENTERO = "#,##0"
FORMATO_DECIMAL = "#,##0.00"
FORMATO_FECHA = "yyyy-mm-dd hh:mm"
# Numeros que son claves y no cantidades (una fecha AAAAMMDD, un año):
# sin separador de miles, que 20260905 no se lea 20.260.905.
FORMATO_CLAVE = "0"

TITULOS_CLAVE = ("año", "ano", "anio", "year", "aamm", "fecha")


def _es_fecha_aaaammdd(valor):
    entero = int(valor)
    if float(valor) != entero or not 19000101 <= entero <= 21001231:
        return False
    mes, dia = (entero // 100) % 100, entero % 100
    return 1 <= mes <= 12 and 1 <= dia <= 31


def _es_titulo_clave(titulos):
    for titulo in titulos:
        texto = str(titulo or "").strip().lower()
        if texto in TITULOS_CLAVE:
            return True
    return False


def _es_numero(valor):
    return isinstance(valor, (int, float)) and not isinstance(valor, bool)


def _formato_de_columna(valores, titulos=()):
    """
    Decide el formato de numero de una columna a partir de una
    muestra de sus valores. Devuelve None si no hay que tocarla
    (columna de texto, vacia, o mezcla de tipos).
    """

    con_valor = [v for v in valores if v is not None and v != ""]

    if not con_valor:
        return None

    if all(hasattr(v, "year") and hasattr(v, "month") for v in con_valor):
        return FORMATO_FECHA

    if not all(_es_numero(v) for v in con_valor):
        return None

    if any(float(v) != int(v) for v in con_valor):
        return FORMATO_DECIMAL

    if all(_es_fecha_aaaammdd(v) for v in con_valor) or _es_titulo_clave(titulos):
        return FORMATO_CLAVE

    return FORMATO_ENTERO


def _texto_de_celda(valor):
    if valor is None:
        return ""
    if isinstance(valor, float):
        # Un 1234.5678901 mide 12 caracteres en memoria pero se ve
        # como 1.234,57 una vez aplicado FORMATO_DECIMAL.
        return f"{valor:,.2f}"
    return str(valor)


def formatear_hoja(ws, fila_encabezado=1):
    """
    Deja una hoja lista para leer:

      - las filas 1..fila_encabezado (titulos de cuadro, encabezados
        de grupo y nombres de columna) con fondo azul oscuro (#244062),
        letra blanca en negrita y centradas;
      - el contorno de cada tabla enmarcado (borde fino), hasta su
        ultima fila;
      - panel inmovilizado justo debajo del encabezado, para que los
        nombres de columna sigan a la vista al bajar;
      - ancho de columna segun lo que realmente hay adentro;
      - separador de miles (y dos decimales cuando la columna los
        tiene) en las columnas numericas, y fecha legible en las de
        fecha.

    fila_encabezado: fila (1-indexada) donde estan los NOMBRES de
    columna. Es 1 en las hojas que escribe pandas solas, 2 en las que
    llevan una fila de encabezados de grupo arriba (Calculo E Costos,
    Calculo RE545, Ofertas SSCC) y 3 en PRORRATA_RETIROS (titulo de
    la hoja + titulo de cada cuadro).
    """

    if ws.max_row <= fila_encabezado:
        # Hoja vacia o sin filas de datos: no hay nada que ordenar.
        return

    # Las columnas de las tablas: las que traen nombre en la fila de
    # encabezado. Una columna sin nombre es la separacion entre dos
    # tablas y no se pinta ni se enmarca.
    columnas_tabla = [
        celda.column
        for celda in ws[fila_encabezado]
        if celda.value is not None and str(celda.value).strip() != ""
    ]

    combinadas = set()
    for rango in ws.merged_cells.ranges:
        for fila in range(rango.min_row, rango.max_row + 1):
            for columna in range(rango.min_col, rango.max_col + 1):
                combinadas.add((fila, columna))

    for fila in ws.iter_rows(min_row=1, max_row=fila_encabezado):
        for celda in fila:
            en_combinada = (celda.row, celda.column) in combinadas
            if celda.value is None and not en_combinada:
                continue
            if celda.value is not None:
                celda.font = LETRA_ENCABEZADO
                celda.alignment = Alignment(
                    horizontal="center", vertical="center", wrap_text=True
                )
            # El titulo de un cuadro o de un grupo (filas de arriba) se
            # pinta igual, combinado o no.
            celda.fill = RELLENO_ENCABEZADO
            celda.border = BORDE

    ws.freeze_panes = ws.cell(row=fila_encabezado + 1, column=1)

    primera_fila_datos = fila_encabezado + 1
    ultima_muestra = min(ws.max_row, primera_fila_datos + FILAS_MUESTRA - 1)

    muestra = [
        [celda.value for celda in fila]
        for fila in ws.iter_rows(
            min_row=primera_fila_datos, max_row=ultima_muestra
        )
    ]

    formatos = {}

    for indice in range(ws.max_column):

        columna = [
            fila[indice] for fila in muestra
            if indice < len(fila)
        ]

        titulos = [
            ws.cell(row=fila, column=indice + 1).value
            for fila in range(1, fila_encabezado + 1)
        ]

        anchos = [
            len(_texto_de_celda(valor))
            for valor in columna + titulos
        ]

        letra = get_column_letter(indice + 1)
        ws.column_dimensions[letra].width = min(
            ANCHO_MAXIMO, max(ANCHO_MINIMO, max(anchos, default=0) + 2)
        )

        formato = _formato_de_columna(columna, titulos)
        if formato:
            formatos[indice + 1] = formato

    if not columnas_tabla:
        return

    celdas = (ws.max_row - fila_encabezado) * len(columnas_tabla)
    if celdas > MAXIMO_CELDAS_FORMATO:
        return

    # Una pasada por los datos: formato de numero y, para el marco, la
    # ultima fila con dato de cada tabla (las tablas de una hoja tienen
    # distinto largo; la corta no se enmarca hasta el fondo de la larga).
    tabla_de = _tablas_contiguas(columnas_tabla)
    ultima_de_tabla = {}
    for fila in ws.iter_rows(
        min_row=primera_fila_datos,
        min_col=min(columnas_tabla),
        max_col=max(columnas_tabla),
    ):
        for celda in fila:
            if celda.value is None:
                continue
            formato = formatos.get(celda.column)
            if formato:
                celda.number_format = formato
            tabla = tabla_de.get(celda.column)
            if tabla is not None:
                ultima_de_tabla[tabla] = celda.row

    # Solo el contorno de cada tabla: enmarcar celda por celda multiplica
    # por diez el tiempo de una hoja grande (medido: 3 s -> 28 s en
    # 30.000 x 60), y el contorno ya separa una tabla de la otra.
    for tabla, ultima in ultima_de_tabla.items():
        columnas = [c for c, t in tabla_de.items() if t == tabla]
        izquierda, derecha = min(columnas), max(columnas)
        for fila in range(primera_fila_datos, ultima + 1):
            for columna in {izquierda, derecha}:
                celda = ws.cell(row=fila, column=columna)
                celda.border = Border(
                    left=_LINEA if columna == izquierda else None,
                    right=_LINEA if columna == derecha else None,
                    bottom=_LINEA if fila == ultima else None,
                )
        for columna in range(izquierda + 1, derecha):
            ws.cell(row=ultima, column=columna).border = Border(bottom=_LINEA)


def _tablas_contiguas(columnas):
    """{columna: numero de tabla}: columnas seguidas son una tabla."""

    tabla_de = {}
    tabla = -1
    anterior = None
    for columna in sorted(columnas):
        if anterior is None or columna != anterior + 1:
            tabla += 1
        tabla_de[columna] = tabla
        anterior = columna
    return tabla_de


def formatear_libro(wb, filas_encabezado=None):
    """
    Aplica formatear_hoja() a todas las hojas del libro.

    filas_encabezado: {nombre de hoja: fila de los nombres de
    columna}. Las hojas que no esten en el diccionario usan 1.
    """

    filas_encabezado = filas_encabezado or {}

    for nombre in wb.sheetnames:
        formatear_hoja(wb[nombre], filas_encabezado.get(nombre, 1))
