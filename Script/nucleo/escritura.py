# -*- coding: utf-8 -*-
"""
Escritura de la planilla de salida (una sola) y del archivo de
control de la corrida.

Son dos archivos, y ninguno de los dos es el de antes:

  - ARCHIVO_SALIDA: las entradas y el calculo JUNTOS en un mismo
    libro (antes eran Consolidado_entradas.xlsx y Pagos_BESS.xlsx),
    ordenado de fin a inicio -- el Resumen primero y las entradas al
    final (ORDEN_HOJAS_SALIDA).
  - ARCHIVO_CONTROL: las hojas de control de la corrida (Alertas,
    Ejecucion, Log), que ya no viven en la planilla de trabajo.

Las dos mitades de la planilla se siguen escribiendo por separado
(escribir_salida / escribir_pagos_bess: cada una tiene su boton en la
ventana y su propio calculo detras), pero como ahora comparten
archivo, cada una PRESERVA las hojas de la otra: lo que no se
recalcula en esta pasada se copia tal cual del libro anterior.
"""

import re
import shutil
import time
from pathlib import Path

import openpyxl
import pandas as pd
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

from .alertas import ALTA, Alerta

from .ecostos import (
    CUADROS_CALCULO_E_COSTOS, FORMULAS_E_COSTOS, ORIGEN_COLUMNAS_E_COSTOS,
    GUION_SIN_PUESTO, SIN_PUESTO_E_COSTOS, TOTALES_E_COSTOS,
)
from .factores_subastas import TITULOS_FMA
from .formato import formatear_libro
from .sellos import sellar
from .unidades import con_unidades
from .utiles import normalizar
from .hojas_entrada import fd_para_hoja, subastas_para_hoja
from .lectura import (
    TITULO_CENTRALES_DICCIONARIO, TITULO_CENTRALES_RESUMEN, titulo_de_cuadro,
)
from .ofertas_sscc import (
    HOJA_OFERTAS_SSCC, TITULO_OFERTAS_POR_DIA, TITULO_RESUMEN_VENTANA,
)
from .parametros import (
    HOJA_ALERTAS, HOJA_CALCULO_ECOSTOS, HOJA_CALCULO_RE545, HOJA_CENTRALES,
    COLUMNA_CENTRAL, HOJA_CMG, HOJA_FMA,
    HOJA_COMPENSACION_CENTRAL, HOJA_EJECUCION, HOJA_FD, HOJA_LOG,
    HOJA_MEDIDORES, HOJA_PRORRATA_RETIROS, HOJA_RESUMEN, HOJA_SUBASTAS,
    ORDEN_HOJAS_CONTROL, ORDEN_HOJAS_SALIDA,
)
from .re545 import COLUMNAS_SALIDA_RE545, GRUPOS_CALCULO_RE545
from .xlsx_rapido import nombres_de_hojas, pegar_hojas


# Las hojas que escribe cada mitad de la planilla. No definen el
# orden del libro (ese es ORDEN_HOJAS_SALIDA, comun a las dos): dicen
# de quien es cada hoja, o sea cual de las dos funciones la puede
# dejar vacia si no se regenero y no habia version anterior. Las de la
# otra mitad no se tocan nunca: se copian tal cual.
_HOJAS_PAGOS = (
    HOJA_RESUMEN, HOJA_CALCULO_ECOSTOS, HOJA_CALCULO_RE545,
    HOJA_COMPENSACION_CENTRAL, HOJA_PRORRATA_RETIROS,
)


def _abrir_existente(ruta):
    """
    El libro que hay hoy en 'ruta' (solo valores), o None si no
    existe o no se puede abrir. Se llama SIEMPRE antes de abrir el
    ExcelWriter: el writer trunca el archivo, asi que lo que no se
    haya leido antes se pierde.
    """

    if ruta is None:
        return None

    ruta = Path(ruta)

    if not ruta.is_file():
        return None

    try:
        return openpyxl.load_workbook(ruta, data_only=True)
    except Exception:
        return None


class _LibroEnZip:
    """
    El libro de salida que ya existe, visto solo por los nombres de sus
    hojas (se leen del indice del zip, sin abrir ninguna hoja). Las
    hojas que no se reescriben ni se copian: quedan en el zip tal cual y
    pegar_hojas() solo reemplaza las nuevas (ver xlsx_rapido.py).
    """

    def __init__(self, ruta):
        self.ruta = Path(ruta)
        self.sheetnames = nombres_de_hojas(ruta)


def _abrir_existente_en_zip(ruta):
    """_LibroEnZip de 'ruta', o None si no existe o no es un .xlsx."""

    if ruta is None or not Path(ruta).is_file():
        return None

    try:
        return _LibroEnZip(ruta)
    except Exception:
        return None


def _archivo_temporal(ruta_salida):
    """Donde se escriben las hojas nuevas antes de pegarlas."""

    ruta_salida = Path(ruta_salida)
    return ruta_salida.with_name(f".~nuevas_{ruta_salida.name}")


class _Cronometro:
    """Cuanto tarda cada tramo de una escritura del libro (para el log)."""

    def __init__(self, registrar):
        self.registrar = registrar
        self.marca = time.time()
        self.inicio = self.marca

    def tramo(self, nombre):
        ahora = time.time()
        self.registrar(f"  escritura: {nombre} ({_segundos(ahora - self.marca)})")
        self.marca = ahora

    def total(self):
        self.registrar(
            f"  escritura: total {_segundos(time.time() - self.inicio)}"
        )


def _segundos(segundos):
    return f"{segundos:.1f} s".replace(".", ",")


def _pegar_en_salida(writer_ruta, ruta_salida, wb_existente, escritas_vacio):
    """
    Pasa las hojas del libro temporal al de salida (xlsx_rapido) y borra
    el temporal. Si el libro existente es otro archivo que el de salida,
    se parte de una copia suya.
    """

    ruta_salida = Path(ruta_salida)

    try:
        if (
            wb_existente is not None
            and wb_existente.ruta.resolve() != ruta_salida.resolve()
        ):
            shutil.copyfile(wb_existente.ruta, ruta_salida)

        pegar_hojas(
            ruta_salida, writer_ruta, ORDEN_HOJAS_SALIDA,
            ignorar=escritas_vacio,
        )
    finally:
        Path(writer_ruta).unlink(missing_ok=True)


# Una hoja de relleno para el libro temporal cuando esta pasada no
# escribe ninguna (openpyxl no guarda un libro sin hojas). No se pega.
_HOJA_RELLENO = "__sin_hojas__"


def _asegurar_una_hoja(writer):
    if not writer.book.sheetnames:
        writer.book.create_sheet(_HOJA_RELLENO)
        return {_HOJA_RELLENO}
    return set()


def _preservar_ajenas(writer, wb_existente, escritas):
    """
    Copia al libro nuevo toda hoja del anterior que esta pasada no
    haya escrito. Es lo que hace posible que las dos mitades de la
    planilla (entradas y calculo) compartan archivo: la que escribe
    no sabe nada de las hojas de la otra, solo que no las toca.
    """

    if wb_existente is None or isinstance(wb_existente, _LibroEnZip):
        # En el zip, lo que no se escribe ya queda como estaba.
        return

    for nombre in wb_existente.sheetnames:
        if nombre in escritas or nombre in writer.book.sheetnames:
            continue
        _copiar_hoja_existente(wb_existente, nombre, writer.book)


def _ordenar_hojas(wb, orden):
    """
    Deja las hojas del libro en el orden pedido. Lo que no este en
    'orden' queda al final, en el orden en que estaba (una hoja que
    agrego el usuario a mano no se pierde ni se mete al medio).
    """

    posicion = {nombre: i for i, nombre in enumerate(orden)}
    wb._sheets.sort(key=lambda ws: posicion.get(ws.title, len(posicion)))


# Fila (1-indexada) donde estan los NOMBRES de columna de cada hoja,
# para formatear_libro(): las dos hojas de calculo llevan arriba la
# fila de encabezados de grupo, PRORRATA_RETIROS el titulo de la hoja
# y el de cada cuadro, y "Ofertas SSCC" el titulo de cada tabla. Es
# una sola tabla para todo el libro (las dos mitades se formatean
# juntas: cualquiera de las dos escrituras formatea tambien las hojas
# preservadas de la otra).
_FILAS_ENCABEZADO = {
    HOJA_CALCULO_ECOSTOS: 3,
    HOJA_CALCULO_RE545: 2,
    HOJA_PRORRATA_RETIROS: 3,
    HOJA_COMPENSACION_CENTRAL: 3,
    HOJA_OFERTAS_SSCC: 3,
    HOJA_CENTRALES: 3,
    HOJA_FMA: 3,
    HOJA_SUBASTAS: 3,
    HOJA_RESUMEN: 3,
    HOJA_MEDIDORES: 3,
    HOJA_CMG: 3,
    HOJA_FD: 3,
}

# Hojas con varios cuadros lado a lado y el titulo de cada uno en esta
# fila: va combinado sobre su cuadro, sin fondo y en el azul de los
# encabezados (la fila 1 queda vacia). Ver formato.formatear_hoja.
_FILAS_TITULO_CUADRO = {
    HOJA_CALCULO_ECOSTOS: 2,
    HOJA_PRORRATA_RETIROS: 2,
    HOJA_COMPENSACION_CENTRAL: 2,
    HOJA_FMA: 2,
    HOJA_SUBASTAS: 2,
    HOJA_RESUMEN: 2,
    HOJA_MEDIDORES: 2,
    HOJA_CMG: 2,
    HOJA_FD: 2,
    HOJA_OFERTAS_SSCC: 2,
    HOJA_CENTRALES: 2,
}

# Comentario con la formula en el nombre de la columna de plata de los
# cuadros 2 y 3 de PRORRATA_RETIROS (clave: la columna donde arranca el
# cuadro, como en la lista de cuadros de escribir_pagos_bess).
_NOTAS_PRORRATA = {
    4: (
        "Pago(e,q) = M(q) × p(e,q) / Σ_k p(k,q)\n\n"
        "M(q): monto a compensar del cuarto de hora q (Cuadro N° 1)\n"
        "p(e,q): prorrata de la empresa e en el cuarto q\n"
        "Σ_k p(k,q): suma de las prorratas del cuarto q (= 1 si viene "
        "normalizada)"
    ),
    9: (
        "Total(e) = Σ_q Pago(e,q)\n\n"
        "Suma de lo que paga la empresa e en todos los cuartos de hora "
        "del mes (Cuadro N° 2)"
    ),
}


def _comentario(texto):
    from openpyxl.comments import Comment

    comentario = Comment(texto, "Balance_SSCC")
    comentario.width = 320
    comentario.height = 120
    return comentario



def _escribir_encabezados_grupo(ws, columnas_internas, grupos, fila=1, columna_inicio=1):
    """
    Escribe, en una fila propia arriba de los nombres de columna, los
    titulos de grupo con celdas combinadas que trae el archivo real
    (ej. "Subastas"/"FD"/"FMA" arriba de las reservas de Calculo
    RE545, "Prorratas (-)"/"Prorratas (+)" en Calculo E Costos).

    columnas_internas: el orden de claves internas con el que se armo
    la hoja ANTES de renombrar (list(NOMBRES_CALCULO_XXX) -- define
    la posicion real de cada grupo en la salida, que es la del orden
    interno, no la letra del archivo real: la salida no reproduce la
    letra de Excel real, solo el orden y el contenido, asi que un
    grupo puede terminar en una letra distinta de la del archivo
    original).
    grupos: tuplas (etiqueta, [claves internas del grupo, en orden]).
    fila: fila de Excel (1-indexada) donde va el titulo de grupo --
    la fila de nombres de columna queda siempre una fila mas abajo.
    columna_inicio: columna de Excel (1-indexada) donde arranca la
    hoja (1 = A, para el bloque principal; mas adelante para una
    tabla escrita al lado, si alguna vez tuviera sus propios grupos).
    """

    for etiqueta, claves in grupos:
        posiciones = [columnas_internas.index(clave) for clave in claves]
        columna_desde = columna_inicio + min(posiciones)
        columna_hasta = columna_inicio + max(posiciones)

        if columna_hasta > columna_desde:
            ws.merge_cells(
                start_row=fila, start_column=columna_desde,
                end_row=fila, end_column=columna_hasta,
            )

        ws.cell(row=fila, column=columna_desde, value=etiqueta)


def escribir_control(
    ruta_control,
    df_log=None,
    registro=None,
    manifiesto=None,
    conciliacion=None,
    periodo="",
    hojas_regeneradas=(),
):
    """
    Escribe el archivo de control de la corrida (ARCHIVO_CONTROL), que
    a pedido del usuario vive APARTE de la planilla de trabajo:

      - "Ejecucion": el estado de la corrida (APROBADA / NO APROBADA),
        el conteo por severidad, la conciliacion de energia y el
        manifiesto de entradas.
      - "Alertas": una fila por alerta de la corrida, con su id,
        severidad, central y clave.
      - "Log": los avisos e incidencias de la generacion de las hojas
        de entrada.

    Cada mitad de la corrida escribe lo suyo y PRESERVA el resto (la
    corrida de entradas deja el Log sin tocar Alertas/Ejecucion, y al
    reves), igual que las hojas de la planilla.

    OJO con las corridas parciales: si hojas_regeneradas trae una sola
    hoja, las otras vienen de una corrida ANTERIOR y el estado no
    habla de ellas. Por eso "Ejecucion" escribe siempre que hojas se
    recalcularon en esta pasada.

    Se salta sola si no hay nada que escribir: asi las dos funciones de
    escritura siguen sirviendo sueltas (pruebas, scripts) sin armar
    una corrida entera.
    """

    if ruta_control is None or (df_log is None and registro is None):
        return None

    ruta_control = Path(ruta_control)
    wb_existente = _abrir_existente(ruta_control)
    escritas = set()

    with pd.ExcelWriter(ruta_control, engine="openpyxl") as writer:

        if registro is not None:

            filas = [
                ("estado", registro.estado()),
                ("periodo", periodo),
                ("corrida", registro.inicio.strftime("%Y-%m-%d %H:%M:%S")),
                ("hojas_recalculadas", ", ".join(hojas_regeneradas)),
            ]

            for severidad, cuantas in registro.conteo().items():
                filas.append((f"alertas_{severidad.lower()}", cuantas))

            if conciliacion:
                filas += list(conciliacion.items())

            # Mismo formato que la planilla: fila 1 vacia, "Cuadro N°"
            # en la 2, margen en A y una columna entre cuadros.
            _, columna = _escribir_tabla_con_titulo(
                writer, HOJA_EJECUCION,
                pd.DataFrame(filas, columns=["campo", "valor"]),
                titulo_de_cuadro(1, "Estado de la corrida"), fila_inicio=1,
            )
            escritas.add(HOJA_EJECUCION)

            if manifiesto is not None and not manifiesto.empty:
                _escribir_tabla_con_titulo(
                    writer, HOJA_EJECUCION, manifiesto,
                    titulo_de_cuadro(
                        2, "Entradas de esta corrida (sha256 del contenido)"
                    ),
                    fila_inicio=1, columna_inicio=columna,
                )

            _escribir_tabla_con_titulo(
                writer, HOJA_ALERTAS, registro.tabla(),
                titulo_de_cuadro(1, "Alertas de la corrida"), fila_inicio=1,
            )
            escritas.add(HOJA_ALERTAS)

        if df_log is not None:
            _escribir_tabla_con_titulo(
                writer, HOJA_LOG, df_log,
                titulo_de_cuadro(1, "Avisos e incidencias"), fila_inicio=1,
            )
            escritas.add(HOJA_LOG)

        _preservar_ajenas(writer, wb_existente, escritas)
        _ordenar_hojas(writer.book, ORDEN_HOJAS_CONTROL)
        hojas_control = (HOJA_EJECUCION, HOJA_ALERTAS, HOJA_LOG)
        formatear_libro(
            writer.book,
            {hoja: 3 for hoja in hojas_control},
            {hoja: 2 for hoja in hojas_control},
        )

    return ruta_control


def escribir_pagos_bess(
    ruta_salida,
    df_ecostos=None,
    df_re545=None,
    df_resumen_re545=None,
    df_compensacion_detalle=None,
    df_compensacion_central=None,
    df_compensacion_empresa=None,
    df_prorrata_retiros=None,
    df_compensacion_cuarto=None,
    df_pagos_suministrador=None,
    df_resumen=None,
    ruta_existente=None,
    hojas_regenerar=None,
    registrar=print,
    registro=None,
    manifiesto=None,
    conciliacion=None,
    periodo="",
    ruta_control=None,
):
    """
    Escribe la MITAD DE CALCULO de la planilla de salida: "Calculo E
    Costos" (si se pasa df_ecostos), "Calculo RE545" (si se pasa
    df_re545, con la tabla resumen df_resumen_re545 al lado si
    tambien se pasa), COMPENSACION_CENTRAL, PRORRATA_RETIROS y el
    Resumen.

    Es el mismo archivo que escribe escribir_salida() con las hojas de
    entrada (el usuario pidio combinarlos): las hojas de entrada que
    ya estan en el libro se copian tal cual, no se tocan.

    hojas_regenerar: None (por defecto) escribe cada hoja para la que
    se paso su DataFrame, sin mas. Si es un set con alguno de los
    nombres de _HOJAS_PAGOS, la(s) que NO esten en el set se copian
    tal cual desde ruta_existente en vez de escribirse desde el
    DataFrame (una hoja que no se pidio actualizar se preserva, no se
    recalcula). Si una hoja a preservar no existe en ruta_existente,
    queda vacia y se registra un aviso.

    ruta_control: si se pasa, ahi se escribe el control de la corrida
    (Alertas + Ejecucion, ver escribir_control) -- en su propio
    archivo, ya no como hojas de esta planilla.
    """

    ruta_salida = Path(ruta_salida)

    regenerar = (
        set(_HOJAS_PAGOS) if hojas_regenerar is None else set(hojas_regenerar)
    )

    # Siempre se lee el libro que hay: aunque se regeneren TODAS las
    # hojas de calculo, las de entrada (la otra mitad de la planilla)
    # hay que preservarlas igual.
    crono = _Cronometro(registrar)
    wb_existente = _abrir_existente_en_zip(ruta_existente or ruta_salida)
    crono.tramo("abrir el libro existente")
    ruta_temporal = _archivo_temporal(ruta_salida)

    avisos_preservacion = []
    escritas = []

    def _preservar_o_avisar(writer, nombre_hoja):
        if _copiar_hoja_existente(wb_existente, nombre_hoja, writer.book):
            return
        pd.DataFrame().to_excel(writer, sheet_name=nombre_hoja, index=False)
        mensaje = (
            f"No se regenero la hoja '{nombre_hoja}' (no se pidio "
            f"actualizarla en esta corrida) y no habia una version "
            f"anterior para preservarla: quedo vacia. Usa su boton "
            f"'Actualizar' en la ventana."
        )
        avisos_preservacion.append(mensaje)

        if registro is not None:
            registro.anotar(Alerta(
                "PAG-001", ALTA, nombre_hoja, mensaje,
                valor_encontrado="hoja vacia",
                valor_esperado="la hoja de una corrida anterior",
                accion="el libro queda con una hoja de pagos vacia",
                origen_control="CONTROL NUEVO",
            ))

    with pd.ExcelWriter(ruta_temporal, engine="openpyxl") as writer:

        if HOJA_RESUMEN in regenerar:
            if df_resumen is not None:
                escritas.append(HOJA_RESUMEN)
                # Fila 1 vacia, titulo de cuadro en la 2, encabezado en
                # la 3 (como PRORRATA_RETIROS).
                df_resumen.to_excel(
                    writer, sheet_name=HOJA_RESUMEN, index=False,
                    startrow=2, startcol=COLUMNA_MARGEN,
                )
                writer.sheets[HOJA_RESUMEN].cell(
                    2, COLUMNA_MARGEN + 1,
                    titulo_de_cuadro(1, "recibe, paga y neto por empresa"),
                )
        else:
            _preservar_o_avisar(writer, HOJA_RESUMEN)

        if HOJA_CALCULO_ECOSTOS in regenerar:
            if df_ecostos is not None:
                escritas.append(HOJA_CALCULO_ECOSTOS)
                letras_e_costos = _escribir_cuadros_e_costos(
                    writer, df_ecostos
                )
        else:
            _preservar_o_avisar(writer, HOJA_CALCULO_ECOSTOS)

        if HOJA_CALCULO_RE545 in regenerar:
            if df_re545 is not None:
                escritas.append(HOJA_CALCULO_RE545)
                df_re545.to_excel(
                    writer,
                    sheet_name=HOJA_CALCULO_RE545,
                    index=False,
                    startrow=1,
                    startcol=COLUMNA_MARGEN,
                )
                _escribir_encabezados_grupo(
                    writer.sheets[HOJA_CALCULO_RE545],
                    COLUMNAS_SALIDA_RE545,
                    GRUPOS_CALCULO_RE545,
                    columna_inicio=COLUMNA_MARGEN + 1,
                )

                if df_resumen_re545 is not None:
                    # Tabla de otro largo, al lado del bloque
                    # principal con una columna en blanco de
                    # separacion (mismo criterio que CSF/CPF de FD).
                    # startrow=1 tambien, para que sus nombres de
                    # columna queden en la misma fila que los del
                    # bloque principal (no tiene encabezado de grupo
                    # propio: no hay evidencia de uno en el archivo
                    # real).
                    df_resumen_re545.to_excel(
                        writer,
                        sheet_name=HOJA_CALCULO_RE545,
                        index=False,
                        startrow=1,
                        startcol=COLUMNA_MARGEN + len(df_re545.columns) + 1,
                    )
        else:
            _preservar_o_avisar(writer, HOJA_CALCULO_RE545)

        if HOJA_COMPENSACION_CENTRAL in regenerar:
            # Tres cuadros: el detalle por central y ciclo/ventana (E
            # Costos y RE545 lado a lado), el total por central (con su
            # empresa) y el total por empresa, que junta las dos.
            cuadros_compensacion = [
                (df_compensacion_detalle, 1,
                 "Cuadro N° 1 — compensación por central y ciclo/ventana"),
                (df_compensacion_central, 7,
                 "Cuadro N° 2 — total por central"),
                (df_compensacion_empresa, 13,
                 "Cuadro N° 3 — total recibido por cada empresa"),
            ]
            if any(df is not None for df, _, _ in cuadros_compensacion):
                escritas.append(HOJA_COMPENSACION_CENTRAL)
                for df_cuadro, col, titulo in cuadros_compensacion:
                    if df_cuadro is None:
                        continue
                    df_cuadro.to_excel(
                        writer, sheet_name=HOJA_COMPENSACION_CENTRAL,
                        index=False, startrow=2, startcol=col,
                    )
                    ws = writer.sheets[HOJA_COMPENSACION_CENTRAL]
                    ws.cell(2, col + 1, titulo)
        else:
            _preservar_o_avisar(writer, HOJA_COMPENSACION_CENTRAL)

        if HOJA_PRORRATA_RETIROS in regenerar:
            if df_prorrata_retiros is not None or df_compensacion_cuarto is not None:
                escritas.append(HOJA_PRORRATA_RETIROS)
                # Tres cuadros, uno al lado del otro, en el orden en que
                # se leen: primero cuanto hay que compensar en cada
                # cuarto de hora, despues como se reparte ese monto
                # entre las empresas que retiraron en ese cuarto, y al
                # final el total del mes de cada empresa.
                cuadros = [
                    (df_compensacion_cuarto, 1,
                     "Cuadro N° 1 — monto a compensar por cuarto de hora"),
                    (df_prorrata_retiros, 4,
                     "Cuadro N° 2 — reparto del monto del cuarto según la prorrata"),
                    (df_pagos_suministrador, 9,
                     "Cuadro N° 3 — total a pagar de cada empresa"),
                ]
                for df_cuadro, col, titulo in cuadros:
                    if df_cuadro is None:
                        continue
                    df_cuadro.to_excel(
                        writer, sheet_name=HOJA_PRORRATA_RETIROS, index=False,
                        startrow=2, startcol=col,
                    )
                    ws = writer.sheets[HOJA_PRORRATA_RETIROS]
                    ws.cell(2, col + 1, titulo)
                    if col in _NOTAS_PRORRATA:
                        # La formula de la columna de plata, como
                        # comentario en su nombre de columna (I3, M3).
                        celda = ws.cell(3, col + df_cuadro.shape[1])
                        celda.comment = _comentario(_NOTAS_PRORRATA[col])
        else:
            _preservar_o_avisar(writer, HOJA_PRORRATA_RETIROS)

        _preservar_ajenas(writer, wb_existente, set())
        _ordenar_hojas(writer.book, ORDEN_HOJAS_SALIDA)

        crono.tramo("armar las hojas")
        _poner_unidades(writer.book)
        formatear_libro(
            writer.book, _FILAS_ENCABEZADO, _FILAS_TITULO_CUADRO
        )
        if HOJA_CALCULO_ECOSTOS in escritas:
            ws_e_costos = writer.sheets[HOJA_CALCULO_ECOSTOS]
            ws_e_costos.freeze_panes = PANEL_E_COSTOS
            _terminar_e_costos(ws_e_costos, letras_e_costos, len(df_ecostos))
            # Formulas sin valor guardado: que Excel las calcule al abrir.
            writer.book.calculation.fullCalcOnLoad = True
        crono.tramo("formato (encabezados, anchos, numeros, marco)")
        relleno = _asegurar_una_hoja(writer)

    crono.tramo("guardar el archivo temporal")
    _pegar_en_salida(ruta_temporal, ruta_salida, wb_existente, relleno)
    crono.tramo("pegar las hojas en el libro")
    # Cuando se escribio cada hoja (ver sellos.py / vigencia.py).
    sellar(ruta_salida, escritas)

    escribir_control(
        ruta_control,
        registro=registro,
        manifiesto=manifiesto,
        conciliacion=conciliacion,
        periodo=periodo,
        hojas_regeneradas=escritas,
    )
    crono.tramo("Control_corrida.xlsx (Alertas y Ejecucion)")
    crono.total()

    for mensaje in avisos_preservacion:
        registrar(f"  [{ALTA}] PAG-001: {mensaje}")

    registrar(f"Archivo generado: {ruta_salida}")

    return ruta_salida




def _poner_unidades(wb):
    """La unidad de cada columna entre corchetes (ver unidades.py)."""

    for nombre in wb.sheetnames:
        con_unidades(wb[nombre], nombre, _FILAS_ENCABEZADO.get(nombre, 1))


def _escribir_cuadros_e_costos(writer, df_ecostos):
    """
    "Calculo E Costos" en sus tres cuadros (CUADROS_CALCULO_E_COSTOS),
    lado a lado con una columna de separacion: fila 1 vacia, "Cuadro
    N° ..." en la 2, nombres en la 3. df_ecostos llega ya renombrado y
    en el orden de COLUMNAS_SALIDA_E_COSTOS (los cuadros seguidos): se
    parte por posicion. Las columnas que vienen de otra hoja llevan el
    origen como comentario en su nombre (ORIGEN_COLUMNAS_E_COSTOS).
    """

    ws_nombre = HOJA_CALCULO_ECOSTOS
    columna = COLUMNA_MARGEN
    desde = 0
    letras = {}
    for numero, (titulo, claves) in enumerate(CUADROS_CALCULO_E_COSTOS, 1):
        cuadro = df_ecostos.iloc[:, desde:desde + len(claves)]
        _, siguiente = _escribir_tabla_con_titulo(
            writer, ws_nombre, cuadro, titulo_de_cuadro(numero, titulo),
            fila_inicio=1, columna_inicio=columna,
        )
        ws = writer.sheets[ws_nombre]
        for posicion, clave in enumerate(claves):
            nota = ORIGEN_COLUMNAS_E_COSTOS.get(clave)
            if nota:
                ws.cell(3, columna + posicion + 1).comment = _comentario(nota)
            letras[clave] = get_column_letter(columna + posicion + 1)
        desde += len(claves)
        columna = siguiente
    return letras


# El cuadro del Monto a compensar va en verde (es el pago), distinto del
# azul de los demas encabezados.
COLOR_MONTO = "375623"
RELLENO_MONTO_DATOS = "E2EFDA"


def _terminar_e_costos(ws, letras, filas):
    """
    Lo que va DESPUES del formato general en Calculo E Costos (que mira
    los valores para decidir anchos y formatos de numero):

      - formula de Excel en las columnas que se calculan con la misma
        fila (FORMULAS_E_COSTOS), en vez del valor;
      - el total del mes en la fila 1, sobre Total C1, Total C2 y el
        Monto a compensar (TOTALES_E_COSTOS);
      - el cuadro del Monto en verde;
      - un guion en los puestos de las curvas monotonas sin fila.
    """

    from openpyxl.styles import Alignment, Font, PatternFill

    primera, ultima = 4, 3 + filas
    marcador = re.compile(r"\{(\w+)\}")

    for clave, plantilla in FORMULAS_E_COSTOS.items():
        if clave not in letras:
            continue
        for fila in range(primera, ultima + 1):
            ws[f"{letras[clave]}{fila}"].value = marcador.sub(
                lambda m: f"{letras[m.group(1)]}{fila}", plantilla
            )

    derecha = Alignment(horizontal="right")
    for clave in SIN_PUESTO_E_COSTOS:
        if clave not in letras:
            continue
        for fila in range(primera, ultima + 1):
            celda = ws[f"{letras[clave]}{fila}"]
            # pandas deja el vacio como texto vacio, no como None.
            if celda.value is None or celda.value == "":
                celda.value = GUION_SIN_PUESTO
                celda.alignment = derecha

    for clave in TOTALES_E_COSTOS:
        if clave not in letras or not filas:
            continue
        celda = ws[f"{letras[clave]}1"]
        celda.value = f"=SUM({letras[clave]}{primera}:{letras[clave]}{ultima})"
        celda.number_format = '"Total: "#,##0'
        celda.font = Font(
            bold=True, color=COLOR_MONTO if clave == "BC" else "244062"
        )

    monto = letras.get("BC")
    if monto:
        verde = PatternFill(
            fill_type="solid", start_color=COLOR_MONTO, end_color=COLOR_MONTO
        )
        ws[f"{monto}3"].fill = verde
        ws[f"{monto}2"].font = Font(bold=True, color=COLOR_MONTO)
        claro = PatternFill(
            fill_type="solid", start_color=RELLENO_MONTO_DATOS,
            end_color=RELLENO_MONTO_DATOS,
        )
        for fila in range(primera, ultima + 1):
            ws[f"{monto}{fila}"].fill = claro


# Columnas que quedan fijas al moverse a la derecha en Calculo E Costos
# (pedido del usuario: "que se inmovilicen C:H", Dia..Configuracion):
# el panel se corta en la I, debajo de los nombres de columna.
PANEL_E_COSTOS = "I4"


def _central_primero(df):
    """El bloque FD con la columna 'Central' al inicio (si la trae)."""

    if df is None or COLUMNA_CENTRAL not in df.columns:
        return df
    orden = [COLUMNA_CENTRAL] + [c for c in df.columns if c != COLUMNA_CENTRAL]
    # Por posicion: el bloque puede traer nombres repetidos.
    posiciones = [list(df.columns).index(COLUMNA_CENTRAL)] + [
        i for i, c in enumerate(df.columns) if c != COLUMNA_CENTRAL
    ]
    return df.iloc[:, posiciones].set_axis(orden, axis=1)


# Columnas de la hoja FD que van alineadas a la izquierda y con dos
# decimales (texto y numeros, pedido del usuario): las de respuesta.
_PREFIJO_RESPUESTA_FD = "Respuesta"


def _respuestas_fd_a_la_izquierda(ws, fila_encabezado=3):
    from openpyxl.styles import Alignment

    columnas = [
        celda.column for celda in ws[fila_encabezado]
        if str(celda.value or "").startswith(_PREFIJO_RESPUESTA_FD)
    ]
    izquierda = Alignment(horizontal="left")
    for columna in columnas:
        for (celda,) in ws.iter_rows(
            min_row=fila_encabezado + 1, min_col=columna, max_col=columna
        ):
            if celda.value is None:
                continue
            celda.alignment = izquierda
            celda.number_format = "0.00"


# Toda hoja de la planilla deja la columna A vacia, de margen (pedido
# del usuario): el primer cuadro empieza en la B y entre un cuadro y el
# siguiente hay UNA columna vacia.
COLUMNA_MARGEN = 1


def _escribir_tabla_con_titulo(
    writer, hoja, df, titulo, fila_inicio=0, columna_inicio=COLUMNA_MARGEN
):
    """
    Escribe 'titulo' en una fila y 'df' (con su propio encabezado)
    justo debajo, dentro de la hoja 'hoja', empezando en fila_inicio y
    columna_inicio.

    Devuelve (fila_siguiente, columna_siguiente): donde empezaria el
    proximo bloque si se apila debajo, o si se pone al lado,
    respectivamente (cada llamada solo usa el que necesite).
    """

    pd.DataFrame([[titulo]]).to_excel(
        writer,
        sheet_name=hoja,
        index=False,
        header=False,
        startrow=fila_inicio,
        startcol=columna_inicio,
    )

    df.to_excel(
        writer,
        sheet_name=hoja,
        index=False,
        startrow=fila_inicio + 1,
        startcol=columna_inicio,
    )

    celda_titulo = writer.sheets[hoja].cell(
        row=fila_inicio + 1, column=columna_inicio + 1
    )
    celda_titulo.font = Font(bold=True)

    # Debajo: +1 titulo, +1 encabezado de df, +len(df) filas, +2 de separacion.
    fila_siguiente = fila_inicio + len(df) + 4

    # Al lado: +len(df.columns) del bloque, +1 columna de separacion.
    columna_siguiente = columna_inicio + len(df.columns) + 1

    return fila_siguiente, columna_siguiente


# Columna Q en indice 0 (A=0): usada para ubicar el bloque CPF de FD
# a la derecha del bloque CSF, en la misma hoja.
_COLUMNA_Q_INDICE = 16


def _copiar_hoja_existente(wb_origen, nombre_hoja, wb_destino):
    """
    Copia una hoja (solo valores, sin formulas ni formato) de un
    workbook openpyxl a otro. La usan escribir_salida() y
    escribir_pagos_bess() para preservar una hoja que no se pidio
    actualizar en esta corrida (ver hojas_regenerar). Devuelve False
    si wb_origen es None o no tiene esa hoja (no hay nada que
    preservar).

    Las celdas COMBINADAS si se copian, aunque el resto del formato
    no: son los encabezados de grupo de Pagos_BESS.xlsx
    (_escribir_encabezados_grupo), y sin esto cada vez que se
    actualiza UNA de las dos hojas la otra se preservaria con el
    texto del encabezado pero sin la combinacion. Con un boton
    "Actualizar" por hoja, preservar es el caso normal, no la
    excepcion.
    """

    if wb_origen is None or nombre_hoja not in wb_origen.sheetnames:
        return False

    if isinstance(wb_origen, _LibroEnZip):
        # Queda en el zip tal cual: no hay nada que copiar.
        return True

    hoja_o = wb_origen[nombre_hoja]
    hoja_d = wb_destino.create_sheet(title=nombre_hoja)

    for fila in hoja_o.iter_rows():
        for celda in fila:
            hoja_d.cell(
                row=celda.row, column=celda.column, value=celda.value
            )

    for rango in hoja_o.merged_cells.ranges:
        hoja_d.merge_cells(str(rango))

    for letra, dim in hoja_o.column_dimensions.items():
        if dim.width:
            hoja_d.column_dimensions[letra].width = dim.width

    return True


_HOJAS_CONSOLIDADO = (
    HOJA_MEDIDORES, HOJA_OFERTAS_SSCC, HOJA_CMG, HOJA_FD, HOJA_SUBASTAS,
    HOJA_CENTRALES, HOJA_FMA,
)


def _escribir_hoja_centrales(writer, resumen, diccionario):
    """
    La hoja "Centrales": 'Resumen BESS' con su titulo y sus
    encabezados, y al lado la hoja 'Diccionario' tal cual (celda por
    celda, sin encabezado propio). El formato y el porque estan en
    lectura.py, junto a leer_centrales_planilla(), que es la que la
    vuelve a leer: las dos puntas se mueven juntas.
    """

    # 'Energía mínima' no se usa en ningun calculo (pedido del usuario:
    # "quitémosla"). Todo lo de Resumen BESS se busca por nombre de
    # columna, asi que sacarla no corre nada.
    resumen = resumen.loc[:, [
        columna for columna in resumen.columns
        if normalizar(columna) != "energia minima"
    ]]

    # Fila 1 vacia, "Cuadro N° ..." en la 2 y encabezados en la 3.
    _, columna = _escribir_tabla_con_titulo(
        writer, HOJA_CENTRALES, resumen, TITULO_CENTRALES_RESUMEN,
        fila_inicio=1,
    )

    pd.DataFrame([[TITULO_CENTRALES_DICCIONARIO]]).to_excel(
        writer, sheet_name=HOJA_CENTRALES, index=False, header=False,
        startrow=1, startcol=columna,
    )

    # El Diccionario real empieza con una fila vacia (sus encabezados
    # van en la 2): sin ella, los encabezados quedan en la fila de
    # encabezados de la hoja y no como primera fila de datos. Quitar
    # filas vacias de arriba no cambia ninguna lectura (la fila de
    # encabezados del Diccionario se busca, ver encabezado_diccionario).
    con_algo = diccionario.notna().any(axis=1).to_numpy()
    primera = int(con_algo.argmax()) if con_algo.any() else 0
    diccionario.iloc[primera:].to_excel(
        writer, sheet_name=HOJA_CENTRALES, index=False, header=False,
        startrow=2, startcol=columna,
    )


def escribir_salida(
    df,
    ruta_salida,
    avisos,
    incidencias,
    df_wxy=None,
    df_resumen_ventana=None,
    df_cmg=None,
    df_fd_csf=None,
    df_fd_cpf=None,
    df_subastas=None,
    ruta_existente=None,
    hojas_regenerar=None,
    registrar=print,
    ruta_control=None,
    centrales=None,
    tablas_fma=None,
):
    """
    Escribe la MITAD DE ENTRADAS de la planilla de salida: la tabla
    Medidores (A:U, una fila por registro), las tablas auxiliares de
    Ofertas SSCC -de otro largo, ver comentario de LETRA_A_CAMPO-
    juntas en una misma hoja (HOJA_OFERTAS_SSCC, una debajo de la
    otra), CMg, FD (el bloque CSF y el bloque CPF lado a lado, de
    distinto largo cada uno - ver construir_fd) y Subastas.

    Es el mismo archivo que escribe escribir_pagos_bess() con las
    hojas de calculo (el usuario pidio combinarlos): esas hojas se
    copian tal cual, no se tocan.

    hojas_regenerar: None (por defecto) regenera las 5 hojas de datos
    con lo que se haya pasado. Si es un set con algunos nombres de
    _HOJAS_CONSOLIDADO, las que NO esten en el set se copian tal cual
    desde ruta_existente en vez de recalcularse -- lo usa
    generar_consolidado() cuando el usuario destilda una entrada en
    su boton "Actualizar". Si una hoja a preservar no existe en
    ruta_existente, queda vacia y se registra un aviso (en el log de
    esta corrida y como fila del Log).

    ruta_control: si se pasa, ahi se escribe la hoja "Log" (avisos e
    incidencias) -- en el archivo de control, ya no como una hoja mas
    de esta planilla.

    centrales: (resumen_bess, diccionario) de Centrales.xlsx, para la
    hoja "Centrales" (ver _escribir_hoja_centrales).

    tablas_fma: el dict de cargar_tablas_fma(), para la hoja "FMA" (una
    tabla al lado de la otra, cada una con su titulo de TITULOS_FMA).
    """

    ruta_salida = Path(ruta_salida)

    regenerar = (
        set(_HOJAS_CONSOLIDADO)
        if hojas_regenerar is None
        else set(hojas_regenerar)
    )

    # Siempre se lee el libro que hay: aunque se regeneren TODAS las
    # hojas de entrada, las de calculo (la otra mitad de la planilla)
    # hay que preservarlas igual.
    crono = _Cronometro(registrar)
    wb_existente = _abrir_existente_en_zip(ruta_existente or ruta_salida)
    crono.tramo("abrir el libro existente")
    ruta_temporal = _archivo_temporal(ruta_salida)

    avisos_preservacion = []

    def _preservar_o_avisar(writer, nombre_hoja):
        if _copiar_hoja_existente(wb_existente, nombre_hoja, writer.book):
            return
        pd.DataFrame().to_excel(writer, sheet_name=nombre_hoja, index=False)
        mensaje = (
            f"No se regenero la hoja '{nombre_hoja}' (no se pidio "
            f"actualizarla en esta corrida) y no habia una version "
            f"anterior para preservarla: quedo vacia. Usa su boton "
            f"'Actualizar' en la ventana."
        )
        avisos_preservacion.append(mensaje)
        registrar(f"  [AVISO] {mensaje}")

    # El Log se arma recien al final: los avisos de preservacion los
    # va agregando _preservar_o_avisar() MIENTRAS se escriben las
    # hojas, asi que construirlo antes (como estaba) los perdia todos.
    # Con el desglose hoja por hoja de la ventana, generar solo una
    # hoja es el caso normal y esos avisos son justamente los que hay
    # que ver.
    def _armar_log():
        registros = (
            [("aviso", a) for a in avisos]
            + [("aviso", a) for a in avisos_preservacion]
            + [("incidencia_soc", i) for i in incidencias]
        )
        return pd.DataFrame(
            registros or [("ok", "Sin observaciones.")],
            columns=["tipo", "detalle"],
        )

    with pd.ExcelWriter(ruta_temporal, engine="openpyxl") as writer:

        if HOJA_MEDIDORES in regenerar:
            _escribir_tabla_con_titulo(
                writer, HOJA_MEDIDORES, df,
                titulo_de_cuadro(1, "Medidores"), fila_inicio=1,
            )
        else:
            _preservar_o_avisar(writer, HOJA_MEDIDORES)

        if HOJA_OFERTAS_SSCC in regenerar:

            columna = COLUMNA_MARGEN

            if df_wxy is not None:
                _, columna = _escribir_tabla_con_titulo(
                    writer,
                    HOJA_OFERTAS_SSCC,
                    df_wxy,
                    TITULO_OFERTAS_POR_DIA,
                    fila_inicio=1,
                    columna_inicio=columna,
                )

            if df_resumen_ventana is not None:
                _escribir_tabla_con_titulo(
                    writer,
                    HOJA_OFERTAS_SSCC,
                    df_resumen_ventana,
                    TITULO_RESUMEN_VENTANA,
                    fila_inicio=1,
                    columna_inicio=columna,
                )
        else:
            _preservar_o_avisar(writer, HOJA_OFERTAS_SSCC)

        if HOJA_CMG in regenerar:
            if df_cmg is not None:
                _escribir_tabla_con_titulo(
                    writer, HOJA_CMG, df_cmg,
                    titulo_de_cuadro(1, "Costo marginal por barra"),
                    fila_inicio=1,
                )
        else:
            _preservar_o_avisar(writer, HOJA_CMG)

        if HOJA_FD in regenerar:

            # Sin las auxiliares (id y la copia de Hora Mes, ver
            # fd_para_hoja) y con una sola columna vacia entre CSF y CPF.
            # 'Central' va primera en cada bloque (la lectura la vuelve
            # a dejar al final). Fila 1 vacia, "Cuadro N° ..." en la 2.
            hoja_csf = _central_primero(fd_para_hoja(df_fd_csf))
            hoja_cpf = _central_primero(fd_para_hoja(df_fd_cpf))

            if hoja_csf is not None:
                _escribir_tabla_con_titulo(
                    writer, HOJA_FD, hoja_csf,
                    titulo_de_cuadro(1, "FD CSF"), fila_inicio=1,
                )

            if hoja_cpf is not None:
                _escribir_tabla_con_titulo(
                    writer, HOJA_FD, hoja_cpf,
                    titulo_de_cuadro(2 if hoja_csf is not None else 1,
                                     "FD CPF"),
                    fila_inicio=1,
                    columna_inicio=COLUMNA_MARGEN + (
                        len(hoja_csf.columns) + 1 if hoja_csf is not None
                        else 0
                    ),
                )
        else:
            _preservar_o_avisar(writer, HOJA_FD)

        if HOJA_SUBASTAS in regenerar:
            if df_subastas is not None:
                # Fila 1 vacia, titulo de cuadro en la 2, encabezado en
                # la 3 (como PRORRATA_RETIROS).
                subastas_para_hoja(df_subastas).to_excel(
                    writer,
                    sheet_name=HOJA_SUBASTAS,
                    index=False,
                    startrow=2,
                    startcol=COLUMNA_MARGEN,
                )
                writer.sheets[HOJA_SUBASTAS].cell(
                    2, COLUMNA_MARGEN + 1,
                    titulo_de_cuadro(1, "subastas adjudicadas del periodo"),
                )
        else:
            _preservar_o_avisar(writer, HOJA_SUBASTAS)

        if HOJA_FMA in regenerar:
            if tablas_fma is not None:
                columna = COLUMNA_MARGEN
                for clave, titulo in TITULOS_FMA.items():
                    tabla = tablas_fma.get(clave)
                    # Fila 1 vacia, "Cuadro N° ..." en la 2 (como
                    # PRORRATA_RETIROS), encabezado en la 3.
                    _, columna = _escribir_tabla_con_titulo(
                        writer, HOJA_FMA,
                        tabla if tabla is not None else pd.DataFrame(),
                        titulo, fila_inicio=1, columna_inicio=columna,
                    )
            else:
                _copiar_hoja_existente(wb_existente, HOJA_FMA, writer.book)
        else:
            _preservar_o_avisar(writer, HOJA_FMA)

        if HOJA_CENTRALES in regenerar and centrales is not None:
            _escribir_hoja_centrales(writer, *centrales)
        else:
            # Es la unica hoja de entrada que NO se deja vacia si falta:
            # un libro de antes de que existiera sigue andando con
            # Centrales.xlsx (ver obtener_centrales), asi que no hay
            # nada que avisar.
            _copiar_hoja_existente(wb_existente, HOJA_CENTRALES, writer.book)

        _preservar_ajenas(writer, wb_existente, set())
        _ordenar_hojas(writer.book, ORDEN_HOJAS_SALIDA)

        crono.tramo("armar las hojas")
        _poner_unidades(writer.book)
        formatear_libro(
            writer.book, _FILAS_ENCABEZADO, _FILAS_TITULO_CUADRO
        )
        if HOJA_FD in regenerar and HOJA_FD in writer.sheets:
            _respuestas_fd_a_la_izquierda(writer.sheets[HOJA_FD])
        crono.tramo("formato (encabezados, anchos, numeros, marco)")
        relleno = _asegurar_una_hoja(writer)

    crono.tramo("guardar el archivo temporal")
    _pegar_en_salida(ruta_temporal, ruta_salida, wb_existente, relleno)
    crono.tramo("pegar las hojas en el libro")
    crono.total()
    # Cuando se escribio cada hoja (ver sellos.py / vigencia.py).
    sellar(ruta_salida, [
        hoja for hoja, escrita in (
            (HOJA_MEDIDORES, df is not None),
            (HOJA_OFERTAS_SSCC,
             df_wxy is not None or df_resumen_ventana is not None),
            (HOJA_CMG, df_cmg is not None),
            (HOJA_FD, df_fd_csf is not None or df_fd_cpf is not None),
            (HOJA_SUBASTAS, df_subastas is not None),
            (HOJA_FMA, tablas_fma is not None),
            (HOJA_CENTRALES, centrales is not None),
        )
        if escrita and hoja in regenerar
    ])

    escribir_control(ruta_control, df_log=_armar_log())

    return ruta_salida
