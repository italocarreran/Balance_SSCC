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

from pathlib import Path

import openpyxl
import pandas as pd
from openpyxl.styles import Font

from .alertas import ALTA, Alerta

from .ecostos import COLUMNAS_SALIDA_E_COSTOS, GRUPOS_CALCULO_E_COSTOS
from .formato import formatear_libro
from .ofertas_sscc import (
    HOJA_OFERTAS_SSCC, TITULO_OFERTAS_POR_DIA, TITULO_RESUMEN_VENTANA,
)
from .parametros import (
    HOJA_ALERTAS, HOJA_CALCULO_ECOSTOS, HOJA_CALCULO_RE545, HOJA_CMG,
    HOJA_COMPENSACION_CENTRAL, HOJA_EJECUCION, HOJA_FD, HOJA_LOG,
    HOJA_MEDIDORES, HOJA_PRORRATA_RETIROS, HOJA_RESUMEN, HOJA_SUBASTAS,
    ORDEN_HOJAS_CONTROL, ORDEN_HOJAS_SALIDA,
)
from .re545 import COLUMNAS_SALIDA_RE545, GRUPOS_CALCULO_RE545


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


def _preservar_ajenas(writer, wb_existente, escritas):
    """
    Copia al libro nuevo toda hoja del anterior que esta pasada no
    haya escrito. Es lo que hace posible que las dos mitades de la
    planilla (entradas y calculo) compartan archivo: la que escribe
    no sabe nada de las hojas de la otra, solo que no las toca.
    """

    if wb_existente is None:
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
    HOJA_CALCULO_ECOSTOS: 2,
    HOJA_CALCULO_RE545: 2,
    HOJA_PRORRATA_RETIROS: 3,
    HOJA_OFERTAS_SSCC: 2,
}


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

            pd.DataFrame(filas, columns=["campo", "valor"]).to_excel(
                writer, sheet_name=HOJA_EJECUCION, index=False
            )
            escritas.add(HOJA_EJECUCION)

            if manifiesto is not None and not manifiesto.empty:
                _escribir_tabla_con_titulo(
                    writer, HOJA_EJECUCION, manifiesto,
                    "Entradas de esta corrida (sha256 del contenido)",
                    fila_inicio=len(filas) + 3,
                )

            registro.tabla().to_excel(
                writer, sheet_name=HOJA_ALERTAS, index=False
            )
            escritas.add(HOJA_ALERTAS)

        if df_log is not None:
            df_log.to_excel(writer, sheet_name=HOJA_LOG, index=False)
            escritas.add(HOJA_LOG)

        _preservar_ajenas(writer, wb_existente, escritas)
        _ordenar_hojas(writer.book, ORDEN_HOJAS_CONTROL)
        formatear_libro(writer.book)

    return ruta_control


def escribir_pagos_bess(
    ruta_salida,
    df_ecostos=None,
    df_re545=None,
    df_resumen_re545=None,
    df_compensacion_ecostos=None,
    df_compensacion_re545=None,
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
    wb_existente = _abrir_existente(ruta_existente or ruta_salida)

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

    with pd.ExcelWriter(ruta_salida, engine="openpyxl") as writer:

        if HOJA_RESUMEN in regenerar:
            if df_resumen is not None:
                escritas.append(HOJA_RESUMEN)
                df_resumen.to_excel(writer, sheet_name=HOJA_RESUMEN, index=False)
        else:
            _preservar_o_avisar(writer, HOJA_RESUMEN)

        if HOJA_CALCULO_ECOSTOS in regenerar:
            if df_ecostos is not None:
                escritas.append(HOJA_CALCULO_ECOSTOS)
                # startrow=1: deja la fila 1 libre para los
                # encabezados de grupo (celdas combinadas), que se
                # escriben aparte con _escribir_encabezados_grupo();
                # los nombres de columna quedan en la fila 2 y los
                # datos desde la fila 3.
                df_ecostos.to_excel(
                    writer,
                    sheet_name=HOJA_CALCULO_ECOSTOS,
                    index=False,
                    startrow=1,
                )
                _escribir_encabezados_grupo(
                    writer.sheets[HOJA_CALCULO_ECOSTOS],
                    COLUMNAS_SALIDA_E_COSTOS,
                    GRUPOS_CALCULO_E_COSTOS,
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
                )
                _escribir_encabezados_grupo(
                    writer.sheets[HOJA_CALCULO_RE545],
                    COLUMNAS_SALIDA_RE545,
                    GRUPOS_CALCULO_RE545,
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
                        startcol=len(df_re545.columns) + 1,
                    )
        else:
            _preservar_o_avisar(writer, HOJA_CALCULO_RE545)

        if HOJA_COMPENSACION_CENTRAL in regenerar:
            cuadros_compensacion = [
                (df_compensacion_ecostos, 1,
                 "Cuadro N° 1 — Calculo E Costos: compensación por central "
                 "y ciclo"),
                (df_compensacion_re545, 7,
                 "Cuadro N° 2 — Calculo RE545: compensación por central "
                 "y ventana"),
                (df_compensacion_empresa, 13,
                 "Cuadro N° 3 — total recibido por cada empresa"),
            ]
            if any(df is not None for df, _, _ in cuadros_compensacion):
                escritas.append(HOJA_COMPENSACION_CENTRAL)
                primero = True
                for df_cuadro, col, titulo in cuadros_compensacion:
                    if df_cuadro is None:
                        continue
                    df_cuadro.to_excel(
                        writer, sheet_name=HOJA_COMPENSACION_CENTRAL,
                        index=False, startrow=2, startcol=col,
                    )
                    ws = writer.sheets[HOJA_COMPENSACION_CENTRAL]
                    if primero:
                        ws.cell(
                            1, 2,
                            "Compensación por central, ciclo/ventana y empresa",
                        )
                        primero = False
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
                    (df_prorrata_retiros, 5,
                     "Cuadro N° 2 — reparto del monto del cuarto según la prorrata"),
                    (df_pagos_suministrador, 11,
                     "Cuadro N° 3 — total a pagar de cada empresa"),
                ]
                primero = True
                for df_cuadro, col, titulo in cuadros:
                    if df_cuadro is None:
                        continue
                    df_cuadro.to_excel(
                        writer, sheet_name=HOJA_PRORRATA_RETIROS, index=False,
                        startrow=2, startcol=col,
                    )
                    ws = writer.sheets[HOJA_PRORRATA_RETIROS]
                    if primero:
                        ws.cell(
                            1, 2,
                            "Prorrata de retiro y cálculo de asignación de pagos",
                        )
                        primero = False
                    ws.cell(2, col + 1, titulo)
        else:
            _preservar_o_avisar(writer, HOJA_PRORRATA_RETIROS)

        _preservar_ajenas(writer, wb_existente, set())
        _ordenar_hojas(writer.book, ORDEN_HOJAS_SALIDA)

        formatear_libro(writer.book, _FILAS_ENCABEZADO)

    escribir_control(
        ruta_control,
        registro=registro,
        manifiesto=manifiesto,
        conciliacion=conciliacion,
        periodo=periodo,
        hojas_regeneradas=escritas,
    )

    for mensaje in avisos_preservacion:
        registrar(f"  [{ALTA}] PAG-001: {mensaje}")

    registrar(f"Archivo generado: {ruta_salida}")

    return ruta_salida




def _escribir_tabla_con_titulo(
    writer, hoja, df, titulo, fila_inicio=0, columna_inicio=0
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

    # Al lado: +len(df.columns) del bloque, +2 columnas de separacion.
    columna_siguiente = columna_inicio + len(df.columns) + 2

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
    wb_existente = _abrir_existente(ruta_existente or ruta_salida)

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

    with pd.ExcelWriter(ruta_salida, engine="openpyxl") as writer:

        if HOJA_MEDIDORES in regenerar:
            df.to_excel(
                writer,
                sheet_name=HOJA_MEDIDORES,
                index=False,
            )
        else:
            _preservar_o_avisar(writer, HOJA_MEDIDORES)

        if HOJA_OFERTAS_SSCC in regenerar:

            columna = 0

            if df_wxy is not None:
                _, columna = _escribir_tabla_con_titulo(
                    writer,
                    HOJA_OFERTAS_SSCC,
                    df_wxy,
                    TITULO_OFERTAS_POR_DIA,
                    columna_inicio=columna,
                )

            if df_resumen_ventana is not None:
                _escribir_tabla_con_titulo(
                    writer,
                    HOJA_OFERTAS_SSCC,
                    df_resumen_ventana,
                    TITULO_RESUMEN_VENTANA,
                    columna_inicio=columna,
                )
        else:
            _preservar_o_avisar(writer, HOJA_OFERTAS_SSCC)

        if HOJA_CMG in regenerar:
            if df_cmg is not None:
                df_cmg.to_excel(
                    writer,
                    sheet_name=HOJA_CMG,
                    index=False,
                )
        else:
            _preservar_o_avisar(writer, HOJA_CMG)

        if HOJA_FD in regenerar:

            if df_fd_csf is not None:
                df_fd_csf.to_excel(
                    writer,
                    sheet_name=HOJA_FD,
                    index=False,
                    startcol=0,
                )

            if df_fd_cpf is not None:
                df_fd_cpf.to_excel(
                    writer,
                    sheet_name=HOJA_FD,
                    index=False,
                    startcol=_COLUMNA_Q_INDICE,
                )
        else:
            _preservar_o_avisar(writer, HOJA_FD)

        if HOJA_SUBASTAS in regenerar:
            if df_subastas is not None:
                df_subastas.to_excel(
                    writer,
                    sheet_name=HOJA_SUBASTAS,
                    index=False,
                )
        else:
            _preservar_o_avisar(writer, HOJA_SUBASTAS)

        _preservar_ajenas(writer, wb_existente, set())
        _ordenar_hojas(writer.book, ORDEN_HOJAS_SALIDA)

        formatear_libro(writer.book, _FILAS_ENCABEZADO)

    escribir_control(ruta_control, df_log=_armar_log())

    return ruta_salida
