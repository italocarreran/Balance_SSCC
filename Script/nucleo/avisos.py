# -*- coding: utf-8 -*-
"""
Diagnostico comun de los cruces que terminarian en cero sin decirlo.
"""

from .alertas import ALTA, Alerta, anotar_muchas
from .utiles import _texto_seguro, _tiene_valor, normalizar


# ============================================================
# CRUCES SIN CORRESPONDENCIA
# ============================================================

def _avisar_claves_sin_mapeo(
    valores, mapa, descripcion, origen, registrar=print, maximo=15,
    id_alerta="MAE-000", severidad=ALTA, etapa="", archivo="", hoja="",
    accion="el resultado queda vacio o en 0",
    origen_control="CONTROL NUEVO",
):
    """Avisa, una vez por valor, las claves ausentes o con dato vacio.

    Los VLOOKUP/diccionarios de la planilla suelen terminar en blanco y
    varios calculos posteriores convierten ese blanco en cero. Este helper
    hace visible el problema antes de esa conversion, sin inundar el log
    con una linea por cada cuarto de hora.

    A la PANTALLA va una linea con hasta `maximo` ejemplos; al registro
    de la corrida va una alerta por cada valor faltante, sin tope (ver
    FD-005 del catalogo de controles). Devuelve la lista completa de
    faltantes.
    """

    # Las columnas de entrada tienen una fila por cuarto de hora
    # (decenas de miles); los valores distintos son unas pocas
    # centrales. Se deduplica ANTES de normalizar para no repetir el
    # normalizar() -- que hace unicodedata + regex -- una vez por fila.
    distintos = {valor for valor in valores if _tiene_valor(valor)}

    faltantes = sorted({
        _texto_seguro(valor)
        for valor in distintos
        if not _tiene_valor(mapa.get(normalizar(valor)))
    })

    if not faltantes:
        return []

    muestra = ", ".join(repr(valor) for valor in faltantes[:maximo])
    resto = len(faltantes) - maximo
    sufijo = f" (y {resto:,} mas, en la hoja 'Alertas')" if resto > 0 else ""

    alertas = [
        Alerta(
            id_alerta=id_alerta,
            severidad=severidad,
            etapa=etapa,
            central=valor,
            clave=valor,
            mensaje=f"{descripcion}: sin correspondencia en {origen}.",
            valor_esperado=f"una fila en {origen}",
            accion=accion,
            archivo=archivo,
            hoja=hoja,
            origen_control=origen_control,
        )
        for valor in faltantes
    ]

    anotar_muchas(
        registrar,
        alertas,
        f"  [{severidad}] {id_alerta}: {descripcion}: "
        f"{len(faltantes):,} valor(es) sin correspondencia en {origen}: "
        f"{muestra}{sufijo}. Los resultados dependientes pueden quedar "
        f"vacios o en 0.",
    )

    return faltantes
