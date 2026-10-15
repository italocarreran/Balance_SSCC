# -*- coding: utf-8 -*-
"""
Lo que quedo viejo: si cambio algo de lo que una fila usa, esa fila (y
todo lo que sale de ella) pasa de "ok" a "pendiente".

Pedido del usuario: *"si yo, por ejemplo, modifico las medidas, todo lo
que esta ok pase a pendiente, si es que depende de medidas. Asi con
todo."*

Las dependencias son las mismas del grafo de "Ejecutar todo"
(orquestador.TAREAS: depende, depende_opcional y requisitos), asi que
el arbol y el plan nunca dicen cosas distintas. La fecha de cada fila:

  - un archivo: su fecha de modificacion;
  - DB subastas/: el Access del periodo mas nuevo;
  - una hoja de Balance_BESS.xlsx: cuando se escribio (sellos.py). Una
    hoja sin sello (de un libro de antes de este cambio) queda sin
    fecha y no avisa nada hasta que se vuelva a generar: la fecha del
    libro no sirve, porque cambia cada vez que se escribe CUALQUIER
    hoja y haria parecer nueva una hoja vieja;
  - una sub-fila de un archivo (hoja 'Resumen BESS' de Centrales.xlsx):
    la del archivo.

Una fila queda vieja si algo de lo que usa es mas nuevo que ella, o si
algo de lo que usa quedo viejo (en cadena). Solo se tocan filas en
"ok": lo que ya estaba pendiente o faltaba sigue igual.
"""

from pathlib import Path

from .sellos import leer_sellos

# Margen para no marcar como vieja una hoja escrita en el mismo segundo
# que su entrada (relojes y sistemas de archivo redondean distinto).
MARGEN_SEGUNDOS = 2.0


def _mtime(ruta):
    try:
        return Path(ruta).stat().st_mtime
    except (OSError, TypeError):
        return None


def _hojas_de_fila():
    """'consolidado:medidores' -> ('Medidores',), etc."""

    from .estructura import SECCIONES_SALIDA

    return {
        f"{grupo}:{seccion}": hojas
        for grupo, (seccion, _, _, hojas) in SECCIONES_SALIDA
    }


def _fechas(rutas, filas):
    """{id de fila: instante} para las filas que tienen fecha."""

    por_id = {fila["id"]: fila for fila in filas}
    libro = rutas.get("salida")
    sellos = leer_sellos(libro) if libro else {}
    hojas_de_fila = _hojas_de_fila()

    fechas = {}
    for fila in filas:
        id_fila = fila["id"]

        # Solo lo que esta: una fila que falta o esta pendiente no tiene
        # fecha propia (su "ruta" puede ser la carpeta donde iria).
        if fila["estado"] != "ok":
            continue

        if id_fila in hojas_de_fila:
            selladas = [
                sellos[hoja] for hoja in hojas_de_fila[id_fila]
                if hoja in sellos
            ]
            fecha = max(selladas) if selladas else None
        elif id_fila == "db_subastas":
            accdb = [_mtime(r) for r in rutas.get("accdb_subastas") or []]
            accdb = [f for f in accdb if f is not None]
            fecha = max(accdb) if accdb else None
        elif fila.get("es_carpeta"):
            fecha = None
        elif fila.get("ruta"):
            fecha = _mtime(fila["ruta"])
        elif ":" in id_fila and id_fila.split(":", 1)[0] in por_id:
            # Hoja de un archivo de entrada (Centrales.xlsx, homologacion).
            padre = por_id[id_fila.split(":", 1)[0]]
            fecha = _mtime(padre.get("ruta")) if not padre.get("es_carpeta") else None
        else:
            fecha = None

        if fecha is not None:
            fechas[id_fila] = fecha

    return fechas


def aplicar_vigencia(rutas, filas):
    """
    Pasa a "pendiente" (con el motivo en el detalle) las filas "ok" que
    quedaron viejas. Modifica `filas` en el lugar y las devuelve.
    """

    from .orquestador import TAREA_POR_ID, TAREAS

    por_id = {fila["id"]: fila for fila in filas}
    fechas = _fechas(rutas, filas)
    viejas = {}

    def usa(tarea):
        """(id de fila, nombre para el motivo) de lo que usa la tarea."""

        for requisito in tarea.requisitos:
            yield requisito
        for dep in tarea.depende + tarea.depende_opcional:
            yield TAREA_POR_ID[dep].fila_estado

    # TAREAS esta en orden de dependencia: lo de arriba ya se resolvio
    # cuando se mira lo de abajo, asi que la cadena sale en una pasada.
    for tarea in TAREAS:
        fila = por_id.get(tarea.fila_estado)
        if fila is None or fila["estado"] != "ok":
            continue

        propia = fechas.get(fila["id"])
        motivo = None

        for id_usado in usa(tarea):
            usada = por_id.get(id_usado)
            if usada is None:
                continue
            nombre = usada.get("etiqueta", id_usado)
            if id_usado in viejas:
                motivo = f"usa {nombre}, que esta desactualizado"
                break
            fecha = fechas.get(id_usado)
            if (
                propia is not None and fecha is not None
                and fecha > propia + MARGEN_SEGUNDOS
            ):
                motivo = f"cambio {nombre} despues de generarlo"
                break

        if motivo:
            viejas[fila["id"]] = motivo
            fila["estado"] = "pendiente"
            fila["desactualizada"] = True
            fila["detalle"] = (
                f"desactualizado: {motivo} - rehacer con el boton ->"
            )

    return filas
