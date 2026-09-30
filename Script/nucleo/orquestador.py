# -*- coding: utf-8 -*-
"""
"Ejecutar todo": el orden en que se hace cada cosa y que depende de que.

Pedido del usuario: *"Quiero un boton que ejecute todo teniendo la
informacion inicial necesaria (centrales, homologacion y ofertas creo),
que el boton se bloquee si falta algo. [...] que se calcule solo lo
restante pero que yo pueda seleccionar si quiero repetir algun calculo
o generador, etc. Hay cosas que pueden ir en paralelo y cosas que
dependen de otras. Con cuidado"*.

Este modulo NO calcula nada nuevo: llama exactamente a las mismas
funciones que los botones sueltos del diagrama (generar_medidas_sae,
traer_csv_cmg, generar_cmg, traer_fd, generar_fma, traer_subastas,
generar_consolidado, generar_pagos_bess). Lo unico que agrega es:

  - el GRAFO: quien depende de quien, para no correr algo con una
    entrada vieja;
  - el PLAN: que esta al dia (no se rehace), que falta (se hace), que
    no se puede hacer todavia y por que;
  - el PARALELISMO: lo que no depende entre si corre junto -- las
    bajadas de red (medidas, CMg, FD, FMA, subastas) son casi todo el
    tiempo de una corrida y no se estorban.

Tres reglas de las que depende que esto sea seguro:

  1. Dos tareas que escriben el MISMO archivo comparten "recurso" y
     nunca corren a la vez (las hojas del consolidado, por ejemplo, se
     mandan todas juntas en UNA llamada a generar_consolidado, que es
     ademas la unica forma de escribir el libro una sola vez).
  2. Una tarea corre solo si sus dependencias estan al dia o se
     recalculan en esta misma corrida. Si no, queda bloqueada y se
     dice por que, en vez de calcular con datos viejos.
  3. Si una tarea falla, no se corre nada que dependa de ella.
"""

import threading
from collections import namedtuple
from concurrent.futures import FIRST_COMPLETED, ThreadPoolExecutor, wait

from .estructura import revisar_estructura
from .externos import Homologacion
from .parametros import (
    ARCHIVO_CENTRALES, ARCHIVO_CMG, ARCHIVO_MEDIDAS_SAE, ARCHIVO_SALIDA,
    HOJA_COMPENSACION_CENTRAL, HOJA_DICCIONARIO, HOJA_RESUMEN_BESS,
)
from .proceso import generar_consolidado, generar_pagos_bess
from .medidas_sae import generar_medidas_sae
from .rutas import validar_aamm
from .traer import (
    generar_cmg, generar_fma, traer_csv_cmg, traer_fd, traer_subastas,
)
from .utiles import ErrorEntrada


# ============================================================
# EL GRAFO
# ============================================================

Grupo = namedtuple("Grupo", "id etiqueta recurso ejecutar")

Tarea = namedtuple(
    "Tarea",
    "id etiqueta grupo seccion fila_estado requisitos depende "
    "depende_opcional",
)


def _correr_medidas_sae(carpeta_base, aamm, secciones, registrar, progreso):
    return generar_medidas_sae(
        carpeta_base, aamm, registrar=registrar, progreso=progreso
    )


def _correr_traer_cmg(carpeta_base, aamm, secciones, registrar, progreso):
    return traer_csv_cmg(
        carpeta_base, aamm, registrar=registrar, progreso=progreso
    )


def _correr_generar_cmg(carpeta_base, aamm, secciones, registrar, progreso):
    return generar_cmg(
        carpeta_base, aamm, registrar=registrar, progreso=progreso
    )


def _correr_traer_fd(carpeta_base, aamm, secciones, registrar, progreso):
    return traer_fd(
        carpeta_base, aamm, registrar=registrar, progreso=progreso
    )


def _correr_fma(carpeta_base, aamm, secciones, registrar, progreso):
    # Las tres salidas de FMA se arman en UNA llamada: comparten la
    # busqueda de version en el arbol del DCO, asi que pedirlas juntas
    # es una sola pasada por la red en vez de tres.
    return generar_fma(
        carpeta_base, aamm, tipos=set(secciones),
        registrar=registrar, progreso=progreso,
    )


def _correr_traer_subastas(carpeta_base, aamm, secciones, registrar, progreso):
    return traer_subastas(
        carpeta_base, aamm, registrar=registrar, progreso=progreso
    )


def _correr_consolidado(carpeta_base, aamm, secciones, registrar, progreso):
    return generar_consolidado(
        carpeta_base, secciones_activas=set(secciones), aamm=aamm,
        registrar=registrar, progreso=progreso,
    )


def _correr_pagos(carpeta_base, aamm, secciones, registrar, progreso):
    return generar_pagos_bess(
        carpeta_base, secciones_activas=set(secciones), aamm=aamm,
        registrar=registrar, progreso=progreso,
    )


GRUPOS = (
    Grupo("medidas_sae", ARCHIVO_MEDIDAS_SAE, "medidas", _correr_medidas_sae),
    Grupo("cmg_csv", "CSV 15-minutal de CMg", "cmg", _correr_traer_cmg),
    Grupo("cmg_xlsx", ARCHIVO_CMG, "cmg", _correr_generar_cmg),
    Grupo("fd_sscc", "FD (SSCC_Desempeño_*)", "fd_fma", _correr_traer_fd),
    Grupo("fma", "Salidas de FMA", "fd_fma", _correr_fma),
    Grupo("subastas_db", "Access de subastas", "subastas",
          _correr_traer_subastas),
    # Los dos escriben el MISMO archivo (una sola planilla), asi que
    # comparten recurso: nunca pueden correr a la vez.
    Grupo(
        "consolidado", f"{ARCHIVO_SALIDA} (entradas)", "salida",
        _correr_consolidado,
    ),
    Grupo(
        "pagos", f"{ARCHIVO_SALIDA} (calculo)", "salida", _correr_pagos,
    ),
)

GRUPO_POR_ID = {grupo.id: grupo for grupo in GRUPOS}


# Requisitos: ids de fila de revisar_estructura() que tienen que estar
# en "ok" para que la tarea pueda correr. Son las ENTRADAS que pone la
# persona (no las genera el programa), que es justo lo que el usuario
# pidio que bloquee el boton.
_REQ_CENTRALES_RESUMEN = ("centrales", f"centrales:{HOJA_RESUMEN_BESS}")
_REQ_CENTRALES = _REQ_CENTRALES_RESUMEN + (f"centrales:{HOJA_DICCIONARIO}",)
_REQ_HOMOLOGACION = ("homologacion", f"homologacion:{Homologacion.HOJA_HOMOL}")

TAREAS = (
    Tarea(
        "medidas_sae", f"{ARCHIVO_MEDIDAS_SAE} (baja el mes de las APIs)",
        "medidas_sae", None, "medidas_sae",
        _REQ_HOMOLOGACION, (), (),
    ),
    Tarea(
        "cmg_csv", "Traer cmg_15min (unidad de red)",
        "cmg_csv", None, "cmg_csv",
        (), (), (),
    ),
    Tarea(
        "cmg_xlsx", f"Generar {ARCHIVO_CMG}",
        "cmg_xlsx", None, "cmg_xlsx",
        # generar_cmg solo necesita las barras de 'Resumen BESS'.
        _REQ_CENTRALES_RESUMEN, ("cmg_csv",), (),
    ),
    Tarea(
        "fd_sscc", "Traer FD (SSCC_Desempeño_* del DCO)",
        "fd_sscc", None, "sscc",
        (), (), (),
    ),
    Tarea(
        "fma_cpf", "Generar fma_cpf", "fma", "cpf", "fma_cpf", (), (), (),
    ),
    Tarea(
        "fma_csf", "Generar fma_csf", "fma", "csf", "fma_csf", (), (), (),
    ),
    Tarea(
        "fma_ctf", "Generar fma_cft", "fma", "ctf", "fma_ctf", (), (), (),
    ),
    Tarea(
        "subastas_db", "Traer subastas (Access del periodo)",
        "subastas_db", None, "db_subastas",
        (), (), (),
    ),
    # --- Las 5 hojas de entrada de la planilla (en UNA escritura) ---
    Tarea(
        "consolidado:medidores", "hoja 'Medidores'",
        "consolidado", "medidores", "consolidado:medidores",
        _REQ_CENTRALES + ("soc",), ("medidas_sae",), (),
    ),
    Tarea(
        # No depende de la hoja 'Medidores' ya escrita:
        # generar_consolidado la reconstruye en memoria a partir de
        # Medidas_SAE (por eso depende de ese paso, no de la hoja).
        "consolidado:ofertas_sscc", "hoja 'Ofertas SSCC'",
        "consolidado", "ofertas_sscc", "consolidado:ofertas_sscc",
        _REQ_CENTRALES + ("soc", "ofertas"), ("medidas_sae",), (),
    ),
    Tarea(
        "consolidado:cmg", "hoja 'CMg'",
        "consolidado", "cmg", "consolidado:cmg",
        (), ("cmg_xlsx",), (),
    ),
    Tarea(
        "consolidado:fd", "hoja 'FD'",
        "consolidado", "fd", "consolidado:fd",
        (), ("fd_sscc",), (),
    ),
    Tarea(
        # El FMA es OPCIONAL (si falta, esa columna queda en 0, igual
        # que la formula original): no bloquea, pero si se genera en
        # esta corrida tiene que ser antes.
        "consolidado:subastas", "hoja 'Subastas'",
        "consolidado", "subastas", "consolidado:subastas",
        _REQ_CENTRALES, ("subastas_db", "fd_sscc"),
        ("fma_cpf", "fma_csf", "fma_ctf"),
    ),
    # --- Las hojas de calculo de la planilla (en UNA escritura) ------
    Tarea(
        "pagos:ecostos", "hoja 'Calculo E Costos'",
        "pagos", "ecostos", "pagos:ecostos",
        _REQ_CENTRALES,
        (
            "consolidado:medidores", "consolidado:ofertas_sscc",
            "consolidado:cmg", "consolidado:fd", "consolidado:subastas",
        ),
        (),
    ),
    Tarea(
        # RE545 no usa el FD.
        "pagos:re545", "hoja 'Calculo RE545'",
        "pagos", "re545", "pagos:re545",
        _REQ_CENTRALES,
        (
            "consolidado:medidores", "consolidado:ofertas_sscc",
            "consolidado:cmg", "consolidado:subastas",
        ),
        (),
    ),
    Tarea(
        # Llego despues que el grafo (venia de otra rama): resume por
        # central lo que calculan las dos hojas de arriba, asi que
        # depende de las dos y del Propietario de 'Resumen BESS'.
        "pagos:compensacion_central", f"hoja '{HOJA_COMPENSACION_CENTRAL}'",
        "pagos", "compensacion_central", "pagos:compensacion_central",
        _REQ_CENTRALES_RESUMEN, ("pagos:ecostos", "pagos:re545"), (),
    ),
    Tarea(
        "pagos:prorrata_retiros", "hoja 'PRORRATA_RETIROS'",
        "pagos", "prorrata_retiros", "pagos:prorrata_retiros",
        ("prorrata_archivo",), ("pagos:ecostos", "pagos:re545"), (),
    ),
    Tarea(
        "pagos:resumen", "hoja 'Resumen'",
        "pagos", "resumen", "pagos:resumen",
        _REQ_CENTRALES_RESUMEN + ("prorrata_archivo",),
        ("pagos:prorrata_retiros",), (),
    ),
)

TAREA_POR_ID = {tarea.id: tarea for tarea in TAREAS}


# Como se llama en la ventana cada requisito que puede bloquear, para
# poder decir "falta X" sin que la ventana tenga que saber de ids.
NOMBRE_REQUISITO = {
    "centrales": ARCHIVO_CENTRALES,
    f"centrales:{HOJA_RESUMEN_BESS}":
        f"{ARCHIVO_CENTRALES}, hoja '{HOJA_RESUMEN_BESS}'",
    f"centrales:{HOJA_DICCIONARIO}":
        f"{ARCHIVO_CENTRALES}, hoja '{HOJA_DICCIONARIO}'",
    "homologacion": "el Excel de homologacion",
    f"homologacion:{Homologacion.HOJA_HOMOL}":
        f"el Excel de homologacion, hoja '{Homologacion.HOJA_HOMOL}'",
    "soc": "el archivo de SoC del periodo",
    "ofertas": "el archivo *OfertasSSCC*",
    "prorrata_archivo": "el Excel de prorrata de retiros del periodo",
}


# Las entradas que pone la persona antes de poder hacer NADA de
# calculo (las que nombro el usuario: "centrales, homologacion y
# ofertas", mas el SoC del periodo). Si falta una de estas, el boton
# se bloquea: no tiene sentido ofrecer una corrida que no va a poder
# terminar.
ENTRADAS_INICIALES = (
    "centrales",
    f"centrales:{HOJA_RESUMEN_BESS}",
    f"centrales:{HOJA_DICCIONARIO}",
    "homologacion",
    f"homologacion:{Homologacion.HOJA_HOMOL}",
    "ofertas",
    "soc",
)

# Entradas que suelen llegar DESPUES (el Excel de prorrata de retiros
# del periodo). No bloquean el boton: las hojas que dependen de ellas
# quedan fuera del plan, a la vista y con el motivo, y se pueden
# tildar a mano cuando el archivo aparezca.
ENTRADAS_TARDIAS = ("prorrata_archivo",)


def _nombre_requisito(id_fila):
    return NOMBRE_REQUISITO.get(id_fila, id_fila)


def dependientes(id_tarea):
    """
    Todo lo que depende (directa o indirectamente) de una tarea, en el
    orden del grafo. Lo usa la ventana: si se fuerza a rehacer algo
    que ya estaba al dia, lo que sale de ahi queda viejo.

    Las dependencias OPCIONALES tambien cuentan aca: si se rehace el
    FMA, la hoja 'Subastas' quedo vieja aunque el FMA no la bloquee.
    """

    pendientes = [id_tarea]
    encontrados = []

    while pendientes:
        actual = pendientes.pop(0)
        for tarea in TAREAS:
            if actual in tarea.depende + tarea.depende_opcional:
                if tarea.id not in encontrados:
                    encontrados.append(tarea.id)
                    pendientes.append(tarea.id)

    return [t.id for t in TAREAS if t.id in encontrados]


def propagar_seleccion(seleccion):
    """
    La seleccion mas todo lo que depende de ella. La ventana la llama
    cuando se tilda una tarea: rehacer Medidas_SAE sin rehacer las
    hojas que salen de ahi deja el libro mezclado entre dos corridas.
    """

    completa = set(seleccion)

    for id_tarea in list(seleccion):
        completa.update(dependientes(id_tarea))

    return completa


# ============================================================
# EL PLAN
# ============================================================

def _seleccion_por_defecto(hechas, estado_fila):
    """
    Lo que el boton propone hacer: todo lo que falta y SE PUEDE hacer.

    Una tarea que no se puede hacer porque le falta una entrada de las
    que llegan tarde (ENTRADAS_TARDIAS: el Excel de prorrata de
    retiros) no entra: si entrara, el boton quedaria bloqueado por
    algo que el usuario todavia no puede resolver, en vez de hacer
    todo lo demas. Queda a la vista igual, como "sin tildar" y
    diciendo que le falta, y se puede tildar a mano.

    Lo mismo, en cadena, con lo que salga de una tarea asi.

    Las ENTRADAS_INICIALES (Centrales, homologacion, ofertas, SoC) NO
    se tratan asi: si falta una, el plan igual las incluye y el boton
    queda bloqueado diciendo que falta -- que es lo que pidio el
    usuario.
    """

    elegidas = {t.id for t in TAREAS if not hechas[t.id]}

    while True:

        descartar = set()

        for tarea in TAREAS:

            if tarea.id not in elegidas:
                continue

            sin_requisito = any(
                estado_fila.get(req) != "ok"
                for req in tarea.requisitos
                if req in ENTRADAS_TARDIAS
            )
            sin_dependencia = any(
                not hechas[dep] and dep not in elegidas
                for dep in tarea.depende
            )

            if sin_requisito or sin_dependencia:
                descartar.add(tarea.id)

        if not descartar:
            return elegidas

        elegidas -= descartar


def planificar(carpeta_base, aamm, seleccion=None, filas=None):
    """
    Que se va a hacer y que no, sin hacer nada.

    seleccion: ids de tarea a correr. None = "lo que falta" (todo lo
    que no este al dia), que es el arranque normal del boton.
    filas: el resultado de revisar_estructura() si la ventana ya lo
    tiene (evita revisar la carpeta dos veces).

    Devuelve un dict:
      tareas       lista de dicts, en orden de ejecucion, con:
                   id, etiqueta, grupo, hecha, seleccionada,
                   faltan_requisitos, faltan_dependencias, estado
                   ("al dia" / "se genera" / "se rehace" / "bloqueada"
                   / "sin tildar")
      bloqueos     lista de textos: que falta para poder ejecutar
      advertencias lista de textos: lo que queda viejo por no tildarlo
      puede_ejecutar  bool
      orden        grupos a correr, en orden de dependencia
    """

    if filas is None:
        _, filas = revisar_estructura(carpeta_base, aamm)

    estado_fila = {fila["id"]: fila["estado"] for fila in filas}

    bloqueos = []

    try:
        validar_aamm(aamm)
    except ErrorEntrada as error:
        bloqueos.append(str(error).split("\n")[0])

    for entrada in ENTRADAS_INICIALES:
        if estado_fila.get(entrada) != "ok":
            bloqueos.append(f"Falta {_nombre_requisito(entrada)}")

    hechas = {
        tarea.id: estado_fila.get(tarea.fila_estado) == "ok"
        for tarea in TAREAS
    }

    if seleccion is None:
        seleccion = _seleccion_por_defecto(hechas, estado_fila)
    else:
        seleccion = set(seleccion)

    tareas = []
    advertencias = []

    for tarea in TAREAS:

        elegida = tarea.id in seleccion

        faltan_requisitos = [
            _nombre_requisito(req) for req in tarea.requisitos
            if estado_fila.get(req) != "ok"
        ]

        faltan_dependencias = [
            TAREA_POR_ID[dep].etiqueta for dep in tarea.depende
            if not hechas[dep] and dep not in seleccion
        ]

        if elegida and (faltan_requisitos or faltan_dependencias):
            estado = "bloqueada"
        elif elegida and hechas[tarea.id]:
            estado = "se rehace"
        elif elegida:
            estado = "se genera"
        elif hechas[tarea.id]:
            estado = "al dia"
        else:
            estado = "sin tildar"

        if estado == "bloqueada":
            for falta in faltan_requisitos:
                bloqueos.append(f"{tarea.etiqueta}: falta {falta}")
            for falta in faltan_dependencias:
                bloqueos.append(
                    f"{tarea.etiqueta}: antes hay que hacer '{falta}'"
                )

        # Lo que YA estaba hecho y sale de algo que se rehace queda
        # viejo si no se tilda tambien.
        if not elegida and hechas[tarea.id]:
            de_donde = [
                dep for dep in tarea.depende + tarea.depende_opcional
                if dep in seleccion
            ]
            if de_donde:
                advertencias.append(
                    f"{tarea.etiqueta} NO se rehace y sale de "
                    f"{', '.join(TAREA_POR_ID[d].etiqueta for d in de_donde)}"
                    f": va a quedar de la corrida anterior."
                )

        tareas.append({
            "id": tarea.id,
            "etiqueta": tarea.etiqueta,
            "grupo": tarea.grupo,
            "hecha": hechas[tarea.id],
            "seleccionada": elegida,
            "faltan_requisitos": faltan_requisitos,
            "faltan_dependencias": faltan_dependencias,
            "estado": estado,
        })

    elegidas = [t for t in tareas if t["seleccionada"]]

    if not elegidas:
        bloqueos.append("No hay nada tildado para ejecutar.")

    return {
        "tareas": tareas,
        "bloqueos": bloqueos,
        "advertencias": advertencias,
        "puede_ejecutar": not bloqueos,
        "orden": _orden_de_grupos(
            {t["id"] for t in elegidas}
        ),
    }


def _grupos_seleccionados(seleccion):
    """{id de grupo: [secciones tildadas de ese grupo]} en orden."""

    grupos = {}

    for tarea in TAREAS:
        if tarea.id in seleccion:
            grupos.setdefault(tarea.grupo, [])
            if tarea.seccion is not None:
                grupos[tarea.grupo].append(tarea.seccion)

    return grupos


def _dependencias_entre_grupos(seleccion):
    """
    {id de grupo: set de grupos que tienen que terminar antes}, solo
    entre los grupos que se van a correr.
    """

    grupos = set(_grupos_seleccionados(seleccion))
    dependencias = {grupo: set() for grupo in grupos}

    for tarea in TAREAS:

        if tarea.id not in seleccion:
            continue

        for dep in tarea.depende + tarea.depende_opcional:
            if dep in seleccion:
                grupo_dep = TAREA_POR_ID[dep].grupo
                if grupo_dep != tarea.grupo:
                    dependencias[tarea.grupo].add(grupo_dep)

    return dependencias


def _orden_de_grupos(seleccion):
    """
    Los grupos a correr en un orden que respeta las dependencias
    (topologico, estable segun el orden de GRUPOS). Es el orden que
    muestra la ventana; la ejecucion real ademas los corre en paralelo
    cuando no dependen entre si.
    """

    dependencias = _dependencias_entre_grupos(seleccion)
    pendientes = [g.id for g in GRUPOS if g.id in dependencias]
    orden = []

    while pendientes:
        listos = [
            g for g in pendientes
            if not (dependencias[g] - set(orden))
        ]
        if not listos:
            # No puede pasar con el grafo de arriba (no tiene ciclos),
            # pero si alguien agrega uno, mejor un orden raro que un
            # cuelgue.
            listos = pendientes[:1]
        orden.extend(listos)
        pendientes = [g for g in pendientes if g not in listos]

    return orden


# ============================================================
# LA CORRIDA
# ============================================================

def ejecutar_plan(
    carpeta_base, aamm, seleccion=None, registrar=print, progreso=None,
    hilos=3, plan=None,
):
    """
    Corre lo tildado respetando el grafo, en paralelo donde se puede.

    seleccion: ids de tarea (None = lo que falte, igual que
    planificar()).
    hilos: cuantos grupos pueden correr a la vez (1 = todo en fila,
    util para depurar).

    Devuelve {id de grupo: "ok" | "error: ..." | "omitido: ..."}.
    Levanta ErrorEntrada al final si algo fallo -- con el detalle de
    que si se hizo, para no dejar la ventana diciendo "listo" cuando
    la mitad no corrio.
    """

    if plan is None:
        plan = planificar(carpeta_base, aamm, seleccion)

    if plan["bloqueos"]:
        raise ErrorEntrada(
            "No se puede ejecutar todavia:\n  - "
            + "\n  - ".join(plan["bloqueos"])
        )

    seleccion = {t["id"] for t in plan["tareas"] if t["seleccionada"]}

    grupos = _grupos_seleccionados(seleccion)
    dependencias = _dependencias_entre_grupos(seleccion)
    orden = plan["orden"]

    for aviso in plan["advertencias"]:
        registrar(f"[AVISO] {aviso}")

    registrar(
        f"Plan: {len(seleccion)} tarea(s) en {len(orden)} paso(s) "
        f"-> {' -> '.join(orden)}"
    )

    candado_registro = threading.Lock()
    candados_recurso = {}
    for grupo in GRUPOS:
        candados_recurso.setdefault(grupo.recurso, threading.Lock())

    resultados = {}
    terminados = set()
    total = len(orden)
    hechos = 0

    def registrar_grupo(id_grupo, mensaje):
        with candado_registro:
            registrar(f"[{id_grupo}] {mensaje}")

    def correr(id_grupo):
        """
        Un grupo entero (una llamada a la funcion de siempre). El
        candado de recurso es lo que garantiza que dos grupos que
        escriben el mismo archivo no se pisen aunque el grafo los deje
        sueltos.
        """

        grupo = GRUPO_POR_ID[id_grupo]

        with candados_recurso[grupo.recurso]:
            registrar_grupo(id_grupo, f"Empezando: {grupo.etiqueta}")
            grupo.ejecutar(
                carpeta_base,
                aamm,
                grupos[id_grupo],
                lambda m: registrar_grupo(id_grupo, m),
                # El progreso fino de cada paso no se mezcla en una
                # sola barra: la de la ventana avanza por paso hecho.
                None,
            )
            registrar_grupo(id_grupo, "Listo.")

    with ThreadPoolExecutor(max_workers=max(1, hilos)) as pool:

        pendientes = list(orden)
        en_vuelo = {}

        while pendientes or en_vuelo:

            for id_grupo in list(pendientes):

                fallidas = sorted(
                    dep for dep in dependencias[id_grupo]
                    if resultados.get(dep, "ok") != "ok"
                )

                if fallidas:
                    resultados[id_grupo] = (
                        f"omitido: no corrio {', '.join(fallidas)}"
                    )
                    registrar(
                        f"[{id_grupo}] NO se corre: no termino bien "
                        f"{', '.join(fallidas)}."
                    )
                    pendientes.remove(id_grupo)
                    terminados.add(id_grupo)
                    hechos += 1
                    continue

                if dependencias[id_grupo] - terminados:
                    continue

                pendientes.remove(id_grupo)
                en_vuelo[pool.submit(correr, id_grupo)] = id_grupo

            if not en_vuelo:
                break

            listos, _ = wait(list(en_vuelo), return_when=FIRST_COMPLETED)

            for futuro in listos:

                id_grupo = en_vuelo.pop(futuro)

                try:
                    futuro.result()
                    resultados[id_grupo] = "ok"
                except Exception as error:  # noqa: BLE001 - se informa
                    resultados[id_grupo] = f"error: {error}"
                    registrar(f"[{id_grupo}] ERROR: {error}")

                terminados.add(id_grupo)
                hechos += 1

                if progreso and total:
                    progreso(int(100 * hechos / total))

    fallidos = {
        id_grupo: detalle for id_grupo, detalle in resultados.items()
        if detalle != "ok"
    }

    registrar("")
    registrar("Resumen de la corrida:")
    for id_grupo in orden:
        registrar(
            f"  {GRUPO_POR_ID[id_grupo].etiqueta}: "
            f"{resultados.get(id_grupo, 'sin correr')}"
        )

    if fallidos:
        raise ErrorEntrada(
            "La corrida termino con problemas:\n  - "
            + "\n  - ".join(
                f"{GRUPO_POR_ID[g].etiqueta}: {d}"
                for g, d in fallidos.items()
            )
        )

    return resultados
