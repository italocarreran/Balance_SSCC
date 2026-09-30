# -*- coding: utf-8 -*-
"""
Los dos procesos completos, de punta a punta.
"""

import pandas as pd
from pathlib import Path

from .externos import desempeno_fd, ofertas_adj

from .alertas import CRITICA, Alerta, Registro, anotar
from .conciliacion import conciliar_energia
from .diccionarios import (
    construir_dic_cmg, construir_dic_resumen_capacidad,
    construir_dic_resumen_eficiencia, construir_dic_resumen_factor,
    construir_mapa_barra,
)
from .ecostos import (
    completar_calculo_e_costos_grupos, construir_calculo_e_costos,
    renombrar_calculo_e_costos,
)
from .manifiesto import construir_manifiesto
from .ecostos_prorratas import construir_dic_mapeo_diccionario
from .escritura import escribir_pagos_bess, escribir_salida
from .estructura import SECCIONES_CONSOLIDADO, SECCIONES_PAGOS
from .fma import (
    TITULO_BLOQUE_FD, TITULO_BLOQUE_FMA_CPF, cargar_tablas_fma,
    construir_dic_bloque_diccionario,
)
from .hojas_entrada import (
    construir_fd, leer_cmg, leer_cmg_consolidado, leer_fd_consolidado,
)
from .lectura import (
    construir_homologacion, leer_centrales, leer_medidas_sae,
)
from .medidores import (
    completar_ofertas_en_medidores, construir_medidores,
    construir_ofertas_sscc, reponer_auxiliares_medidores,
)
from .ofertas_sscc import leer_ofertas_sscc_consolidado
from .parametros import (
    ARCHIVO_CENTRALES, ARCHIVO_CMG, ARCHIVO_MEDIDAS_SAE,
    CARPETA_DB_SUBASTAS, CARPETA_FD_FMA, HOJA_DICCIONARIO,
    HOJA_CALCULO_ECOSTOS, HOJA_CALCULO_RE545,
)
from .prorrata_retiros import (
    buscar_archivo_prorrata, construir_prorrata_retiros,
    leer_prorrata_retiros,
)
from .compensacion import (
    construir_compensacion_central, construir_resumen,
)
from .re545 import (
    completar_calculo_re545, completar_checks_resumen_re545,
    construir_calculo_re545, renombrar_calculo_re545,
)
from .re545_componentes import calcular_componentes_re545
from .re545_resumen import construir_resumen_ventanas_re545
from .rutas import (
    buscar_archivo_ofertas, buscar_archivo_sscc_desempeno,
    buscar_soc, periodo_desde_aamm, resolver_rutas, validar_aamm,
)
from .soc import extraer_soc
from .subastas_accdb import (
    construir_mapa_propietario, construir_subastas_desde_accdb,
)
from .utiles import ErrorEntrada


def generar_consolidado(
    carpeta_base, aamm, secciones_activas, registrar=print, progreso=None
):
    """
    Genera/actualiza las hojas de ENTRADA de la planilla de salida
    (ARCHIVO_SALIDA), recalculando solo las
    hojas de las secciones pedidas (ids de SECCIONES_CONSOLIDADO) y
    preservando el resto tal cual estaba en el archivo existente (ver
    escribir_salida). Si el archivo no existe, se crea. La usan los
    botones "Actualizar" de las filas-hoja del diagrama (y el
    "Actualizar" de la fila del archivo, que manda todas).

    secciones_activas: iterable de ids de SECCIONES_CONSOLIDADO
    ("medidores", "ofertas_sscc", "cmg", "fd", "subastas") a
    recalcular esta vez.

    La dependencia entre "medidores" y "ofertas_sscc" se dio vuelta a
    pedido del usuario: ahora es "Ofertas SSCC" la que necesita
    Medidores (de ahi salen las centrales y las ventanas), y no al
    reves. Consecuencia practica: actualizar SOLO "Medidores" no abre
    el archivo *OfertasSSCC* ni lo exige; actualizar "Ofertas SSCC" si
    recalcula Medidores en memoria, pero no lo reescribe si no se
    pidio.
    """

    def avanzar(valor):
        if progreso:
            progreso(valor)

    secciones_activas = set(secciones_activas)
    ids_validos = {seccion[0] for seccion in SECCIONES_CONSOLIDADO}
    desconocidas = secciones_activas - ids_validos

    if desconocidas:
        raise ErrorEntrada(
            f"Seccion(es) desconocida(s): {sorted(desconocidas)}"
        )

    if not secciones_activas:
        raise ErrorEntrada(
            "No se tildo ninguna entrada para generar/actualizar."
        )

    hojas_regenerar = set()
    for id_seccion, _, _, hojas in SECCIONES_CONSOLIDADO:
        if id_seccion in secciones_activas:
            hojas_regenerar.update(hojas)

    rutas = resolver_rutas(carpeta_base)

    # Centrales.xlsx lo necesitan dos secciones ("medidores" y
    # "subastas"): se abre UNA vez por corrida, no una por seccion
    # (pedido del usuario: "que no se abran planillas innecesarias").
    centrales_leidas = {}

    def leer_centrales_una_vez():
        if "datos" not in centrales_leidas:
            registrar(f"Leyendo {ARCHIVO_CENTRALES}...")
            centrales_leidas["datos"] = leer_centrales(rutas["centrales"])
        return centrales_leidas["datos"]

    df_medidores = None
    avisos, incidencias = [], []
    df_wxy = df_resumen_ventana = None
    df_cmg = df_fd_csf = df_fd_cpf = df_subastas = None

    avanzar(5)

    quiere_ofertas = "ofertas_sscc" in secciones_activas

    if "medidores" in secciones_activas or quiere_ofertas:

        aamm_val = validar_aamm(aamm)

        if not rutas["medidas_sae"].is_file():
            raise ErrorEntrada(
                f"No se encontro {rutas['medidas_sae']}"
            )

        if not rutas["centrales"].is_file():
            raise ErrorEntrada(
                f"No se encontro {rutas['centrales']}"
            )

        # El archivo de ofertas SOLO hace falta para la hoja "Ofertas
        # SSCC": Medidores ya no depende de el (pedido del usuario de
        # independizar una hoja de la otra).
        archivo_ofertas = None

        if quiere_ofertas:
            archivo_ofertas = buscar_archivo_ofertas(rutas["ofertas_dir"])
            if not archivo_ofertas:
                raise ErrorEntrada(
                    f"No se encontro ningun archivo *OfertasSSCC* en "
                    f"{rutas['ofertas_dir']}"
                )

        archivo_soc = buscar_soc(rutas["medidas_dir"], aamm_val)
        anio, mes = periodo_desde_aamm(aamm_val)

        registrar(f"Periodo indicado: {anio}-{mes:02d} ({aamm_val})")

        _, diccionario = leer_centrales_una_vez()
        mapa = construir_homologacion(diccionario)
        registrar(f"  homologaciones cargadas: {len(mapa):,}")
        avanzar(20)

        registrar(f"Leyendo {archivo_soc.name}...")
        df_soc, incidencias = extraer_soc(archivo_soc, mapa)
        registrar(
            f"  bloques leidos: "
            f"{df_soc['central'].nunique()}   "
            f"registros: {len(df_soc):,}"
        )

        for incidencia in incidencias:
            registrar(f"  [SOC] {incidencia}")

        avanzar(35)

        registrar(f"Leyendo {ARCHIVO_MEDIDAS_SAE}...")
        df_sae = leer_medidas_sae(rutas["medidas_sae"])
        registrar(f"  filas: {len(df_sae):,}")
        avanzar(50)

        registrar("Construyendo Medidores...")
        df_medidores, avisos = construir_medidores(
            df_sae, df_soc, mes, registrar=registrar
        )

        if quiere_ofertas:
            registrar("Construyendo Ofertas SSCC...")
            df_wxy, df_resumen_ventana, avisos_ofertas = (
                construir_ofertas_sscc(
                    df_medidores, archivo_ofertas, diccionario, anio, mes,
                    registrar=registrar,
                )
            )
            avisos.extend(avisos_ofertas)

        for aviso in avisos:
            registrar(f"  [AVISO] {aviso}")

    avanzar(60)

    if "cmg" in secciones_activas:
        if not rutas["cmg"].is_file():
            raise ErrorEntrada(f"No se encontro {rutas['cmg']}")
        registrar(f"Leyendo {ARCHIVO_CMG}...")
        df_cmg = leer_cmg(rutas["cmg"], registrar=registrar)

    avanzar(72)

    if "fd" in secciones_activas:
        archivo_sscc = buscar_archivo_sscc_desempeno(
            rutas["sscc_desempeno_dir"]
        )
        if not archivo_sscc:
            raise ErrorEntrada(
                f"No se encontro ningun archivo SSCC_Desempeño_* en "
                f"{rutas['sscc_desempeno_dir']}"
            )
        registrar(f"Leyendo {archivo_sscc.name}...")
        df_fd_csf, df_fd_cpf = construir_fd(
            archivo_sscc, registrar=registrar
        )

    avanzar(84)

    if "subastas" in secciones_activas:

        aamm_val = validar_aamm(aamm)
        hay_accdb = bool(
            ofertas_adj.accdb_presentes(rutas["subastas_dir"], aamm_val)
        )

        if hay_accdb:
            # Camino normal desde ahora: el origen real de las
            # subastas son los Access, no la planilla 3.
            mapa_propietarios = {}
            dic_cpf = {}
            diccionario_centrales = None

            if rutas["centrales"].is_file():
                resumen_bess, diccionario_centrales = leer_centrales_una_vez()
                mapa_propietarios = construir_mapa_propietario(resumen_bess)
                dic_cpf = construir_dic_bloque_diccionario(
                    diccionario_centrales, TITULO_BLOQUE_FMA_CPF
                )
            else:
                registrar(
                    f"  [AVISO] no se encontro {rutas['centrales']}: la "
                    f"columna Propietario de Subastas queda vacia."
                )

            registrar(
                f"Leyendo las salidas de FMA de {CARPETA_FD_FMA}/..."
            )
            tablas_fma = cargar_tablas_fma(
                rutas["sscc_desempeno_dir"], aamm_val, registrar=registrar
            )

            # El FD (Subastas!P) y el Vector de Participacion CSF que
            # necesita el FMA salen los dos del archivo SSCC_Desempeño_*.
            tablas_fd = None
            dic_unidad_fd = {}

            archivo_sscc_fd = buscar_archivo_sscc_desempeno(
                rutas["sscc_desempeno_dir"]
            )

            if archivo_sscc_fd:
                registrar(f"Leyendo el FD de {archivo_sscc_fd.name}...")
                try:
                    tablas_fd = desempeno_fd.construir_tablas_fd(
                        archivo_sscc_fd, registrar=registrar
                    )
                except desempeno_fd.ErrorDesempeno as error:
                    raise ErrorEntrada(str(error)) from error

                if rutas["centrales"].is_file():
                    dic_unidad_fd = construir_dic_bloque_diccionario(
                        diccionario_centrales, TITULO_BLOQUE_FD
                    ) or construir_dic_mapeo_diccionario(diccionario_centrales)
            else:
                registrar(
                    f"  [AVISO] no hay ningun SSCC_Desempeño_* en "
                    f"{CARPETA_FD_FMA}/ (se baja con el boton 'Traer "
                    f"FD'): las columnas FD y el Vector de Participacion "
                    f"CSF quedan sin dato."
                )

            if not dic_cpf:
                registrar(
                    f"  [AVISO] {ARCHIVO_CENTRALES} no tiene la columna "
                    f"'FMA_CPF' en la hoja {HOJA_DICCIONARIO} "
                    f"(Configuración -> central como la nombra fma_cpf): "
                    f"el FMA de las filas CPF se busca con el nombre tal "
                    f"cual y va a quedar en 0. Es el formato de "
                    f"Diccionario de una sola tabla, con los encabezados "
                    f"Balance_BESS | FD | Subastas | Ofertas | FMA_CPF."
                )

            registrar(
                f"Leyendo las subastas de {CARPETA_DB_SUBASTAS}/ "
                f"(periodo {aamm_val})..."
            )
            df_subastas = construir_subastas_desde_accdb(
                rutas["subastas_dir"],
                aamm_val,
                mapa_propietarios=mapa_propietarios,
                tablas_fma=tablas_fma,
                dic_cpf=dic_cpf,
                tablas_fd=tablas_fd,
                dic_unidad_fd=dic_unidad_fd,
                registrar=registrar,
            )

        else:
            # La planilla 3 (3_REMUNERACIÓN_SUBASTAS_E_ID_*) dejo de
            # ser un respaldo: el usuario confirmo que ya no se usa, y
            # tampoco era el origen (se arma pegando lo que sale de
            # estos mismos Access). Sin Access no hay subastas.
            raise ErrorEntrada(
                f"No hay de donde sacar las subastas del periodo "
                f"{aamm_val}: {rutas['db_subastas_dir']} no tiene "
                f"ningun Access del periodo. Traelos con el boton "
                f"'Traer' de esa fila."
            )

    avanzar(92)

    registrar(f"Escribiendo {rutas['salida'].name}...")
    escribir_salida(
        df_medidores,
        rutas["salida"],
        avisos,
        incidencias,
        df_wxy,
        df_resumen_ventana,
        df_cmg,
        df_fd_csf,
        df_fd_cpf,
        df_subastas,
        ruta_existente=rutas["salida"],
        hojas_regenerar=hojas_regenerar,
        registrar=registrar,
        ruta_control=rutas["control"],
    )

    avanzar(100)
    registrar(f"Listo: {rutas['salida']}")

    return rutas["salida"]


def _leer_entradas_del_consolidado(ruta_salida, quiere_fd, registrar):
    """
    Las hojas de entrada de la planilla que necesitan las dos
    hojas de calculo, leidas con UNA sola apertura del archivo.

    Devuelve (df_medidores, df_wxy, df_resumen_ventana, df_cmg,
    df_subastas, df_fd_csf, df_fd_cpf); los dos ultimos son None si
    no se pidio "ecostos" (RE545 no usa FD).

    Existe para dos cosas que pidio el usuario:
      - que no se abran planillas innecesarias: un
        pd.read_excel(ruta, sheet_name=...) por hoja volvia a parsear
        el libro entero cinco veces, y es el archivo mas grande del
        caso;
      - que lo que ya esta en el consolidado salga de ahi: ninguna de
        estas cinco hojas se vuelve a calcular ni se lee de su origen.
    """

    ruta_salida = Path(ruta_salida)

    with pd.ExcelFile(ruta_salida) as libro:

        registrar(f"Leyendo hoja 'Medidores' de {ruta_salida.name}...")

        try:
            df_medidores = pd.read_excel(libro, sheet_name="Medidores")
        except ValueError as error:
            raise ErrorEntrada(
                f"{ruta_salida.name} no tiene la hoja 'Medidores' "
                f"todavia. Genera esa hoja primero "
                f"(tildando 'Medidores + Ofertas SSCC')."
            ) from error

        if df_medidores.empty:
            raise ErrorEntrada(
                f"La hoja 'Medidores' de {ruta_salida.name} esta "
                f"vacia. Genera esa hoja primero "
                f"(tildando 'Medidores + Ofertas SSCC')."
            )

        registrar(f"  filas: {len(df_medidores):,}")

        # La hoja ya no trae Copia_Ventana ("Ciclo de Carga del mes"):
        # es copia de Ventana y se repone aca, sin recalcular nada.
        # Ver COLUMNAS_AUXILIARES_MEDIDORES.
        df_medidores = reponer_auxiliares_medidores(df_medidores)

        # Las tres columnas que salen de Ofertas SSCC (R, S, T) ya no
        # viven en la hoja Medidores: se reconstruyen a partir de las
        # dos tablas de la hoja "Ofertas SSCC". T es la que decide,
        # fila por fila, si la energia va a "Calculo E Costos" o a
        # "Calculo RE545", asi que sin esto no hay calculo posible.
        registrar(f"Leyendo hoja 'Ofertas SSCC' de {ruta_salida.name}...")
        df_wxy, df_resumen_ventana = leer_ofertas_sscc_consolidado(
            ruta_salida, registrar=registrar, libro=libro
        )

        registrar(f"Leyendo hoja 'CMg' de {ruta_salida.name}...")
        df_cmg = leer_cmg_consolidado(
            ruta_salida, registrar=registrar, libro=libro
        )

        registrar(f"Leyendo hoja 'Subastas' de {ruta_salida.name}...")

        try:
            df_subastas = pd.read_excel(libro, sheet_name="Subastas")
        except ValueError as error:
            raise ErrorEntrada(
                f"{ruta_salida.name} no tiene la hoja 'Subastas' "
                f"todavia. Genera esa hoja primero "
                f"(tildando 'Subastas')."
            ) from error

        if df_subastas.empty:
            raise ErrorEntrada(
                f"La hoja 'Subastas' de {ruta_salida.name} esta "
                f"vacia. Genera esa hoja primero "
                f"(tildando 'Subastas')."
            )

        df_fd_csf = df_fd_cpf = None

        if quiere_fd:
            # AM:AR salen de la hoja 'FD' del consolidado, no de
            # releer el SSCC_Desempeño_*: el consolidado es la foto de
            # las entradas con la que se calcula todo.
            registrar(f"Leyendo hoja 'FD' de {ruta_salida.name}...")
            df_fd_csf, df_fd_cpf = leer_fd_consolidado(
                ruta_salida, registrar=registrar, libro=libro
            )

    return (
        df_medidores, df_wxy, df_resumen_ventana, df_cmg, df_subastas,
        df_fd_csf, df_fd_cpf,
    )


def generar_pagos_bess(
    carpeta_base, secciones_activas, registrar=print, progreso=None, aamm=None
):
    """
    Genera/actualiza Pagos_BESS.xlsx, recalculando solo las hojas de
    las secciones pedidas (ids de SECCIONES_PAGOS: "ecostos",
    "re545") y preservando el resto tal cual estaba en el archivo
    existente (ver escribir_pagos_bess/hojas_regenerar) -- mismo
    criterio que generar_consolidado()/SECCIONES_CONSOLIDADO. Si el
    archivo no existe, se crea. La usan los botones "Actualizar" de
    las filas-hoja del diagrama.

    No recalcula Medidores ni Subastas: los lee tal cual estan en
    la misma planilla, que debe generarse primero con su
    propio boton "Actualizar". Centrales.xlsx y cmg.xlsx si se leen/
    recalculan frescos. El archivo SSCC_Desempeño_* solo se exige si
    se pide "ecostos" -- "re545" no usa FD.
    """

    def avanzar(valor):
        if progreso:
            progreso(valor)

    # Todo lo que el calculo escriba al log pasa por este Registro: las
    # alertas quedan guardadas con su id y severidad, no solo impresas.
    registro = Registro(salida=registrar)
    registrar = registro

    secciones_activas = set(secciones_activas)
    ids_validos = {seccion[0] for seccion in SECCIONES_PAGOS}
    desconocidas = secciones_activas - ids_validos

    if desconocidas:
        raise ErrorEntrada(
            f"Seccion(es) desconocida(s): {sorted(desconocidas)}"
        )

    if not secciones_activas:
        raise ErrorEntrada(
            "No se tildo ninguna seccion para generar/actualizar."
        )

    quiere_ecostos = "ecostos" in secciones_activas
    quiere_re545 = "re545" in secciones_activas
    quiere_compensacion = "compensacion_central" in secciones_activas
    quiere_prorrata = "prorrata_retiros" in secciones_activas
    quiere_resumen = "resumen" in secciones_activas

    hojas_regenerar = set()
    for id_seccion, _, _, hojas in SECCIONES_PAGOS:
        if id_seccion in secciones_activas:
            hojas_regenerar.update(hojas)

    rutas = resolver_rutas(carpeta_base)

    # ----------------------------------------------------------
    # LECTURA DE ENTRADAS: SOLO LO QUE PIDEN LAS SECCIONES TILDADAS
    #
    # Pedido del usuario ("que no se abran planillas innecesarias"):
    # antes esta funcion abria Medidores, Ofertas SSCC, Subastas,
    # Centrales.xlsx y cmg.xlsx SIEMPRE, incluso para recalcular solo
    # PRORRATA_RETIROS o el Resumen -- dos hojas que no miran ninguna
    # de esas entradas, sino las dos hojas de calculo ya escritas.
    #
    #   Medidores / Ofertas SSCC / Subastas / CMg  -> solo si hay que
    #     recalcular alguna de las dos hojas de calculo;
    #   Centrales.xlsx -> ademas para el Resumen (el Propietario de
    #     cada central);
    #   la hoja 'FD' del consolidado -> solo para E Costos (RE545 no
    #     usa FD), como ya era.
    # ----------------------------------------------------------

    quiere_calculo = quiere_ecostos or quiere_re545

    df_medidores = None
    df_wxy = df_resumen_ventana = None
    df_cmg = df_subastas = None
    df_fd_csf = df_fd_cpf = None
    resumen = diccionario = None
    mapa_barra = dic_cmg = dic_factor = umbral_soc_minimo = None
    dic_eficiencia = dic_capacidad = None
    periodo_medidores = validar_aamm(aamm) if aamm else ""

    if quiere_calculo:

        # El consolidado solo hace falta para las dos hojas de
        # calculo: PRORRATA_RETIROS y el Resumen salen de las hojas ya
        # escritas en Pagos_BESS.xlsx, asi que ni se abre.
        if not rutas["salida"].is_file():
            raise ErrorEntrada(
                f"No se encontro {rutas['salida']}. Primero hay que "
                f"generar las hojas de entrada de la planilla "
                f"(Medidores, Ofertas SSCC, CMg, FD y Subastas) con sus "
                f"botones 'Actualizar'."
            )

        (
            df_medidores, df_wxy, df_resumen_ventana, df_cmg, df_subastas,
            df_fd_csf, df_fd_cpf,
        ) = _leer_entradas_del_consolidado(
            rutas["salida"], quiere_ecostos, registrar
        )

        # La hoja Medidores de la que salen los dos calculos tiene que
        # ser de un solo mes: si trae dos, el libro quedo mezclado
        # entre corridas de periodos distintos y todo lo que siga paga
        # mal.
        meses = sorted(
            int(m) for m in pd.to_numeric(
                df_medidores["Mes"], errors="coerce"
            ).dropna().unique()
        )
        periodo_medidores = "-".join(f"{m:02d}" for m in meses)

        if len(meses) > 1:
            anotar(registrar, Alerta(
                "PER-001", CRITICA, "Traspaso",
                f"La hoja 'Medidores' de {rutas['salida'].name} tiene "
                f"{len(meses)} meses distintos ({periodo_medidores}): el "
                f"libro quedo mezclado entre corridas de periodos distintos.",
                valor_encontrado=periodo_medidores,
                valor_esperado="un unico mes",
                accion="corrida NO APROBADA",
                archivo=rutas["salida"].name, hoja="Medidores",
                origen_control="CATALOGO PER-001",
            ))

    avanzar(10)

    if quiere_calculo or quiere_resumen:

        if not rutas["centrales"].is_file():
            raise ErrorEntrada(f"No se encontro {rutas['centrales']}")

        registrar(f"Leyendo {ARCHIVO_CENTRALES}...")
        resumen, diccionario = leer_centrales(rutas["centrales"])

    if quiere_calculo:

        df_medidores, _ = completar_ofertas_en_medidores(
            df_medidores, df_wxy, df_resumen_ventana, diccionario,
            registrar=registrar,
        )

        mapa_barra = construir_mapa_barra(resumen)
        dic_factor, umbral_soc_minimo = construir_dic_resumen_factor(resumen)

        # Eficiencia y Capacidad solo las usa RE545 (V y U/BC/BN).
        if quiere_re545:
            dic_eficiencia = construir_dic_resumen_eficiencia(resumen)
            dic_capacidad = construir_dic_resumen_capacidad(resumen)

        dic_cmg = construir_dic_cmg(df_cmg)

    avanzar(40)

    df_ecostos = None
    df_re545 = None
    df_resumen_re545 = None
    df_prorrata_retiros = None
    df_compensacion_cuarto = None
    df_pagos_suministrador = None
    df_compensacion_ecostos = None
    df_compensacion_re545 = None
    df_compensacion_empresa = None
    df_resumen = None

    # Las dos hojas, con los nombres INTERNOS de columna. Las que se
    # escriben (df_ecostos/df_re545) ya pasaron por su renombrar_*(), y
    # ahi "Energia_Positiva" se llama "Descarga kWh": la conciliacion
    # tiene que mirar estas, no aquellas.
    df_ecostos_base = None
    df_re545_base = None

    if quiere_ecostos:

        registrar(
            "Construyendo Calculo E Costos (etapa base: H + CMg + "
            "traspaso de Medidores)..."
        )
        df_ecostos = construir_calculo_e_costos(
            df_medidores, mapa_barra, dic_cmg, registrar=registrar
        )
        avanzar(50)

        # df_fd_csf/df_fd_cpf ya vienen leidos de la hoja 'FD' del
        # consolidado (_leer_entradas_del_consolidado), no de releer
        # el SSCC_Desempeño_*: el consolidado es la foto de las
        # entradas con la que se calcula todo. Si la hoja FD esta
        # vieja, se regenera con su propio boton 'Actualizar'.
        avanzar(60)

        registrar(
            "Completando L, M, N, O, R, S, T, U, W, X, Y, AB, AC, AD, AE, "
            "AF, AG, AH, AI, AJ, AK, AL, AM, AN, AO, AP, AQ, AR, AS, AT, "
            "AU, AV, AW, AX, AZ..."
        )
        df_ecostos_base = completar_calculo_e_costos_grupos(
            df_ecostos, df_subastas, dic_factor, umbral_soc_minimo,
            diccionario, df_fd_csf, df_fd_cpf,
            registrar=registrar,
        )
        df_ecostos = renombrar_calculo_e_costos(df_ecostos_base)

    avanzar(70)

    if quiere_re545:

        registrar(
            "Construyendo Calculo RE545 (etapa base: traspaso de "
            "Medidores + L, M, N, O, S, U, V)..."
        )
        df_re545 = construir_calculo_re545(
            df_medidores, mapa_barra, dic_cmg, registrar=registrar
        )
        df_re545_base = completar_calculo_re545(
            df_re545, df_subastas, umbral_soc_minimo, dic_capacidad,
            dic_eficiencia, registrar=registrar,
        )

        # AY ("Oferta Completa") del resumen por central+ventana sale
        # de la misma tabla que alimenta Medidores!T: la tabla
        # "Resumen ventana oferta" de la hoja 'Ofertas SSCC' del
        # consolidado, leida arriba. Antes se recalculaba aca a partir
        # de la hoja Medidores; ahora se usa directamente la generada,
        # que es el mismo dato y una cuenta menos.
        df_resumen_re545 = construir_resumen_ventanas_re545(
            df_re545_base, df_resumen_ventana, dic_capacidad,
            registrar=registrar,
        )

        registrar("  Calculo RE545: Componente 1 y Componente 2 (BI:CE)...")
        for interno, serie in calcular_componentes_re545(
            df_re545_base, df_resumen_re545, dic_factor,
            registrar=registrar,
        ).items():
            df_re545_base[interno] = serie

        df_resumen_re545 = completar_checks_resumen_re545(
            df_resumen_re545, df_re545_base
        )

        df_re545 = renombrar_calculo_re545(df_re545_base)

    avanzar(90)

    archivo_prorrata = None
    if quiere_prorrata or quiere_resumen or quiere_compensacion:
        # Estas hojas tienen botones independientes. Si las hojas de calculo
        # no se recalcularon en esta misma accion, se consumen las versiones
        # ya guardadas en la planilla.
        def calculo_existente(nombre):
            try:
                return pd.read_excel(
                    rutas["salida"], sheet_name=nombre, header=1
                )
            except (FileNotFoundError, ValueError) as error:
                raise ErrorEntrada(
                    f"Primero calcula la hoja '{nombre}' de "
                    f"{rutas['salida'].name}."
                ) from error

        df_ecostos_asignacion = (
            df_ecostos if df_ecostos is not None
            else calculo_existente(HOJA_CALCULO_ECOSTOS)
        )
        df_re545_asignacion = (
            df_re545 if df_re545 is not None
            else calculo_existente(HOJA_CALCULO_RE545)
        )
        # Quien RECIBE. El Resumen tambien lo necesita (es su columna
        # RECIBE), asi que se calcula una sola vez para las dos hojas.
        if quiere_compensacion or quiere_resumen:
            registrar("Compensacion por central y ciclo/ventana...")
            (
                df_compensacion_ecostos,
                df_compensacion_re545,
                df_compensacion_empresa,
            ) = construir_compensacion_central(
                df_ecostos_asignacion, df_re545_asignacion, resumen,
                registrar=registrar,
            )
            registrar(
                f"  E Costos: {len(df_compensacion_ecostos):,} grupos; "
                f"RE545: {len(df_compensacion_re545):,} grupos; "
                f"{len(df_compensacion_empresa):,} empresas reciben."
            )

        # Quien PAGA. Solo aca hace falta el Excel de la prorrata.
        if quiere_prorrata or quiere_resumen:
            archivo_prorrata = buscar_archivo_prorrata(
                rutas["prorrata_retiros_dir"], aamm
            )
            if archivo_prorrata is None:
                raise ErrorEntrada(
                    f"No se encontro Prorrata_Retiros_AAMM_pre/def.xlsx en "
                    f"{rutas['prorrata_retiros_dir']}."
                )
            registrar(f"Leyendo {archivo_prorrata.name}/Prorrata 15min...")
            fuente_prorrata = leer_prorrata_retiros(
                archivo_prorrata, registrar=registrar
            )
            (
                df_compensacion_cuarto,
                df_prorrata_retiros,
                df_pagos_suministrador,
            ) = construir_prorrata_retiros(
                fuente_prorrata, df_ecostos_asignacion, df_re545_asignacion,
                registrar=registrar,
            )
            registrar(
                f"  {len(df_compensacion_cuarto):,} cuartos de hora con monto; "
                f"{len(df_prorrata_retiros):,} asignaciones; "
                f"{len(df_pagos_suministrador):,} suministradores."
            )

        if quiere_resumen:
            registrar("Resumen: RECIBE, PAGA y NETO por empresa...")
            df_resumen = construir_resumen(
                df_compensacion_empresa, df_pagos_suministrador,
                registrar=registrar,
            )

    registrar(f"Escribiendo {rutas['salida'].name}...")
    # TRA: la energia de Medidores tiene que repartirse entera entre
    # las dos hojas. Va antes de escribir: si no cuadra, la corrida
    # queda NO APROBADA y eso se escribe en el libro.
    registrar("Conciliando energia Medidores -> E Costos + RE545...")
    # Y va dentro de un try: si la conciliacion revienta por algo
    # inesperado, lo unico que se pierde tiene que ser la conciliacion
    # -no todo lo calculado-. Antes un error aca dejaba la corrida sin
    # archivo y sin nada que revisar (el KeyError 'Energia_Positiva'
    # que reporto el usuario).
    try:
        conciliacion = conciliar_energia(
            df_medidores, df_ecostos_base, df_re545_base,
            registrar=registrar,
        )
    except Exception as error:  # noqa: BLE001 - se informa, no se tapa
        anotar(registrar, Alerta(
            "TRA-010", CRITICA, "Traspaso",
            f"No se pudo conciliar la energia: {type(error).__name__}: "
            f"{error}.",
            accion="se escribe igual el archivo, sin conciliacion",
            origen_control="CATALOGO TRA-010",
        ))
        conciliacion = {
            "conciliacion_completa": False,
            "error": f"{type(error).__name__}: {error}",
        }

    # cmg.xlsx ya no es una entrada de esta etapa: el CMg sale de la
    # hoja 'CMg' del consolidado, que si esta en el manifiesto.
    manifiesto = construir_manifiesto([
        ("Planilla de entradas", rutas["salida"]),
        ("Centrales", rutas["centrales"]),
        ("Prorrata retiros", archivo_prorrata),
    ])

    # Todo por nombre: son once tablas y el orden posicional ya se
    # presto una vez a confusion.
    escribir_pagos_bess(
        rutas["salida"],
        df_ecostos=df_ecostos,
        df_re545=df_re545,
        df_resumen_re545=df_resumen_re545,
        df_compensacion_ecostos=df_compensacion_ecostos,
        df_compensacion_re545=df_compensacion_re545,
        df_compensacion_empresa=df_compensacion_empresa,
        df_prorrata_retiros=df_prorrata_retiros,
        df_compensacion_cuarto=df_compensacion_cuarto,
        df_pagos_suministrador=df_pagos_suministrador,
        df_resumen=df_resumen,
        ruta_existente=rutas["salida"],
        hojas_regenerar=hojas_regenerar,
        registrar=registrar,
        registro=registro,
        manifiesto=manifiesto,
        conciliacion=conciliacion,
        periodo=periodo_medidores,
        ruta_control=rutas["control"],
    )

    avanzar(100)
    registrar(f"Listo: {rutas['salida']}")

    for linea in registro.resumen(
        periodo=periodo_medidores,
        hojas_regeneradas=sorted(hojas_regenerar),
    ).split("\n"):
        registrar(linea)

    return rutas["salida"]
