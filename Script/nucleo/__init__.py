# -*- coding: utf-8 -*-
"""
Nucleo de calculo: una etapa por modulo.

Sin dependencias de interfaz: todo lo de aca se puede correr y testear
sin abrir la ventana.

Este archivo es solo la fachada. `nucleo.lo_que_sea` sigue funcionando
igual que cuando todo vivia en un unico nucleo.py, pero cada cosa esta
ahora en el modulo de su etapa:

    parametros.py            Nombres de archivo, carpeta y hoja del caso; mapeo A:I.
    utiles.py                Normalizacion y el error de entrada: lo que usa todo.
    alertas.py               El registro de alertas y el estado de la corrida.
    avisos.py                Diagnostico comun de los cruces que terminan en cero.
    conciliacion.py          Medidores = E Costos + RE545 (control TRA).
    manifiesto.py            Que archivo exacto alimento esta corrida.
    rutas.py                 Rutas del caso y busqueda de los archivos de entrada.
    estructura.py            El arbol de carpetas y archivos que dibuja la ventana.
    lectura.py               Medidas_SAE.xlsx y el maestro Centrales.xlsx.
    soc.py                   El SoC por bloques, desde el archivo del SCADA.
    ofertas_sscc.py          Ofertas SSCC y las columnas R, S, T, V de Medidores.
    hojas_entrada.py         Hojas CMg, FD y Subastas del consolidado.
    subastas_accdb.py        Subastas desde su origen: los Access OfertasSSCCAdj*.
    fma.py                   Subastas!Q (FMA) y Subastas!P (FD).
    diccionarios.py          Diccionarios del maestro y del CMg: los usan las dos hojas.
    ecostos.py               Calculo E Costos: traspaso base y orquestacion.
    columnas_compartidas.py  L, M y N/O: iguales en las dos hojas de calculo.
    ecostos_columnas.py      E Costos, etapa 2: R, S, T, U, W, X, Y, AB:AF.
    ecostos_prorratas.py     E Costos, etapa 3: AG:AL, AM:AR (FD), AS:AV.
    ecostos_ciclo.py         E Costos, etapa 4: Ciclo, umbrales, AW, AX, AZ.
    re545.py                 Calculo RE545: traspaso base, S, U, V y orquestacion.
    re545_reservas.py        RE545: los tres bloques de reservas (AC:AT) y AU.
    re545_resumen.py         RE545: la tabla por central+ventana (AW:BG) y BV.
    re545_componentes.py     RE545: Componente 1 y Componente 2 (BI:CE).
    prorrata_retiros.py      PRORRATA_RETIROS: quien paga, cuarto a cuarto.
    compensacion.py          COMPENSACION_CENTRAL y Resumen: quien recibe.
    medidores.py             La hoja Medidores: columnas calculadas y armado.
    escritura.py             Escritura de los dos libros de salida.
    proceso.py               Los dos procesos completos, de punta a punta.
    origenes.py              De donde viene lo que se trae: el "Origen:" del diagrama.
    traer.py                 Botones 'Traer'/'Generar': lo que el programa baja o arma.
    medidas_sae.py           Medidas_SAE.xlsx: los cuatro pasos de un viaje.

El orden de los imports de abajo es el de esa lista: de las piezas
compartidas hacia las etapas que las usan.

"""

from . import externos
from .externos import (
    Claves_Balance,
    Descarga_PRMTE,
    ErrorMedidas,
    Generacion_Real,
    Homologacion,
    desempeno_fd,
    extrae_cmg,
    fma_subastas,
    indicadores_dco,
    indices_fma,
    ofertas_adj,
)

from .parametros import (
    ARCHIVO_CENTRALES, ARCHIVO_CMG, ARCHIVO_CONTROL, ARCHIVO_MEDIDAS_SAE,
    ARCHIVO_SALIDA, CARPETA_AUXILIARES, CARPETA_CMG,
    CARPETA_DB_SUBASTAS, CARPETA_FD_FMA, CARPETA_FD_FMA_ANTIGUA,
    CARPETA_MEDIDAS, CARPETA_OFERTAS, CARPETA_PRORRATA_RETIROS,
    CARPETA_SUBASTAS,
    CARPETA_TRABAJO_MEDIDAS, COLUMNAS_AI,
    COLUMNAS_AUXILIARES_MEDIDORES, COLUMNAS_MEDIDORES_SALIDA,
    COLUMNAS_VACIAS,
    EXTENSIONES_EXCEL, HOJA_ALERTAS, HOJA_CALCULO_ECOSTOS,
    HOJA_CALCULO_RE545, HOJA_CMG, HOJA_EJECUCION, HOJA_FD, HOJA_LOG,
    HOJA_MEDIDORES, HOJA_PRORRATA_RETIROS, HOJA_RESUMEN, HOJA_SUBASTAS,
    ORDEN_HOJAS_CONTROL, ORDEN_HOJAS_SALIDA,
    HOJA_CMG_ORIGEN, HOJA_CPF_HORARIO, HOJA_CSF_HORARIO,
    HOJA_DICCIONARIO, HOJA_MEDIDAS_SAE, HOJA_RESUMEN_BESS,
    COLUMNAS_OFERTAS_EN_MEDIDORES, INICIO_VENTANA, LETRA_A_CAMPO,
    PATRON_AAMM,
    PATRON_NOMBRE_OFERTAS, PATRON_NOMBRE_SSCC_DESEMPENO,
    UMBRAL_SOC,
)
from .utiles import (
    ErrorEntrada, _columna_clave_vba, _entero_a_texto, _es_numero,
    _normaliza_valor_vba, _texto_seguro, _tiene_valor, _valor_clave,
    normalizar,
)
from .alertas import (
    ALTA, APROBADA, APROBADA_CON_ADVERTENCIAS, COLUMNAS_ALERTA, CRITICA,
    FALLIDA, INFO, MEDIA, NO_APROBADA, SEVERIDADES, Alerta, Registro,
    anotar, anotar_muchas,
)
from .avisos import _avisar_claves_sin_mapeo
from .conciliacion import (
    TOLERANCIA_ABSOLUTA, TOLERANCIA_RELATIVA, conciliar_energia,
)
from .manifiesto import (
    COLUMNAS_MANIFIESTO, construir_manifiesto, sha256_de,
)
from .rutas import (
    _buscar_archivo_excel_mas_reciente, _es_archivo_de_soc,
    buscar_archivo_ofertas, buscar_archivo_sscc_desempeno,
    SUBCARPETAS_CASO, buscar_soc, crear_estructura_caso,
    CARPETA_DEL_PERIODO, CARPETA_DE_OTRO_PERIODO, CARPETA_SIN_PERIODO,
    carpeta_corresponde_al_periodo, nombre_caso_sugerido,
    periodo_desde_aamm, periodos_en_el_nombre,
    resolver_rutas, validar_aamm,
)
from .origenes import ORIGENES, origen, ruta_origen
from .estructura import (
    SECCIONES_CONSOLIDADO, SECCIONES_PAGOS, SECCIONES_SALIDA, _fila,
    _filas_de_hojas,
    hojas_con_datos, hojas_de, revisar_estructura,
)
from .prorrata_retiros import (
    COLUMNAS_ORIGEN, HOJA_ORIGEN,
    buscar_archivo_prorrata, construir_prorrata_retiros,
    leer_prorrata_retiros,
)
from .compensacion import (
    construir_compensacion_central, construir_compensacion_total,
    construir_resumen,
)
from .lectura import (
    ROL_BALANCE_BESS, ROL_FD, ROL_FMA_CPF, ROL_OFERTAS, ROL_SUBASTAS,
    TITULOS_DICCIONARIO, encabezado_diccionario, filas_diccionario,
    mapa_diccionario,
    _bloques_columnas_diccionario, _leer_hoja_con_encabezado,
    _leer_resumen_bess, construir_homologacion, leer_centrales,
    leer_medidas_sae,
)
from .soc import (
    _extraer_nombre_desde_ruta_scada, detectar_bloques,
    detectar_fila_nombres, extraer_soc,
)
from .ofertas_sscc import (
    _contiene_bess_o_sae, _es_respuesta_si, _homologar_fge,
    _limpiar_nombre_mostrar, _mapas_homologacion_fge,
    _normalizar_nombre_clave, _normalizar_periodo,
    _servicio_termina_en_rs, _valor_oferta_binario, calcular_r,
    calcular_s, calcular_t, cargar_resumen_en_medidores,
    construir_resumen_ofertas_sscc, construir_resumen_ventana_oferta,
    HOJA_OFERTAS_SSCC, TITULO_OFERTAS_POR_DIA, TITULO_RESUMEN_VENTANA,
    leer_ofertas_sscc_consolidado,
)
from .hojas_entrada import (
    NOMBRES_FD_CPF, NOMBRES_FD_CSF, NOMBRES_SUBASTAS,
    _construir_bloque_fd_cpf, _construir_bloque_fd_csf,
    _contiene_bess_o_sae_sin_bat, _dia_hora_mes_fd,
    _filtrar_bess_sae_posicional, _ordenar_subastas_por_hora_mes,
    construir_fd, leer_cmg, leer_cmg_consolidado, leer_fd_consolidado,
)
from .subastas_accdb import (
    COLUMNA_ENERGIA_SSCC_ACCDB, _control_desde_servicio,
    _sub_baj_desde_servicio, calcular_hora_mes_subastas,
    construir_mapa_propietario, construir_subastas_desde_accdb,
)
from .fma import (
    CTF_MAS_BUSCA_EN_LAS_DOS_TABLAS, TITULO_BLOQUE_FD,
    TITULO_BLOQUE_FMA_CPF, VECTOR_PARTICIPACION_CSF_SIN_DATO,
    _clave_numerica_o_texto, _dic_desde_tabla, _numero_o_cero,
    calcular_fd_subastas, calcular_fma_subastas, cargar_tablas_fma,
    construir_dic_bloque_diccionario,
)
from .diccionarios import (
    _buscar_cmg, _mapa_resumen_bess_por_nombre, _normaliza_cuarto,
    construir_dic_cmg, construir_dic_resumen_capacidad,
    construir_dic_resumen_eficiencia, construir_dic_resumen_factor,
    construir_mapa_barra,
)
from .ecostos import (
    COLUMNAS_AUXILIARES_E_COSTOS, COLUMNAS_SALIDA_E_COSTOS,
    GRUPOS_CALCULO_E_COSTOS, NOMBRES_CALCULO_E_COSTOS,
    completar_calculo_e_costos_grupos, construir_calculo_e_costos,
    renombrar_calculo_e_costos,
)
from .columnas_compartidas import (
    _construir_set_subastas_tipo, calcular_l, calcular_m, calcular_n_o,
)
from .ecostos_columnas import (
    _calcular_asignacion_energia, calcular_ae_af, calcular_r_ecostos,
    calcular_s_t_u, calcular_w_x, calcular_y_ab_ac_ad,
)
from .ecostos_prorratas import (
    _calcular_bloque, _calcular_costo_ponderado, calcular_as_at,
    calcular_au_av, calcular_fd_prorrateado, calcular_prorratas,
    construir_dic_fd_bloque, construir_dic_mapeo_diccionario,
    unidades_bloque_fd,
    construir_dic_prorrata, construir_prorrata_sscc,
)
from .ecostos_ciclo import (
    _clave_central_ciclo, calcular_aw_ax, calcular_az,
    calcular_subastas_ciclo, construir_dic_umbrales_subastas,
)
from .re545 import (
    COLUMNAS_AUXILIARES_RE545, COLUMNAS_SALIDA_RE545,
    GRUPOS_CALCULO_RE545, NOMBRES_CALCULO_RE545, calcular_s_re545,
    calcular_u_v_re545, completar_calculo_re545,
    completar_checks_resumen_re545, construir_calculo_re545,
    renombrar_calculo_re545,
)
from .re545_reservas import (
    TIPOS_RESERVA_RE545, _BLOQUES_RESERVA_RE545,
    calcular_reservas_re545, construir_dic_reservas_subastas,
)
from .re545_resumen import (
    NOMBRES_RESUMEN_RE545, _clave_grupo_re545, _mapa_resumen_por_grupo,
    calcular_bv_re545, construir_resumen_ventanas_re545,
)
from .re545_componentes import (
    calcular_bi_bj_re545, calcular_bk_bl_bm_bs_re545,
    calcular_componentes_re545,
)
from .medidores import (
    calcular_clave_auxiliar, calcular_indicador_soc, calcular_ventana,
    completar_ofertas_en_medidores, construir_medidores,
    construir_ofertas_sscc, reponer_auxiliares_medidores,
)
from .formato import formatear_hoja, formatear_libro
from .escritura import (
    _COLUMNA_Q_INDICE, _HOJAS_CONSOLIDADO,
    _HOJAS_PAGOS, _copiar_hoja_existente, _escribir_encabezados_grupo,
    _escribir_tabla_con_titulo, escribir_control, escribir_pagos_bess,
    escribir_salida,
)
from .proceso import generar_consolidado, generar_pagos_bess
from .orquestador import (
    GRUPOS, GRUPO_POR_ID, TAREAS, TAREA_POR_ID, dependientes, ejecutar_plan,
    planificar, propagar_seleccion,
)
from .traer import (
    barras_desde_resumen_bess, generar_cmg, generar_fma, traer_csv_cmg,
    traer_fd, traer_subastas,
)
from .medidas_sae import (
    _resumir_diagnostico_medidas, generar_medidas_sae,
)

# La API publica del paquete: lo que `nucleo.<algo>` ofrece. Estan
# tambien los helpers con guion bajo, porque las pruebas los usan
# por nombre.
__all__ = [
    "ALTA", "APROBADA", "APROBADA_CON_ADVERTENCIAS", "COLUMNAS_ALERTA",
    "CRITICA", "FALLIDA", "INFO", "MEDIA", "NO_APROBADA", "SEVERIDADES",
    "Alerta", "Registro", "anotar", "anotar_muchas",
    "TOLERANCIA_ABSOLUTA", "TOLERANCIA_RELATIVA", "conciliar_energia",
    "COLUMNAS_MANIFIESTO", "construir_manifiesto", "sha256_de",
    "externos", "Claves_Balance", "Descarga_PRMTE", "ErrorMedidas",
    "Generacion_Real", "Homologacion", "desempeno_fd", "extrae_cmg",
    "fma_subastas", "indicadores_dco", "indices_fma", "ofertas_adj",
    "ARCHIVO_CENTRALES", "ARCHIVO_CMG", "ARCHIVO_MEDIDAS_SAE",
    "ARCHIVO_SALIDA", "ARCHIVO_CONTROL", "CARPETA_AUXILIARES",
    "CARPETA_CMG", "CARPETA_DB_SUBASTAS", "CARPETA_FD_FMA",
    "CARPETA_FD_FMA_ANTIGUA", "CARPETA_MEDIDAS", "CARPETA_OFERTAS",
    "CARPETA_PRORRATA_RETIROS",
    "CARPETA_SUBASTAS", "CARPETA_TRABAJO_MEDIDAS", "COLUMNAS_AI",
    "COLUMNAS_VACIAS", "COLUMNAS_AUXILIARES_MEDIDORES",
    "COLUMNAS_MEDIDORES_SALIDA", "COLUMNAS_AUXILIARES_E_COSTOS",
    "COLUMNAS_SALIDA_E_COSTOS", "COLUMNAS_AUXILIARES_RE545",
    "COLUMNAS_SALIDA_RE545", "reponer_auxiliares_medidores",
    "formatear_hoja", "formatear_libro", "leer_cmg_consolidado",
    "GRUPOS", "GRUPO_POR_ID", "TAREAS", "TAREA_POR_ID", "dependientes",
    "ejecutar_plan", "planificar", "propagar_seleccion",
    "EXTENSIONES_EXCEL", "HOJA_CALCULO_ECOSTOS", "HOJA_ALERTAS",
    "HOJA_EJECUCION", "HOJA_LOG", "HOJA_MEDIDORES", "HOJA_CMG",
    "HOJA_FD", "HOJA_SUBASTAS", "ORDEN_HOJAS_SALIDA",
    "ORDEN_HOJAS_CONTROL", "SECCIONES_SALIDA", "escribir_control",
    "HOJA_CALCULO_RE545", "HOJA_PRORRATA_RETIROS", "HOJA_RESUMEN",
    "HOJA_CMG_ORIGEN", "HOJA_CPF_HORARIO",
    "HOJA_CSF_HORARIO", "HOJA_DICCIONARIO", "HOJA_MEDIDAS_SAE",
    "HOJA_RESUMEN_BESS", "INICIO_VENTANA",
    "LETRA_A_CAMPO", "COLUMNAS_OFERTAS_EN_MEDIDORES",
    "PATRON_AAMM", "PATRON_NOMBRE_OFERTAS",
    "PATRON_NOMBRE_SSCC_DESEMPENO",
    "UMBRAL_SOC", "ErrorEntrada", "_columna_clave_vba",
    "_entero_a_texto", "_es_numero", "_normaliza_valor_vba",
    "_texto_seguro", "_tiene_valor", "_valor_clave", "normalizar",
    "_avisar_claves_sin_mapeo", "_buscar_archivo_excel_mas_reciente",
    "_es_archivo_de_soc", "buscar_archivo_ofertas",
    "buscar_archivo_sscc_desempeno",
    "buscar_soc", "periodo_desde_aamm", "resolver_rutas",
    "SUBCARPETAS_CASO", "crear_estructura_caso", "nombre_caso_sugerido",
    "CARPETA_DEL_PERIODO", "CARPETA_DE_OTRO_PERIODO", "CARPETA_SIN_PERIODO",
    "carpeta_corresponde_al_periodo", "periodos_en_el_nombre",
    "validar_aamm", "SECCIONES_CONSOLIDADO", "SECCIONES_PAGOS", "_fila",
    "COLUMNAS_ORIGEN", "HOJA_ORIGEN",
    "buscar_archivo_prorrata", "leer_prorrata_retiros",
    "construir_prorrata_retiros", "construir_compensacion_total",
    "construir_compensacion_central", "construir_resumen",
    "_filas_de_hojas", "hojas_con_datos", "hojas_de",
    "revisar_estructura", "ROL_BALANCE_BESS", "ROL_FD", "ROL_FMA_CPF",
    "ROL_OFERTAS", "ROL_SUBASTAS", "TITULOS_DICCIONARIO",
    "encabezado_diccionario", "filas_diccionario", "mapa_diccionario",
    "_bloques_columnas_diccionario",
    "_leer_hoja_con_encabezado", "_leer_resumen_bess",
    "construir_homologacion", "leer_centrales", "leer_medidas_sae",
    "_extraer_nombre_desde_ruta_scada", "detectar_bloques",
    "detectar_fila_nombres", "extraer_soc", "_contiene_bess_o_sae",
    "_es_respuesta_si", "_homologar_fge", "_limpiar_nombre_mostrar",
    "_mapas_homologacion_fge", "_normalizar_nombre_clave",
    "_normalizar_periodo", "_servicio_termina_en_rs",
    "_valor_oferta_binario", "calcular_r", "calcular_s", "calcular_t",
    "cargar_resumen_en_medidores", "construir_resumen_ofertas_sscc",
    "construir_resumen_ventana_oferta", "NOMBRES_FD_CPF",
    "NOMBRES_FD_CSF", "NOMBRES_SUBASTAS", "_construir_bloque_fd_cpf",
    "_construir_bloque_fd_csf", "_contiene_bess_o_sae_sin_bat",
    "_dia_hora_mes_fd", "_filtrar_bess_sae_posicional",
    "_ordenar_subastas_por_hora_mes", "construir_fd",
    "leer_cmg", "leer_fd_consolidado", "COLUMNA_ENERGIA_SSCC_ACCDB",
    "_control_desde_servicio", "_sub_baj_desde_servicio",
    "calcular_hora_mes_subastas", "construir_mapa_propietario",
    "construir_subastas_desde_accdb", "CTF_MAS_BUSCA_EN_LAS_DOS_TABLAS",
    "TITULO_BLOQUE_FD", "TITULO_BLOQUE_FMA_CPF",
    "VECTOR_PARTICIPACION_CSF_SIN_DATO", "_clave_numerica_o_texto",
    "_dic_desde_tabla", "_numero_o_cero", "calcular_fd_subastas",
    "calcular_fma_subastas", "cargar_tablas_fma",
    "construir_dic_bloque_diccionario", "_buscar_cmg",
    "_mapa_resumen_bess_por_nombre", "_normaliza_cuarto",
    "construir_dic_cmg", "construir_dic_resumen_capacidad",
    "construir_dic_resumen_eficiencia", "construir_dic_resumen_factor",
    "construir_mapa_barra", "GRUPOS_CALCULO_E_COSTOS",
    "NOMBRES_CALCULO_E_COSTOS", "completar_calculo_e_costos_grupos",
    "renombrar_calculo_e_costos",
    "construir_calculo_e_costos", "_construir_set_subastas_tipo",
    "calcular_l", "calcular_m", "calcular_n_o",
    "_calcular_asignacion_energia", "calcular_ae_af",
    "calcular_r_ecostos", "calcular_s_t_u", "calcular_w_x",
    "calcular_y_ab_ac_ad", "_calcular_bloque",
    "_calcular_costo_ponderado", "calcular_as_at", "calcular_au_av",
    "calcular_fd_prorrateado", "calcular_prorratas",
    "construir_dic_fd_bloque", "construir_dic_mapeo_diccionario",
    "unidades_bloque_fd",
    "construir_dic_prorrata", "construir_prorrata_sscc",
    "_clave_central_ciclo", "calcular_aw_ax", "calcular_az",
    "calcular_subastas_ciclo", "construir_dic_umbrales_subastas",
    "GRUPOS_CALCULO_RE545", "NOMBRES_CALCULO_RE545", "calcular_s_re545",
    "calcular_u_v_re545", "completar_calculo_re545",
    "completar_checks_resumen_re545", "construir_calculo_re545",
    "renombrar_calculo_re545", "TIPOS_RESERVA_RE545",
    "_BLOQUES_RESERVA_RE545", "calcular_reservas_re545",
    "construir_dic_reservas_subastas", "NOMBRES_RESUMEN_RE545",
    "_clave_grupo_re545", "_mapa_resumen_por_grupo",
    "calcular_bv_re545", "construir_resumen_ventanas_re545",
    "calcular_bi_bj_re545", "calcular_bk_bl_bm_bs_re545",
    "calcular_componentes_re545", "calcular_clave_auxiliar",
    "calcular_indicador_soc", "calcular_ventana", "construir_medidores",
    "construir_ofertas_sscc", "completar_ofertas_en_medidores",
    "HOJA_OFERTAS_SSCC", "TITULO_OFERTAS_POR_DIA",
    "TITULO_RESUMEN_VENTANA", "leer_ofertas_sscc_consolidado",
    "_COLUMNA_Q_INDICE", "_HOJAS_CONSOLIDADO",
    "_HOJAS_PAGOS", "_copiar_hoja_existente",
    "_escribir_encabezados_grupo", "_escribir_tabla_con_titulo",
    "escribir_pagos_bess", "escribir_salida", "generar_consolidado",
    "generar_pagos_bess", "barras_desde_resumen_bess", "generar_cmg",
    "generar_fma", "traer_csv_cmg", "traer_fd", "traer_subastas",
    "_resumir_diagnostico_medidas", "generar_medidas_sae",
]
