# -*- coding: utf-8 -*-
"""
Reportes en palabras sobre los resultados de una planilla 11_PAGOS_BESS.

Pedido del usuario (2026-11-10), dos cosas que comparten este modulo:

  1. "en el control de corrida [...] un reporte sobre la corrida, tipo si
     una central tiene cero ingreso que diga algo tipo, segun FD no tuvo
     participacion, o FD=0" -> generar_reporte_corrida(): hoja "Reporte"
     de Control_corrida.xlsx, con su boton "Generar reporte".
  2. "un archivo que compare resultados entre el archivo y el mes o
     version anterior [...] tal central aumento su ingreso por RE545 dado
     que las ofertas respecto al otro aumentaron tanto" ->
     comparar_planillas() (la ventana es Herramientas/Comparar_Meses.py).

"Ordenados, sin mucho ruido": una fila por central, una sola frase por
central, y solo cuando hay algo que decir.

Todo sale de lo que ya esta en la planilla (no recalcula nada):
  - lo que recibe: "Monto a compensar" de Calculo E Costos y Calculo
    RE545, por central (lo mismo que suma COMPENSACION_CENTRAL);
  - la energia: Descarga/Carga y Adj SSCC de Calculo E Costos, y la
    reserva con FD y FMA de Calculo RE545;
  - lo adjudicado: hoja Subastas (traducida al nombre Balance_BESS por el
    Diccionario), por servicio;
  - el FD: hoja FD, el Factor de Desempeño promedio por central;
  - las ofertas: hoja Ofertas SSCC, dias con oferta completa;
  - la empresa: Propietario de la hoja Centrales.
Una hoja que no esta deja sus columnas vacias (NaN), y las frases no la
usan.
"""

import math
from pathlib import Path

import pandas as pd

from .lector import abrir_excel, leer_excel
from .lectura import (
    fila_de_encabezado, leer_centrales_planilla, leer_hoja_con_fila_vacia,
    sin_margen_izquierdo,
)
from .parametros import (
    HOJA_CALCULO_ECOSTOS, HOJA_CALCULO_RE545, HOJA_FD, HOJA_MEDIDORES,
    HOJA_OFERTAS_SSCC, HOJA_SUBASTAS,
)
from .unidades import sin_unidades
from .utiles import ErrorEntrada, _texto_seguro, normalizar

# Las columnas del resumen por central (las de plata primero).
C_CENTRAL = "Central"
C_EMPRESA = "Empresa"
C_ECOSTOS = "Recibe E Costos [CLP]"
C_RE545 = "Recibe RE545 [CLP]"
C_TOTAL = "Recibe total [CLP]"
C_DESCARGA = "Descarga [kWh]"
C_CARGA = "Carga [kWh]"
C_CMG = "CMg promedio [CLP/kWh]"
C_HORAS_ADJ = "Horas con Adj SSCC"
C_ADJ_CPF = "Adjudicado CPF [MW]"
C_ADJ_CSF = "Adjudicado CSF [MW]"
C_RESERVA = "Reserva con FD y FMA [kWh]"
C_FD_CPF = "FD CPF promedio"
C_FD_CSF = "FD CSF promedio"
C_OFERTA_DIAS = "Días con oferta completa"
C_DIAS = "Días en el mes"
C_HORAS_SUB = "Horas adjudicadas"
# Como se arma el monto (por ciclo en E Costos, por ventana en RE545):
# max(0, Total C1 - Total C2). C1 es la energia ajustada por FD y
# prorrata, valorizada con la curva monotona; C2 la energia medida al
# CMg. Ahi entra el FD: en la energia ajustada de C1.
# Horas en que la Respuesta del FD dice "No participó" (con FD = 1
# igual: el factor no lo muestra, pero el monto queda en cero).
C_NP_CSF = "Horas FD CSF 'No participó'"
C_NP_CPF = "Horas FD CPF 'No participó'"
C_FDH_CSF = "Horas en FD CSF"
# Lo mismo, pero SOLO en las horas en que la central tenia ese servicio
# adjudicado: "No participó" fuera de esas horas es lo normal.
C_NPA_CPF = "Horas adjudicadas CPF con 'No participó'"
C_NPA_CSF = "Horas adjudicadas CSF con 'No participó'"
C_HA_CPF = "Horas adjudicadas CPF en el FD"
C_HA_CSF = "Horas adjudicadas CSF en el FD"
C_FDH_CPF = "Horas en FD CPF"
C_EC_C1 = "E Costos Total C1 [CLP]"
C_EC_C2 = "E Costos Total C2 [CLP]"
C_EC_AJUSTADA = "Descarga ajustada FD-prorrata [kWh]"
C_EC_CICLOS = "Ciclos con monto"
C_EC_CICLOS_MES = "Ciclos del mes"
C_RE_C1 = "RE545 Total C1 [CLP]"
C_RE_C2 = "RE545 Total C2 [CLP]"
C_RE_VENTANAS = "Ventanas con monto"
C_RE_VENTANAS_MES = "Ventanas del mes"

COLUMNAS_RESUMEN = (
    C_CENTRAL, C_EMPRESA, C_ECOSTOS, C_RE545, C_TOTAL, C_DIAS, C_DESCARGA,
    C_CARGA, C_CMG, C_HORAS_ADJ, C_HORAS_SUB, C_ADJ_CPF, C_ADJ_CSF, C_RESERVA,
    C_FD_CPF, C_FD_CSF, C_NP_CPF, C_NP_CSF, C_FDH_CPF, C_FDH_CSF,
    C_NPA_CPF, C_NPA_CSF, C_HA_CPF, C_HA_CSF, C_OFERTA_DIAS,
    C_EC_C1, C_EC_C2, C_EC_AJUSTADA, C_EC_CICLOS, C_EC_CICLOS_MES,
    C_RE_C1, C_RE_C2, C_RE_VENTANAS, C_RE_VENTANAS_MES,
)

# Un cambio menor a esto (en % y en plata) no se comenta.
CAMBIO_MINIMO_PCT = 10.0
CAMBIO_MINIMO_CLP = 1000.0


# ------------------------------------------------------------
# Lectura
# ------------------------------------------------------------

def _hoja_de_calculo(libro, hoja):
    """Una hoja de calculo con sus nombres de columna (sin unidades)."""

    arriba = leer_excel(libro, sheet_name=hoja, header=None, nrows=4)
    df = leer_excel(libro, sheet_name=hoja, header=fila_de_encabezado(arriba) or 1)
    df = sin_unidades(sin_margen_izquierdo(df), hoja)
    return df.loc[:, ~df.columns.duplicated()]


def _col(df, nombre):
    """La columna 'nombre' (la primera, si viene repetida) o None."""

    if df is None or nombre not in df:
        return None
    columna = df[nombre]
    if isinstance(columna, pd.DataFrame):
        columna = columna.iloc[:, 0]
    return columna


def _num(serie):
    return pd.to_numeric(serie, errors="coerce")


def _por_central(df, columna_central, columna_valor, como="sum"):
    """{central normalizada: valor agregado} (vacio si falta algo)."""

    central = _col(df, columna_central)
    valor = _col(df, columna_valor)
    if central is None or valor is None:
        return {}
    tabla = pd.DataFrame({"c": central.map(normalizar), "v": _num(valor)})
    tabla = tabla[tabla["c"] != ""]
    if como == "sum":
        # Sin ningun valor (una columna con formula sin valor guardado)
        # es "sin dato", no 0.
        return tabla.groupby("c")["v"].sum(min_count=1).to_dict()
    return tabla.groupby("c")["v"].agg(como).to_dict()


def completar_formulas(df, plantillas, nombres):
    """
    Las columnas con formula de una hoja de calculo que vienen VACIAS (un
    libro escrito antes de que se guardara el valor calculado, y que no se
    abrio ni guardo en Excel) se calculan aca con la MISMA plantilla que
    escribe la hoja (FORMULAS_E_COSTOS / FORMULAS_RE545), desde las
    columnas con valores. Solo aritmetica (+ - * / y parentesis) y el
    IF(COUNT(a,b)<2,"",x) que deja vacio si falta un dato (eso ya lo hace
    el NaN); una plantilla con otra funcion queda como esta.
    """

    import re

    marcador = re.compile(r"\{(\w+)\}")
    si_falta = re.compile(r'^IF\(COUNT\([^)]*\)<2,"",(.*)\)$')
    solo_cuentas = re.compile(r"^[\s\d.+\-*/()v\[\]'\w]*$")
    en_curso = set()

    def tiene_valor(nombre):
        return nombre in df and _num(_col(df, nombre)).notna().any()

    def calcular(clave):
        nombre = nombres.get(clave, clave)
        if tiene_valor(nombre) or clave not in plantillas or clave in en_curso:
            return
        en_curso.add(clave)
        expresion = plantillas[clave].lstrip("=")
        hallado = si_falta.match(expresion)
        if hallado:
            expresion = hallado.group(1)
        if re.search(r"[A-Z]{2,}\(", marcador.sub("x", expresion)):
            return  # otra funcion de Excel: no se calcula
        valores = {}
        for otra in marcador.findall(expresion):
            calcular(otra)
            columna = nombres.get(otra, otra)
            if columna not in df:
                return
            valores[otra] = _num(_col(df, columna)).to_numpy(dtype=float)
        codigo = marcador.sub(lambda m: f"v['{m.group(1)}']", expresion)
        if not solo_cuentas.match(codigo):
            return
        df[nombre] = eval(codigo, {"__builtins__": {}}, {"v": valores})  # noqa: S307

    for clave in plantillas:
        calcular(clave)
    return df


def _leer(libro, hoja, funcion, registrar):
    if hoja not in libro.sheet_names:
        return None
    try:
        return funcion()
    except Exception as error:  # noqa: BLE001 - el reporte sigue sin ella
        registrar(f"  [AVISO] no se pudo leer la hoja '{hoja}': {error}")
        return None


def resumen_por_central(ruta_planilla, registrar=print):
    """
    DataFrame con una fila por central (COLUMNAS_RESUMEN) de la planilla
    11_PAGOS_BESS en 'ruta_planilla'.
    """

    from .hojas_entrada import (
        leer_fd_consolidado, subastas_con_nombre_balance_bess,
    )
    from .ofertas_sscc import leer_ofertas_sscc_consolidado
    from .subastas_accdb import construir_mapa_propietario

    ruta_planilla = Path(ruta_planilla)
    if not ruta_planilla.is_file():
        raise ErrorEntrada(f"No esta {ruta_planilla}.")

    callado = lambda *_: None  # noqa: E731

    registrar(f"Leyendo {ruta_planilla.name}...")
    with abrir_excel(ruta_planilla) as libro:
        ecostos = _leer(libro, HOJA_CALCULO_ECOSTOS,
                        lambda: _hoja_de_calculo(libro, HOJA_CALCULO_ECOSTOS), registrar)
        re545 = _leer(libro, HOJA_CALCULO_RE545,
                      lambda: _hoja_de_calculo(libro, HOJA_CALCULO_RE545), registrar)
        subastas = _leer(libro, HOJA_SUBASTAS,
                         lambda: leer_hoja_con_fila_vacia(libro, HOJA_SUBASTAS), registrar)
        fd = _leer(libro, HOJA_FD, lambda: leer_fd_consolidado(
            ruta_planilla, registrar=callado, libro=libro), registrar)
        ofertas = _leer(libro, HOJA_OFERTAS_SSCC, lambda: leer_ofertas_sscc_consolidado(
            ruta_planilla, registrar=callado, libro=libro)[0], registrar)
        medidores = None
        if ecostos is None:
            medidores = _leer(libro, HOJA_MEDIDORES,
                              lambda: leer_hoja_con_fila_vacia(libro, HOJA_MEDIDORES),
                              registrar)
        try:
            centrales = leer_centrales_planilla(ruta_planilla, libro=libro)
        except Exception:  # noqa: BLE001
            centrales = None

    resumen_bess, diccionario = centrales or (None, None)

    # Las columnas con formula que vengan vacias se calculan (ver
    # completar_formulas).
    from .ecostos import FORMULAS_E_COSTOS, NOMBRES_CALCULO_E_COSTOS
    from .re545 import FORMULAS_RE545, NOMBRES_CALCULO_RE545

    if ecostos is not None:
        ecostos = completar_formulas(ecostos, FORMULAS_E_COSTOS, NOMBRES_CALCULO_E_COSTOS)
    if re545 is not None:
        re545 = completar_formulas(re545, FORMULAS_RE545, NOMBRES_CALCULO_RE545)

    datos = {}  # columna -> {central normalizada: valor}
    nombres = {}  # central normalizada -> como se escribe

    def anotar(df, columna):
        serie = _col(df, columna)
        if serie is not None:
            for v in serie:
                if _texto_seguro(v):
                    nombres.setdefault(normalizar(v), _texto_seguro(v))

    # --- plata y energia de las hojas de calculo -------------------------
    for df, columna in ((ecostos, C_ECOSTOS), (re545, C_RE545)):
        if df is not None:
            anotar(df, "Configuracion")
            datos[columna] = _por_central(df, "Configuracion", "Monto a compensar")
    if ecostos is not None:
        datos[C_DESCARGA] = _por_central(ecostos, "Configuracion", "Descarga kWh")
        # La carga viene negativa: se guarda como energia (positiva), asi
        # "mas carga" es un numero mas grande y su % se lee bien.
        datos[C_CARGA] = {
            c: abs(v) if _hay(v) else v
            for c, v in _por_central(ecostos, "Configuracion", "Carga kWh").items()
        }
        datos[C_CMG] = _por_central(ecostos, "Configuracion", "CMg", "mean")
        adj = _por_central(ecostos, "Configuracion", "Adj SSCC")
        # Adj SSCC va por cuarto de hora: 4 cuartos = 1 hora.
        datos[C_HORAS_ADJ] = {c: v / 4 for c, v in adj.items()}
    elif medidores is not None and _col(medidores, "Gen_Unidad") is not None:
        anotar(medidores, "clave")
        energia = _num(_col(medidores, "Gen_Unidad"))
        clave = _col(medidores, "clave").map(normalizar)
        datos[C_DESCARGA] = energia.clip(lower=0).groupby(clave).sum().to_dict()
        datos[C_CARGA] = (-energia.clip(upper=0)).groupby(clave).sum().to_dict()
    if re545 is not None:
        datos[C_RESERVA] = _por_central(re545, "Configuracion", "Reserva con FD y FMA")

    # C1, C2 y cuantos ciclos/ventanas quedaron con monto.
    for df, c1, c2, extra, grupo, con_monto, del_mes in (
        (ecostos, C_EC_C1, C_EC_C2, (C_EC_AJUSTADA, "Descarga ajustada FD-prorrata kWh"),
         "Ciclo de Carga del mes", C_EC_CICLOS, C_EC_CICLOS_MES),
        (re545, C_RE_C1, C_RE_C2, None,
         "Ventana de valorizacion", C_RE_VENTANAS, C_RE_VENTANAS_MES),
    ):
        if df is None:
            continue
        nombre_c1 = "Total [CLP]" if df is ecostos else "Total C1"
        nombre_c2 = "Total" if df is ecostos else "Total C2"
        datos[c1] = _por_central(df, "Configuracion", nombre_c1)
        datos[c2] = _por_central(df, "Configuracion", nombre_c2)
        if extra is not None:
            datos[extra[0]] = _por_central(df, "Configuracion", extra[1])
        central, ciclo, monto = (
            _col(df, "Configuracion"), _col(df, grupo), _col(df, "Monto a compensar")
        )
        if central is not None and ciclo is not None and monto is not None:
            por_ciclo = (
                pd.DataFrame({"c": central.map(normalizar).to_numpy(),
                              "g": ciclo.to_numpy(), "m": _num(monto).to_numpy()})
                .dropna(subset=["g"]).groupby(["c", "g"])["m"].sum().reset_index()
            )
            por_ciclo = por_ciclo[por_ciclo["c"] != ""]
            datos[del_mes] = por_ciclo.groupby("c").size().astype(float).to_dict()
            datos[con_monto] = (
                por_ciclo.assign(con=por_ciclo["m"] > 0).groupby("c")["con"].sum()
                .astype(float).to_dict()
            )

    # Dias del mes con medidas de la central: los de las dos hojas de
    # calculo juntas (cada cuarto de hora va a una o a la otra), o los de
    # Medidores si no hay calculo. Una central que entro a mitad de mes
    # tiene menos (pedido del usuario: "este mes estuvo todo el mes y el
    # mes pasado solo una parte").
    dias_de = {}
    for df, columna_central in ((ecostos, "Configuracion"), (re545, "Configuracion"),
                                (medidores, "clave")):
        central, dia = _col(df, columna_central), _col(df, "Dia")
        if central is None or dia is None:
            continue
        mes = _col(df, "Mes")
        clave_dia = dia.astype(str) if mes is None else mes.astype(str) + "-" + dia.astype(str)
        for c, d in zip(central.map(normalizar), clave_dia):
            if c:
                dias_de.setdefault(c, set()).add(d)
    if dias_de:
        datos[C_DIAS] = {c: float(len(d)) for c, d in dias_de.items()}

    horas_adjudicadas = {}   # servicio -> {(central, hora del mes)}
    # --- lo adjudicado en Subastas, con el nombre Balance_BESS ---------------
    if subastas is not None and _col(subastas, "Configuración") is not None:
        subastas = subastas_con_nombre_balance_bess(
            subastas, diccionario, registrar=callado
        )
        anotar(subastas, "Configuración")
        servicio = _col(subastas, "Concepto").map(lambda v: _texto_seguro(v)[:3].upper()) \
            if _col(subastas, "Concepto") is not None else None
        horas = _col(subastas, "Hora_mes")
        if horas is not None:
            # Horas distintas con algo adjudicado (cualquier servicio).
            tabla_horas = pd.DataFrame({
                "c": _col(subastas, "Configuración").map(normalizar).to_numpy(),
                "h": horas.to_numpy(),
            }).dropna()
            datos[C_HORAS_SUB] = (
                tabla_horas.drop_duplicates().groupby("c")["h"].size().astype(float).to_dict()
            )
        if horas is not None and servicio is not None:
            centrales_sub = _col(subastas, "Configuración").map(normalizar)
            for codigo in ("CPF", "CSF"):
                es = servicio.eq(codigo).to_numpy()
                horas_adjudicadas[codigo] = {
                    (c, int(h)) for c, h in zip(centrales_sub[es], horas[es])
                    if c and pd.notna(h)
                }
        energia = _col(subastas, "Energía SSCC")
        if servicio is not None and energia is not None:
            for codigo, columna in (("CPF", C_ADJ_CPF), ("CSF", C_ADJ_CSF)):
                parte = subastas[servicio.eq(codigo).to_numpy()]
                datos[columna] = _por_central(parte, "Configuración", "Energía SSCC")

    # --- FD: el factor de desempeno promedio de las horas que existen -------
    if fd is not None:
        for bloque, clave, columna in ((fd[0], "Fd_CSF", C_FD_CSF),
                                       (fd[1], "Fd_CPF", C_FD_CPF)):
            if bloque is None:
                continue
            factor = next((c for c in bloque.columns if clave in str(c)), None)
            if factor is None or _col(bloque, "Central") is None:
                continue
            horas = _col(bloque, "Hora Mes")
            if horas is not None:
                bloque = bloque[_num(horas).notna().to_numpy()]
            datos[columna] = _por_central(bloque, "Central", factor, "mean")
            # "No participó" en la Respuesta (columnas H y V de la hoja):
            # puede venir con FD = 1 (docs/Trazabilidad_FD_a_DB_Y_planilla3.md,
            # seccion 25), asi que el promedio del FD no lo ve.
            respuestas = [c for c in bloque.columns if "respuesta" in normalizar(c)]
            if respuestas:
                no_participo = pd.Series(False, index=range(len(bloque)))
                for nombre in respuestas:
                    no_participo |= (
                        _col(bloque, nombre).reset_index(drop=True)
                        .map(lambda v: normalizar(v).startswith("no particip"))
                    )
                central = _col(bloque, "Central").reset_index(drop=True).map(normalizar)
                tabla_np = pd.DataFrame({"c": central, "np": no_participo})
                tabla_np = tabla_np[tabla_np["c"] != ""]
                sin, total = (
                    (C_NP_CSF, C_FDH_CSF) if clave == "Fd_CSF" else (C_NP_CPF, C_FDH_CPF)
                )
                datos[sin] = tabla_np.groupby("c")["np"].sum().astype(float).to_dict()
                datos[total] = tabla_np.groupby("c").size().astype(float).to_dict()
                codigo = "CSF" if clave == "Fd_CSF" else "CPF"
                horas_fd = _col(bloque, "Hora Mes")
                if codigo in horas_adjudicadas and horas_fd is not None:
                    tabla_np["h"] = _num(horas_fd).reset_index(drop=True)
                    adjudicada = [
                        (c, int(h)) in horas_adjudicadas[codigo] if pd.notna(h) else False
                        for c, h in zip(tabla_np["c"], tabla_np["h"])
                    ]
                    dentro = tabla_np[adjudicada]
                    sin_a, total_a = (
                        (C_NPA_CSF, C_HA_CSF) if codigo == "CSF" else (C_NPA_CPF, C_HA_CPF)
                    )
                    datos[sin_a] = dentro.groupby("c")["np"].sum().astype(float).to_dict()
                    datos[total_a] = dentro.groupby("c").size().astype(float).to_dict()

    # --- ofertas -------------------------------------------------------------
    if ofertas is not None and _col(ofertas, "Oferta completa") is not None:
        datos[C_OFERTA_DIAS] = _por_central(ofertas, "Central", "Oferta completa")

    # --- empresa -------------------------------------------------------------
    empresa_de = {}
    if resumen_bess is not None:
        empresa_de = construir_mapa_propietario(resumen_bess)
        nombre_col = next(
            (c for c in resumen_bess.columns if "nombre" in normalizar(c)), None
        )
        if nombre_col is not None:
            anotar(resumen_bess, nombre_col)

    # Las sumas para la descomposicion en pesos (responsables.py).
    from .responsables import AGREGADOS, agregados

    datos.update(agregados(
        ecostos, re545,
        subastas if subastas is not None and _col(subastas, "Configuración") is not None else None,
    ))

    centrales_norm = sorted(nombres, key=lambda c: nombres[c].upper())
    filas = []
    for c in centrales_norm:
        fila = {C_CENTRAL: nombres[c], C_EMPRESA: empresa_de.get(c, "")}
        for columna in (*COLUMNAS_RESUMEN[2:], *AGREGADOS):
            if columna == C_TOTAL:
                continue
            fila[columna] = datos[columna].get(c, 0.0) if columna in datos else math.nan
        partes = [fila[C_ECOSTOS], fila[C_RE545]]
        fila[C_TOTAL] = (
            math.nan if all(math.isnan(p) for p in partes)
            else sum(0.0 if math.isnan(p) else p for p in partes)
        )
        filas.append(fila)

    return pd.DataFrame(filas, columns=[*COLUMNAS_RESUMEN, *AGREGADOS])


# ------------------------------------------------------------
# Frases
# ------------------------------------------------------------

def _hay(valor):
    return valor is not None and not (isinstance(valor, float) and math.isnan(valor))


def plata(valor):
    """1234567.8 -> '$1.234.568' (como se escribe en Chile)."""

    signo = "-" if valor < 0 else ""
    return f"{signo}${abs(valor):,.0f}".replace(",", ".")


def numero(valor, decimales=0):
    texto = f"{valor:,.{decimales}f}"
    return texto.replace(",", "X").replace(".", ",").replace("X", ".")


def explicar_central(fila):
    """
    Por que una central no recibio nada (o no recibio por un lado), en
    una frase. "" si no hay nada que decir (recibio por los dos).
    """

    total = fila.get(C_TOTAL)
    if not _hay(total):
        return "Falta calcular E Costos y RE545."

    adj_cpf, adj_csf = fila.get(C_ADJ_CPF), fila.get(C_ADJ_CSF)
    adjudicada = (_hay(adj_cpf) and adj_cpf > 0) or (_hay(adj_csf) and adj_csf > 0)
    sin_subastas = not _hay(adj_cpf) and not _hay(adj_csf)
    descarga, carga = fila.get(C_DESCARGA), fila.get(C_CARGA)
    fd_cpf, fd_csf = fila.get(C_FD_CPF), fila.get(C_FD_CSF)
    dias = fila.get(C_OFERTA_DIAS)
    horas_adj = fila.get(C_HORAS_ADJ)

    def motivos():
        """Del mas de fondo al mas fino; el primero que aplica manda."""
        if _hay(descarga) and _hay(carga) and descarga == 0 and carga == 0:
            yield "Sin energía medida en el mes (Medidores = 0)."
        if not sin_subastas and not adjudicada:
            yield "No se adjudicó nada en Subastas en el mes."
        if adjudicada and _hay(horas_adj) and horas_adj == 0:
            yield ("Adjudicada en Subastas, pero ninguna hora cruza con "
                   "Medidores (Adj SSCC = 0): revisar su nombre en el Diccionario.")
        for nombre, adj, sin_a, total_a, sin, total in (
            ("CPF", adj_cpf, fila.get(C_NPA_CPF), fila.get(C_HA_CPF),
             fila.get(C_NP_CPF), fila.get(C_FDH_CPF)),
            ("CSF", adj_csf, fila.get(C_NPA_CSF), fila.get(C_HA_CSF),
             fila.get(C_NP_CSF), fila.get(C_FDH_CSF)),
        ):
            # Mejor en las horas adjudicadas; si no se pudo cruzar, todas.
            que = "horas adjudicadas"
            if not (_hay(total_a) and total_a > 0):
                sin_a, total_a, que = sin, total, "horas"
            if not (_hay(adj) and adj > 0 and _hay(sin_a) and _hay(total_a) and total_a > 0):
                continue
            # Solo si no participo en NINGUNA: con los datos reales de oct,
            # "No participó" en la mayoria de las horas CPF adjudicadas es lo
            # normal (85-95 % en casi todas, y reciben igual).
            if sin_a >= total_a:
                yield (f"Según FD no tuvo participación en {nombre} ('No participó' "
                       f"en todas sus {numero(total_a)} {que}, aunque el FD diga 1).")
        servicios_sin_fd = [
            nombre for nombre, adj, fd in (("CPF", adj_cpf, fd_cpf), ("CSF", adj_csf, fd_csf))
            if _hay(adj) and adj > 0 and _hay(fd) and fd == 0
        ]
        if servicios_sin_fd:
            yield (f"Según FD no tuvo participación (FD "
                   f"{' y '.join(servicios_sin_fd)} = 0 en todo el mes).")
        if _hay(dias) and dias == 0:
            yield "Ningún día con oferta completa."

    if total == 0:
        motivo = next(motivos(), None)
        return motivo or (
            "La valorización no superó el costo en ningún ciclo ni ventana "
            "(Total C1 ≤ Total C2: el monto nunca es negativo)."
        )

    ecostos, re545 = fila.get(C_ECOSTOS), fila.get(C_RE545)
    if _hay(re545) and re545 == 0 and _hay(ecostos) and ecostos > 0:
        if _hay(dias) and dias == 0:
            # Sin oferta completa la energia de la ventana no va a RE545
            # (Medidores!T): esa es la razon de fondo.
            return ("Solo recibe por E Costos: ningún día con oferta completa "
                    "(sin oferta completa, la energía no va a RE545).")
        motivo = next(motivos(), None)
        return "Solo recibe por E Costos" + (f": {motivo[0].lower()}{motivo[1:]}" if motivo else ".")
    if _hay(ecostos) and ecostos == 0 and _hay(re545) and re545 > 0:
        return "Solo recibe por RE545."
    return ""


C_PERDIDA_FD = "Dejó de recibir por FD [CLP]"


def _n_de_m(con, total):
    if not _hay(con) or not _hay(total) or total == 0:
        return ""
    return f"{numero(con)} de {numero(total)}"


def reporte_corrida(resumen):
    """
    (lineas de arriba, tabla por central, ficha) para el control:
      - tabla: lo que recibe cada central por lado, cuantos ciclos y
        ventanas pagaron, cuanto DEJO de recibir por FD < 1 (el mismo
        escenario de la comparacion, con FD = 1: responsables.perdida_por_fd)
        y, si hay algo que decir, por que;
      - ficha: un renglon por central y concepto, con su valor y una alerta
        cuando algo esta fuera de lo normal (ver ficha_corrida).
    """

    from .responsables import perdida_por_fd

    filas = resumen.to_dict("records")
    perdidas = [perdida_por_fd(f) for f in filas]
    tabla = resumen[[C_CENTRAL, C_EMPRESA, C_ECOSTOS, C_RE545, C_TOTAL]].copy()
    tabla["Ciclos con monto"] = [
        _n_de_m(f.get(C_EC_CICLOS), f.get(C_EC_CICLOS_MES)) for f in filas
    ]
    tabla["Ventanas con monto"] = [
        _n_de_m(f.get(C_RE_VENTANAS), f.get(C_RE_VENTANAS_MES)) for f in filas
    ]
    tabla[C_PERDIDA_FD] = [
        (math.nan if any(not _hay(v) for v in p.values()) and not any(_hay(v) and v > 0 for v in p.values())
         else sum(v for v in p.values() if _hay(v)))
        if p else math.nan
        for p in perdidas
    ]
    tabla["Observación"] = [explicar_central(f) for f in filas]

    total = resumen[C_TOTAL].sum(min_count=1)
    sin_ingreso = resumen[resumen[C_TOTAL].fillna(-1) == 0][C_CENTRAL].tolist()
    lineas = []
    if _hay(total):
        lineas.append(f"Total del mes: {plata_corta(total)} entre {len(resumen)} centrales "
                      f"(E Costos {plata_corta(resumen[C_ECOSTOS].sum())} · "
                      f"RE545 {plata_corta(resumen[C_RE545].sum())}).")
        mayores = resumen.sort_values(C_TOTAL, ascending=False).head(3)
        lineas.append("Las que más reciben: " + ", ".join(
            f"{f[C_CENTRAL]} {plata_corta(f[C_TOTAL])} ({numero(f[C_TOTAL] / total * 100)} %)"
            for _, f in mayores.iterrows() if total
        ) + ".")
    perdida = tabla[C_PERDIDA_FD].sum(min_count=1)
    if _hay(perdida) and perdida >= CAMBIO_MINIMO_CLP:
        mas = tabla.dropna(subset=[C_PERDIDA_FD]).sort_values(C_PERDIDA_FD, ascending=False).head(3)
        lineas.append(
            f"Por FD bajo 1 se dejaron de recibir {plata_corta(perdida)} en el mes "
            f"(con FD = 1 se habría recibido {plata_corta(total + perdida)}). Las que más: "
            + ", ".join(f"{f[C_CENTRAL]} {plata_corta(f[C_PERDIDA_FD])}" for _, f in mas.iterrows())
            + "."
        )
    if sin_ingreso:
        lineas.append(
            f"{len(sin_ingreso)} central(es) sin ingreso: {', '.join(sin_ingreso)} "
            f"(el motivo va en su fila)."
        )
    elif _hay(total):
        lineas.append("Todas las centrales reciben algo.")
    alertas = ficha_corrida(resumen)
    con_alerta = alertas[alertas["Alerta"] != ""]
    if len(con_alerta):
        lineas.append(
            f"{con_alerta[C_CENTRAL].nunique()} central(es) con alertas en la ficha "
            f"({len(con_alerta)} en total): ver la hoja 'Reporte ficha'."
        )
    return lineas, tabla, alertas


# La ficha: (columna del resumen, concepto, unidad). Va todo lo que la
# central tiene (lo que es 0 o no esta, no va), con su alerta.
_FICHA = (
    (C_TOTAL, "Recibe total", "CLP"),
    (C_ECOSTOS, "Recibe E Costos", "CLP"),
    (C_RE545, "Recibe RE545", "CLP"),
    (C_EC_CICLOS, "E Costos: ciclos con monto", "ciclos"),
    (C_EC_C1, "E Costos: Total C1", "CLP"),
    (C_EC_C2, "E Costos: Total C2", "CLP"),
    (C_RE_VENTANAS, "RE545: ventanas con monto", "ventanas"),
    (C_RE_C1, "RE545: Total C1", "CLP"),
    (C_RE_C2, "RE545: Total C2", "CLP"),
    (C_ADJ_CPF, "Adjudicado CPF", "MW"),
    (C_ADJ_CSF, "Adjudicado CSF", "MW"),
    (C_HORAS_SUB, "Horas adjudicadas", "h"),
    (C_FD_CPF, "FD CPF promedio", "FD"),
    (C_FD_CSF, "FD CSF promedio", "FD"),
    (C_NPA_CPF, "CPF: horas adjudicadas con 'No participó'", "h"),
    (C_NPA_CSF, "CSF: horas adjudicadas con 'No participó'", "h"),
    (C_OFERTA_DIAS, "Días con oferta completa", "días"),
    (C_DIAS, "Días en el mes", "días"),
    (C_DESCARGA, "Descarga", "kWh"),
    (C_CARGA, "Carga", "kWh"),
    (C_RESERVA, "Reserva con FD y FMA", "kWh"),
    (C_CMG, "CMg promedio", "CLP/kWh"),
)


def _alerta(columna, fila):
    """El texto de alerta de un concepto de la ficha, o ''."""

    valor = fila.get(columna)
    if columna == C_TOTAL and _hay(valor) and valor == 0:
        return "Sin ingreso"
    for adj, fd, sin, total, servicio in (
        (C_ADJ_CPF, C_FD_CPF, C_NPA_CPF, C_HA_CPF, "CPF"),
        (C_ADJ_CSF, C_FD_CSF, C_NPA_CSF, C_HA_CSF, "CSF"),
    ):
        adjudicada = _hay(fila.get(adj)) and fila.get(adj) > 0
        if not adjudicada:
            continue
        if columna == sin and _hay(valor) and _hay(fila.get(total)) and fila[total] > 0:
            # Solo NINGUNA: la mayoria de las horas CPF con "No participó"
            # es lo normal (ver explicar_central).
            if valor >= fila[total]:
                return f"No participó en ninguna de sus horas con {servicio} adjudicado"
        if columna == fd and _hay(valor) and valor < 0.5:
            return f"FD {servicio} bajo 0,5 con {servicio} adjudicado"
    adjudicada = any(_hay(fila.get(c)) and fila.get(c) > 0 for c in (C_ADJ_CPF, C_ADJ_CSF))
    sin_ofertas = _hay(fila.get(C_OFERTA_DIAS)) and fila.get(C_OFERTA_DIAS) == 0
    if columna == C_OFERTA_DIAS and sin_ofertas and adjudicada:
        return "Sin oferta completa: su energía no va a RE545"
    # Solo si todo pasa por E Costos (sin oferta completa): si no, E Costos
    # cubre pocos dias y tener pocos ciclos con monto es lo normal.
    if columna == C_EC_CICLOS and sin_ofertas and _hay(valor) \
            and _hay(fila.get(C_EC_CICLOS_MES)) and fila[C_EC_CICLOS_MES] > 0 \
            and (fila.get(C_DESCARGA) or 0) > 0 and valor < fila[C_EC_CICLOS_MES] / 4:
        return "Menos de 1 de cada 4 ciclos con monto (C2 sobre C1 en el resto)"
    if columna == C_HORAS_SUB and _hay(valor) and valor > 0 \
            and _hay(fila.get(C_HORAS_ADJ)) and abs(fila[C_HORAS_ADJ] - valor) >= 1:
        return ("Las horas con Adj SSCC no coinciden con las adjudicadas: "
                "revisar su nombre en el Diccionario")
    return ""


def ficha_corrida(resumen):
    """Un renglon por central y concepto (Central, Concepto, Valor, Unidad, Alerta)."""

    filas = []
    for f in resumen.to_dict("records"):
        for columna, concepto, unidad in _FICHA:
            valor = f.get(columna)
            alerta = _alerta(columna, f)
            if (not _hay(valor) or valor == 0) and not alerta and columna != C_TOTAL:
                continue
            total = {C_EC_CICLOS: C_EC_CICLOS_MES, C_RE_VENTANAS: C_RE_VENTANAS_MES,
                     C_NPA_CPF: C_HA_CPF, C_NPA_CSF: C_HA_CSF}.get(columna)
            filas.append({
                C_CENTRAL: f[C_CENTRAL], "Concepto": concepto,
                "Valor": valor, "De": f.get(total) if total else math.nan,
                "Unidad": unidad, "Alerta": alerta,
            })
    return pd.DataFrame(filas, columns=[C_CENTRAL, "Concepto", "Valor", "De", "Unidad", "Alerta"])


def formato_ficha(ws):
    """La ficha: numero segun su unidad, alerta en rojo, franja por central, filtro."""

    from openpyxl.styles import Font, PatternFill

    fila_enc = FILA_ENCABEZADO_REPORTE
    columnas = {str(c.value): c.column for c in ws[fila_enc] if c.value is not None}
    if "Alerta" not in columnas or "Concepto" not in columnas:
        return
    franja = PatternFill(fill_type="solid", start_color="F5F8FC", end_color="F5F8FC")
    rojo = PatternFill(fill_type="solid", start_color="F8CBAD", end_color="F8CBAD")
    formatos = {"CLP": "#,##0", "FD": "0.00", "CLP/kWh": "#,##0.0"}
    anterior, banda = None, False
    for fila in range(fila_enc + 1, ws.max_row + 1):
        central = ws.cell(fila, columnas[C_CENTRAL]).value
        if central is None:
            continue
        if central != anterior:
            banda, anterior = not banda, central
        formato = formatos.get(ws.cell(fila, columnas["Unidad"]).value, "#,##0")
        for nombre in ("Valor", "De"):
            ws.cell(fila, columnas[nombre]).number_format = formato if nombre == "Valor" else "#,##0"
        if banda:
            for col in range(min(columnas.values()), max(columnas.values()) + 1):
                ws.cell(fila, col).fill = franja
        alerta = ws.cell(fila, columnas["Alerta"])
        if alerta.value:
            alerta.fill = rojo
            alerta.font = Font(color="9C0006", bold=True)
    for nombre, ancho in ((C_CENTRAL, 30), ("Concepto", 32), ("Valor", 18), ("De", 8),
                          ("Unidad", 10), ("Alerta", 70)):
        ws.column_dimensions[ws.cell(fila_enc, columnas[nombre]).column_letter].width = ancho
    primera = ws.cell(fila_enc, min(columnas.values())).coordinate
    ultima = ws.cell(ws.max_row, max(columnas.values())).coordinate
    ws.auto_filter.ref = f"{primera}:{ultima}"
    ws.freeze_panes = ws.cell(fila_enc + 1, min(columnas.values()) + 2).coordinate


# ------------------------------------------------------------
# Comparacion entre dos planillas
# ------------------------------------------------------------

# Que mira cada lado para explicar un cambio, y como se dice.
_MOTORES_RE545 = (
    (C_NP_CPF, "horas FD CPF 'No participó'", "h"),
    (C_NP_CSF, "horas FD CSF 'No participó'", "h"),
    (C_FD_CPF, "FD CPF promedio", ""),
    (C_FD_CSF, "FD CSF promedio", ""),
    (C_ADJ_CPF, "adjudicado CPF", "MW"),
    (C_ADJ_CSF, "adjudicado CSF", "MW"),
    (C_HORAS_SUB, "horas adjudicadas", "h"),
    (C_RESERVA, "reserva con FD y FMA", "kWh"),
    (C_OFERTA_DIAS, "días con oferta completa", "días"),
    (C_DIAS, "días en el mes", "días"),
    (C_CMG, "CMg promedio", "CLP/kWh"),
)
_MOTORES_ECOSTOS = (
    (C_NP_CPF, "horas FD CPF 'No participó'", "h"),
    (C_NP_CSF, "horas FD CSF 'No participó'", "h"),
    (C_FD_CPF, "FD CPF promedio", ""),
    (C_FD_CSF, "FD CSF promedio", ""),
    (C_ADJ_CPF, "adjudicado CPF", "MW"),
    (C_ADJ_CSF, "adjudicado CSF", "MW"),
    (C_HORAS_SUB, "horas adjudicadas", "h"),
    (C_EC_AJUSTADA, "descarga ajustada FD-prorrata", "kWh"),
    (C_DESCARGA, "descarga", "kWh"),
    (C_CARGA, "carga", "kWh"),
    (C_CMG, "CMg promedio", "CLP/kWh"),
    (C_DIAS, "días en el mes", "días"),
)


# Lo que se resta en el monto (Total C2): moverse en contra lo explica.
_RESTAN = (C_EC_C2, C_RE_C2, C_NP_CPF, C_NP_CSF)


def _pct(nuevo, viejo):
    if not _hay(nuevo) or not _hay(viejo):
        return math.nan
    if viejo == 0:
        return math.inf if nuevo != 0 else 0.0
    return (nuevo - viejo) / abs(viejo) * 100.0


def _texto_pct(pct):
    if math.isinf(pct):
        return "aparece"
    return f"{'+' if pct > 0 else ''}{numero(pct, 0)} %"


def _motores(actual, anterior, motores, sube):
    """Las 2 entradas que mas cambiaron en el mismo sentido que la plata."""

    candidatos = []
    for columna, nombre, unidad in motores:
        nuevo, viejo = actual.get(columna), anterior.get(columna)
        pct = _pct(nuevo, viejo)
        if not _hay(pct) or abs(pct) < CAMBIO_MINIMO_PCT:
            continue
        crece = nuevo > viejo
        # C2 se RESTA en el monto: si baja, el monto sube.
        empuja_arriba = (not crece) if columna in _RESTAN else crece
        if empuja_arriba != sube:
            continue
        decimales = 2 if columna in (C_FD_CPF, C_FD_CSF) else 0
        viejo_txt = numero(viejo, decimales)
        nuevo_txt = numero(nuevo, decimales)
        if unidad == "CLP":
            viejo_txt, nuevo_txt, unidad = plata(viejo), plata(nuevo), ""
        detalle = f"{viejo_txt} → {nuevo_txt}{(' ' + unidad) if unidad else ''}"
        plural = nombre.startswith(("los ", "las "))
        verbo = ("subieron" if crece else "bajaron") if plural else (
            "subió" if crece else "bajó"
        )
        candidatos.append((abs(pct) if not math.isinf(pct) else 1e9,
                           f"{nombre} {verbo} ({detalle})"))
    candidatos.sort(reverse=True)
    return [texto for _, texto in candidatos[:2]]


def plata_corta(valor):
    """1.234.567.890 -> '$1.235 MM'; los montos chicos, enteros."""

    if _hay(valor) and abs(valor) >= 1e6:
        signo = "-" if valor < 0 else ""
        return f"{signo}${numero(abs(valor) / 1e6, 0)} MM"
    return plata(valor) if _hay(valor) else "sin dato"


# Como se arma el monto de cada lado: (ciclos o ventanas con monto, del
# mes, Total C1, Total C2).
_MECANICA = {
    "E Costos": (C_EC_CICLOS, C_EC_CICLOS_MES, C_EC_C1, C_EC_C2, "ciclos"),
    "RE545": (C_RE_VENTANAS, C_RE_VENTANAS_MES, C_RE_C1, C_RE_C2, "ventanas"),
}


def _mecanica(lado, actual, anterior):
    """'ciclos con monto 4 → 30 de 32; C1 $518 MM → $471 MM y C2 ...'."""

    con, del_mes, c1, c2, que = _MECANICA[lado]
    partes = []
    nuevo, viejo = actual.get(con), anterior.get(con)
    if _hay(nuevo) and _hay(viejo) and nuevo != viejo:
        total = actual.get(del_mes)
        partes.append(
            f"{que} con monto {numero(viejo)} → {numero(nuevo)}"
            + (f" de {numero(total)}" if _hay(total) else "")
        )
    totales = []
    for columna, nombre in ((c1, "C1"), (c2, "C2")):
        nuevo, viejo = actual.get(columna), anterior.get(columna)
        if _hay(nuevo) and _hay(viejo) and (nuevo or viejo):
            totales.append((nombre, viejo, nuevo, abs(_pct(nuevo, viejo))))
    if any(_hay(p) and p >= CAMBIO_MINIMO_PCT for *_, p in totales):
        partes.append(" y ".join(
            f"{nombre} {plata_corta(viejo)} → {plata_corta(nuevo)}"
            for nombre, viejo, nuevo, _ in totales
        ))
    return "; ".join(partes)


def _cambio_general(actual, anterior, columna):
    """
    La variacion % "de todo el sistema" de una entrada (la mediana entre
    las centrales que estan en los dos meses), o NaN. Sirve para no
    repetir en cada central algo que le paso a todas (el CMg).
    """

    a = actual.set_index(actual[C_CENTRAL].map(normalizar))[columna]
    b = anterior.set_index(anterior[C_CENTRAL].map(normalizar))[columna]
    pcts = [
        _pct(a[c], b[c]) for c in set(a.index) & set(b.index)
        if _hay(a[c]) and _hay(b[c]) and b[c] != 0
    ]
    pcts = [p for p in pcts if _hay(p) and not math.isinf(p)]
    return float(pd.Series(pcts).median()) if len(pcts) >= 3 else math.nan


# Entradas que suelen moverse para todas las centrales a la vez: se
# dicen una vez arriba, y en una central solo si se aparta de eso.
_GENERALES = (C_CMG,)
_APARTE_PCT = 15.0


def _entradas_que_mas_cambiaron(actual, anterior, motores, cuantas=3, generales=None):
    """'CMg promedio 22,1 → 6,8 CLP/kWh (-69 %), ...': las que mas se
    movieron (10 % o mas), en cualquier sentido."""

    mecanica = {c for valores in _MECANICA.values() for c in valores[:4]}
    candidatos = []
    for columna, nombre, unidad in motores:
        if columna in mecanica:
            continue
        nuevo, viejo = actual.get(columna), anterior.get(columna)
        pct = _pct(nuevo, viejo)
        if not _hay(pct) or abs(pct) < CAMBIO_MINIMO_PCT:
            continue
        general = (generales or {}).get(columna)
        if _hay(general) and not math.isinf(pct) and abs(pct - general) < _APARTE_PCT:
            continue  # le paso a todas: va en el resumen de arriba
        decimales = 2 if columna in (C_FD_CPF, C_FD_CSF) else (1 if columna == C_CMG else 0)
        texto = (
            f"{nombre} {numero(viejo, decimales)} → {numero(nuevo, decimales)}"
            + (f" {unidad}" if unidad else "")
            + (f" ({_texto_pct(pct)})" if not math.isinf(pct) else "")
        )
        candidatos.append((abs(pct) if not math.isinf(pct) else 1e9, texto))
    candidatos.sort(reverse=True)
    return [texto for _, texto in candidatos[:cuantas]]


def _frase_lado(lado, nuevo, viejo, actual, anterior, motores, generales=None):
    """
    Una frase corta por lado: cuanto cambio, como se movio el monto
    (ciclos/ventanas con monto, Total C1 y C2; el monto es max(0, C1 - C2)
    por ciclo o ventana) y las entradas que mas cambiaron. El detalle
    completo va en la hoja "Cambios por central".
    """

    delta = (nuevo if _hay(nuevo) else 0.0) - (viejo if _hay(viejo) else 0.0)
    if abs(delta) < CAMBIO_MINIMO_CLP:
        return None
    pct = _pct(nuevo, viejo)
    cabeza = (
        f"{lado} {'sube' if delta > 0 else 'baja'} {plata_corta(abs(delta))}"
        + (f" ({_texto_pct(pct)})" if _hay(pct) else "")
    )
    mecanica = _mecanica(lado, actual, anterior)
    entradas = _entradas_que_mas_cambiaron(actual, anterior, motores, generales=generales)
    texto = cabeza + (f": {mecanica}" if mecanica else "")
    if not mecanica and not entradas:
        texto += ", sin un cambio grande en sus entradas"
    return texto + ".", entradas


def comparar_resumenes(actual, anterior):
    """
    (lineas de arriba, tabla por central con los dos meses y la
    explicacion). actual/anterior: resumen_por_central() de cada planilla.
    """

    a = actual.set_index(actual[C_CENTRAL].map(normalizar))
    b = anterior.set_index(anterior[C_CENTRAL].map(normalizar))
    claves = sorted(set(a.index) | set(b.index),
                    key=lambda c: (a[C_CENTRAL].get(c) or b[C_CENTRAL].get(c)).upper())

    generales = {c: _cambio_general(actual, anterior, c) for c in _GENERALES}
    filas, sin_cambios = [], []
    for clave in claves:
        fila_a = a.loc[clave].to_dict() if clave in a.index else {}
        fila_b = b.loc[clave].to_dict() if clave in b.index else {}
        nombre = fila_a.get(C_CENTRAL) or fila_b.get(C_CENTRAL)
        empresa = fila_a.get(C_EMPRESA) or fila_b.get(C_EMPRESA) or ""
        total_a = fila_a.get(C_TOTAL, math.nan)
        total_b = fila_b.get(C_TOTAL, math.nan)

        entradas = []
        if not fila_b:
            explicacion = f"Nueva: no estaba en la planilla anterior. Recibe {plata(total_a)}." \
                if _hay(total_a) else "Nueva: no estaba en la planilla anterior."
            # Si no recibe (o solo por un lado), el porque, como en el
            # reporte de la corrida (pedido del usuario: "tenia
            # adjudicacion, pero en todos los FD salia que no participa").
            motivo = explicar_central(fila_a)
            if motivo:
                explicacion += f" {motivo}"
        elif not fila_a:
            explicacion = (f"Ya no está en esta planilla (antes recibía "
                           f"{plata(total_b)}).") if _hay(total_b) else "Ya no está en esta planilla."
        else:
            lados = [
                _frase_lado("RE545", fila_a.get(C_RE545), fila_b.get(C_RE545),
                            fila_a, fila_b, _MOTORES_RE545, generales),
                _frase_lado("E Costos", fila_a.get(C_ECOSTOS), fila_b.get(C_ECOSTOS),
                            fila_a, fila_b, _MOTORES_ECOSTOS, generales),
            ]
            lados = [l for l in lados if l]
            frases = [frase for frase, _ in lados]
            # Lo que mas cambio, sin repetir lo que comparten los dos lados.
            for _, de_un_lado in lados:
                for entrada in de_un_lado:
                    if entrada not in entradas:
                        entradas.append(entrada)
            entradas = entradas[:4]
            if frases and _hay(total_a) and total_a == 0:
                # Quedo en cero: el porque.
                frases.append(f"Queda en $0: {explicar_central(fila_a)}")
            if not frases and _hay(total_a) and total_a == 0:
                frases.append(f"Sigue en $0. {explicar_central(fila_a)}")
            elif not frases:
                sin_cambios.append(nombre)
            explicacion = "\n".join(frases)

        delta = (total_a if _hay(total_a) else 0.0) - (total_b if _hay(total_b) else 0.0)
        filas.append({
            C_CENTRAL: nombre, C_EMPRESA: empresa,
            "Total anterior [CLP]": total_b, "Total actual [CLP]": total_a,
            "Diferencia [CLP]": delta, "Diferencia [%]": _pct(total_a, total_b),
            "Qué pasó": explicacion,
            "Lo que más cambió": "\n".join(f"• {e}" for e in entradas),
        })

    tabla = pd.DataFrame(filas)
    # Los responsables en pesos (responsables.py): el principal y el
    # desglose, en vez de solo lo que se movio.
    from .responsables import tablas_responsables

    responsables_mes, _, textos = tablas_responsables(actual, anterior)
    if len(tabla):
        tabla.insert(
            tabla.columns.get_loc("Qué pasó"), "Principal responsable",
            [textos.get(c, ("", ""))[0] for c in tabla[C_CENTRAL]],
        )
        tabla.insert(
            tabla.columns.get_loc("Qué pasó"), "Desglose en $",
            [textos.get(c, ("", ""))[1] for c in tabla[C_CENTRAL]],
        )
        tabla = tabla.drop(columns=["Lo que más cambió"])
    if len(tabla):
        tabla["Diferencia [%]"] = tabla["Diferencia [%]"].replace([math.inf], math.nan)
        # Lo que mas cambio, arriba; sin cambios, al final.
        tabla = tabla.reindex(
            tabla["Diferencia [CLP]"].abs().sort_values(ascending=False, kind="stable").index
        ).reset_index(drop=True)

    total_a = actual[C_TOTAL].sum(min_count=1)
    total_b = anterior[C_TOTAL].sum(min_count=1)
    lineas = []
    if _hay(total_a) and _hay(total_b):
        delta = total_a - total_b
        lineas.append(
            f"Total del mes: {plata_corta(total_b)} → {plata_corta(total_a)} "
            f"({'+' if delta >= 0 else '-'}{plata_corta(abs(delta))}"
            + (f", {_texto_pct(_pct(total_a, total_b))}" if total_b else "") + ")."
        )
        for lado, col in (("RE545", C_RE545), ("E Costos", C_ECOSTOS)):
            va, vb = actual[col].sum(), anterior[col].sum()
            lineas.append(f"{lado}: {plata_corta(vb)} → {plata_corta(va)}.")
    if len(responsables_mes):
        principales = [
            f"{'+' if fila['Aporte [CLP]'] >= 0 else '-'}{plata_corta(abs(fila['Aporte [CLP]']))} "
            f"{fila['Responsable']}"
            for _, fila in responsables_mes.head(4).iterrows()
        ]
        lineas.append("Responsables del cambio del mes: " + ", ".join(principales)
                      + " (detalle en la hoja 'Responsables').")

    if len(tabla):
        cambio = tabla[tabla["Diferencia [CLP]"].abs() >= CAMBIO_MINIMO_CLP]
        if len(cambio):
            mayor = cambio.iloc[0]
            lineas.append(
                f"El mayor cambio: {mayor[C_CENTRAL]} "
                f"({'+' if mayor['Diferencia [CLP]'] >= 0 else '-'}"
                f"{plata_corta(abs(mayor['Diferencia [CLP]']))})."
            )
    if sin_cambios:
        lineas.append(f"Sin cambios relevantes: {', '.join(sin_cambios)}.")
    return lineas, tabla


def comparar_planillas(ruta_actual, ruta_anterior, registrar=print):
    """
    Compara dos planillas 11_PAGOS_BESS y escribe el reporte en la carpeta
    Log del caso de la actual (Comparacion_<actual>_vs_<anterior>.xlsx).
    Devuelve (ruta del reporte, lineas de arriba, tabla por central).
    """

    ruta_actual, ruta_anterior = Path(ruta_actual), Path(ruta_anterior)
    actual = resumen_por_central(ruta_actual, registrar)
    anterior = resumen_por_central(ruta_anterior, registrar)
    lineas, tabla = comparar_resumenes(actual, anterior)

    from .parametros import CARPETA_LOG

    carpeta = ruta_actual.parent / CARPETA_LOG
    carpeta.mkdir(parents=True, exist_ok=True)
    destino = carpeta / (
        f"Comparacion_{_periodo(ruta_actual)}_vs_{_periodo(ruta_anterior)}.xlsx"
    )
    cambios = cambios_por_central(actual, anterior, list(tabla[C_CENTRAL]))
    from .responsables import tablas_responsables

    mes, por_central, _ = tablas_responsables(actual, anterior, list(tabla[C_CENTRAL]))
    escribir_reporte(destino, [
        ("Comparacion", lineas, tabla,
         f"{ruta_actual.name} contra {ruta_anterior.name}"),
        ("Responsables", mes, por_central,
         "De dónde sale la diferencia del mes (la suma da la diferencia)"),
        ("Cambios por central", [], cambios,
         "Cada concepto de cada central: anterior, actual y si cambió"),
    ])
    registrar(f"Reporte: {destino}")
    return destino, lineas, tabla


def _periodo(ruta):
    """'11_PAGOS_BESS_2610.xlsx' -> '2610' (o el nombre sin extension)."""

    import re

    hallado = re.search(r"(\d{4})$", Path(ruta).stem)
    return hallado.group(1) if hallado else Path(ruta).stem


# La sub tabla de cambios (pedido del usuario: "que se vea mas ordenado
# las diferencias [...] una sub tabla [...] por cada cosa que diga el
# cambio y su variacion y si cambio o no"): un renglon por central y
# concepto, en este orden -- la plata, como se arma el monto y despues
# las entradas que lo mueven. (columna, concepto, unidad, va siempre).
CONCEPTOS_CAMBIO = (
    (C_TOTAL, "Recibe total", "CLP", True),
    (C_ECOSTOS, "Recibe E Costos", "CLP", True),
    (C_RE545, "Recibe RE545", "CLP", True),
    (C_EC_C1, "E Costos: Total C1", "CLP", False),
    (C_EC_C2, "E Costos: Total C2", "CLP", False),
    (C_EC_CICLOS, "E Costos: ciclos con monto", "ciclos", False),
    (C_RE_C1, "RE545: Total C1", "CLP", False),
    (C_RE_C2, "RE545: Total C2", "CLP", False),
    (C_RE_VENTANAS, "RE545: ventanas con monto", "ventanas", False),
    (C_FD_CPF, "FD CPF promedio", "FD", True),
    (C_FD_CSF, "FD CSF promedio", "FD", True),
    (C_NP_CPF, "FD CPF: horas 'No participó'", "h", False),
    (C_NP_CSF, "FD CSF: horas 'No participó'", "h", False),
    (C_DIAS, "Días en el mes", "días", False),
    (C_HORAS_SUB, "Horas adjudicadas", "h", False),
    (C_ADJ_CPF, "Adjudicado CPF", "MW", False),
    (C_ADJ_CSF, "Adjudicado CSF", "MW", False),
    (C_OFERTA_DIAS, "Días con oferta completa", "días", False),
    # Es lo mismo que "Horas adjudicadas" cuando el nombre de la central
    # cruza bien entre Subastas y Medidores: va solo si no coincide.
    (C_HORAS_ADJ, "Horas con Adj SSCC (no coincide con las adjudicadas)", "h", False),
    (C_DESCARGA, "Descarga", "kWh", False),
    (C_CARGA, "Carga", "kWh", False),
    (C_EC_AJUSTADA, "Descarga ajustada FD-prorrata", "kWh", False),
    (C_RESERVA, "Reserva con FD y FMA", "kWh", False),
    (C_CMG, "CMg promedio", "CLP/kWh", False),
)

COLUMNAS_CAMBIOS = (
    C_CENTRAL, "Concepto", "Unidad", "Anterior", "Actual", "Diferencia",
    "Variación [%]", "Cambio",
)
SUBE, BAJA, IGUAL, SIN_DATO = "▲ Sube", "▼ Baja", "Igual", "Sin dato"
NUEVO, YA_NO = "Nuevo", "Ya no está"


def _cambio(nuevo, viejo):
    """(diferencia, variacion %, 'Sube'/'Baja'/'Igual'/'Sin dato')."""

    if not _hay(nuevo) or not _hay(viejo):
        return math.nan, math.nan, SIN_DATO
    diferencia = nuevo - viejo
    pct = _pct(nuevo, viejo)
    igual = (abs(pct) < 1.0) if not math.isinf(pct) else False
    if igual or diferencia == 0:
        return diferencia, (0.0 if not math.isinf(pct) else math.nan), IGUAL
    return (diferencia, pct if not math.isinf(pct) else math.nan,
            SUBE if diferencia > 0 else BAJA)


def cambios_por_central(actual, anterior, orden=None):
    """
    La sub tabla: un renglon por central y concepto (CONCEPTOS_CAMBIO),
    con anterior, actual, diferencia, variacion y si cambio. Sin ruido:
    un concepto que es 0 o no esta en los dos meses no va, salvo la
    plata y el FD (que van siempre). orden: las centrales en ese orden
    (el de la tabla de la comparacion, lo que mas cambio arriba).
    """

    a = actual.set_index(actual[C_CENTRAL].map(normalizar))
    b = anterior.set_index(anterior[C_CENTRAL].map(normalizar))
    claves = list(dict.fromkeys(
        [normalizar(c) for c in (orden or [])] + sorted(set(a.index) | set(b.index))
    ))
    filas = []
    for clave in claves:
        en_a, en_b = clave in a.index, clave in b.index
        if not en_a and not en_b:
            continue
        fila_a = a.loc[clave].to_dict() if en_a else {}
        fila_b = b.loc[clave].to_dict() if en_b else {}
        nombre = fila_a.get(C_CENTRAL) or fila_b.get(C_CENTRAL)
        for columna, concepto, unidad, siempre in CONCEPTOS_CAMBIO:
            nuevo, viejo = fila_a.get(columna, math.nan), fila_b.get(columna, math.nan)
            if columna == C_HORAS_ADJ and all(
                not _hay(fila.get(C_HORAS_ADJ)) or not _hay(fila.get(C_HORAS_SUB))
                or abs(fila[C_HORAS_ADJ] - fila[C_HORAS_SUB]) < 1
                for fila in (fila_a, fila_b) if fila
            ):
                continue
            vacio_a = not _hay(nuevo) or nuevo == 0
            vacio_b = not _hay(viejo) or viejo == 0
            if vacio_a and vacio_b and not siempre:
                continue
            if not en_b:
                diferencia, pct, cambio = math.nan, math.nan, NUEVO
            elif not en_a:
                diferencia, pct, cambio = math.nan, math.nan, YA_NO
            else:
                diferencia, pct, cambio = _cambio(nuevo, viejo)
            filas.append({
                C_CENTRAL: nombre, "Concepto": concepto, "Unidad": unidad,
                "Anterior": viejo, "Actual": nuevo, "Diferencia": diferencia,
                "Variación [%]": pct, "Cambio": cambio,
            })
    return pd.DataFrame(filas, columns=list(COLUMNAS_CAMBIOS))


def formato_cambios(ws):
    """
    La hoja de la sub tabla: cada numero con el formato de su unidad, la
    columna Cambio en colores (Sube verde, Baja rojo, Igual y Sin dato en
    gris), una franja por central para que se lean juntas y filtro en el
    encabezado (para mirar una central o solo lo que cambio).
    """

    from openpyxl.styles import Font, PatternFill

    def relleno(color):
        return PatternFill(fill_type="solid", start_color=color, end_color=color)

    # Verde y rojo solo en la plata (lo que recibe la central): en lo
    # demas "sube" no es ni bueno ni malo (que el Total C2 baje sube el
    # monto), asi que va en gris, mas marcado si cambio 10 % o mas.
    plata_sube = (relleno("E2EFDA"), Font(color="375623", bold=True))
    plata_baja = (relleno("F8CBAD"), Font(color="9C0006", bold=True))
    fuerte = (None, Font(color="262626", bold=True))
    leve = (None, Font(color="808080"))
    colores = {
        IGUAL: (None, Font(color="A6A6A6")),
        SIN_DATO: (None, Font(color="A6A6A6", italic=True)),
        NUEVO: (relleno("DDEBF7"), Font(color="1F4E78", bold=True)),
        YA_NO: (relleno("EDEDED"), Font(color="595959", bold=True)),
    }
    franja = relleno("F5F8FC")
    fila_enc = FILA_ENCABEZADO_REPORTE
    columnas = {
        str(c.value): c.column for c in ws[fila_enc] if c.value is not None
    }
    if "Cambio" not in columnas:
        return
    primera_col = min(columnas.values())
    ultima_col = max(columnas.values())
    formatos = {"CLP": "#,##0", "FD": "0.00", "CLP/kWh": "#,##0.0"}
    anterior_central, banda = None, False
    for fila in range(fila_enc + 1, ws.max_row + 1):
        central = ws.cell(fila, columnas[C_CENTRAL]).value
        if central is None:
            continue
        if central != anterior_central:
            banda, anterior_central = not banda, central
        unidad = ws.cell(fila, columnas["Unidad"]).value
        formato = formatos.get(unidad, "#,##0")
        for nombre in ("Anterior", "Actual", "Diferencia"):
            ws.cell(fila, columnas[nombre]).number_format = formato
        ws.cell(fila, columnas["Variación [%]"]).number_format = '+0.0" %";-0.0" %";0" %"'
        if banda:
            for col in range(primera_col, ultima_col + 1):
                ws.cell(fila, col).fill = franja
        celda = ws.cell(fila, columnas["Cambio"])
        concepto = str(ws.cell(fila, columnas["Concepto"]).value or "")
        variacion = ws.cell(fila, columnas["Variación [%]"]).value
        if celda.value in (SUBE, BAJA):
            if concepto.startswith("Recibe"):
                fondo, letra = plata_sube if celda.value == SUBE else plata_baja
            elif isinstance(variacion, (int, float)) and abs(variacion) < CAMBIO_MINIMO_PCT:
                fondo, letra = leve
            else:
                fondo, letra = fuerte
        else:
            fondo, letra = colores.get(celda.value, (None, None))
        if fondo is not None:
            celda.fill = fondo
        if letra is not None:
            celda.font = letra
    for nombre, ancho in ((C_CENTRAL, 30), ("Concepto", 34), ("Unidad", 10),
                          ("Anterior", 18), ("Actual", 18), ("Diferencia", 18),
                          ("Variación [%]", 13), ("Cambio", 12)):
        letra = ws.cell(fila_enc, columnas[nombre]).column_letter
        ws.column_dimensions[letra].width = ancho
    primera = ws.cell(fila_enc, primera_col).coordinate
    ultima = ws.cell(ws.max_row, ultima_col).coordinate
    ws.auto_filter.ref = f"{primera}:{ultima}"
    ws.freeze_panes = ws.cell(fila_enc + 1, primera_col + 2).coordinate


# ------------------------------------------------------------
# Escritura (mismo aspecto que el resto de los libros)
# ------------------------------------------------------------

def escribir_hoja_reporte(writer, hoja, lineas, tabla, titulo):
    """
    Una hoja de reporte con el formato de las demas (fila 1 vacia,
    "Cuadro N°" en la 2, nombres en la 3, margen en A): el cuadro 1 son
    las lineas de resumen (una por fila) y a su derecha el cuadro 2, la
    tabla por central.
    """

    from .escritura import _escribir_tabla_con_titulo
    from .lectura import titulo_de_cuadro

    columna = 1
    numero = 1
    # Cuadro 1: las lineas de resumen, o una tabla (hoja Responsables).
    primera = lineas if isinstance(lineas, pd.DataFrame) else (
        pd.DataFrame({"Resumen": lineas}) if lineas else None
    )
    if primera is not None and len(primera):
        _, columna = _escribir_tabla_con_titulo(
            writer, hoja, primera, titulo_de_cuadro(1, titulo), fila_inicio=1,
        )
        numero = 2
    _escribir_tabla_con_titulo(
        writer, hoja, tabla, titulo_de_cuadro(numero, "Por central"),
        fila_inicio=1, columna_inicio=columna,
    )


FILA_ENCABEZADO_REPORTE = 3
FILA_TITULOS_REPORTE = 2
_TEXTOS_LARGOS = (
    "Resumen", "Observación", "Qué pasó", "Principal responsable", "Desglose en $",
    "Qué se movió",
)


def dar_formato_reporte(ws):
    """
    Despues de formatear_hoja: los textos largos (resumen, observacion,
    explicacion) anchos y con ajuste de linea, la observacion con algo
    que decir en amarillo, y cada numero con su formato (plata sin
    decimales, % con uno, FD con dos).
    """

    from openpyxl.styles import Alignment, PatternFill

    amarillo = PatternFill(fill_type="solid", start_color="FFF2CC", end_color="FFF2CC")
    fila_enc = FILA_ENCABEZADO_REPORTE
    for celda_nombre in ws[fila_enc]:
        if celda_nombre.value is None:
            continue
        nombre = str(celda_nombre.value)
        letra = celda_nombre.column_letter
        celdas = [ws.cell(f, celda_nombre.column) for f in range(fila_enc + 1, ws.max_row + 1)]
        if nombre in _TEXTOS_LARGOS:
            ws.column_dimensions[letra].width = {
                "Resumen": 60, "Qué pasó": 70, "Principal responsable": 55,
                "Desglose en $": 80, "Qué se movió": 60,
            }.get(nombre, 90)
            for celda in celdas:
                celda.alignment = Alignment(
                    wrap_text=True, vertical="top", horizontal="left"
                )
                if nombre == "Observación" and celda.value:
                    celda.fill = amarillo
            continue
        if nombre == "Aporte [CLP]":
            from openpyxl.styles import Font

            for celda in celdas:
                if isinstance(celda.value, (int, float)):
                    celda.font = Font(bold=True, color="375623" if celda.value >= 0 else "9C0006")
        if "[%]" in nombre:
            formato = '0.0" %"'
        elif "[CLP]" in nombre:
            formato = "#,##0"
        elif "FD" in nombre:
            formato = "0.00"
        else:
            formato = "#,##0"
        for celda in celdas:
            if isinstance(celda.value, (int, float)) and not isinstance(celda.value, bool):
                celda.number_format = formato
    # La tabla de responsables por central, con filtro (para mirar una
    # central o un responsable).
    nombres_enc = {str(c.value): c.column for c in ws[fila_enc] if c.value is not None}
    if "Lado" in nombres_enc and "Qué se movió" in nombres_enc:
        desde = ws.cell(fila_enc, nombres_enc["Lado"] - 1).coordinate
        hasta = ws.cell(ws.max_row, nombres_enc["Qué se movió"]).coordinate
        ws.auto_filter.ref = f"{desde}:{hasta}"
    # Todo arriba: con un texto largo en la fila, los numeros y el nombre
    # de la central quedan a la altura de su primera linea.
    for fila in ws.iter_rows(min_row=fila_enc + 1):
        for celda in fila:
            if celda.value is not None:
                actual = celda.alignment
                celda.alignment = Alignment(
                    horizontal=actual.horizontal, vertical="top",
                    wrap_text=actual.wrap_text,
                )


def escribir_reporte(destino, hojas):
    """Un libro de reporte. hojas: [(nombre, lineas, tabla, titulo)]."""

    from .formato import formatear_libro

    destino = Path(destino)
    from .escritura import escribir_entero

    with pd.ExcelWriter(destino, engine="openpyxl") as writer:
        escribir_entero(writer)
        for nombre, lineas, tabla, titulo in hojas:
            escribir_hoja_reporte(writer, nombre, lineas, tabla, titulo)
        nombres = [nombre for nombre, *_ in hojas]
        formatear_libro(
            writer.book,
            {n: FILA_ENCABEZADO_REPORTE for n in nombres},
            {n: FILA_TITULOS_REPORTE for n in nombres},
        )
        for nombre in nombres:
            dar_formato_reporte(writer.sheets[nombre])
            formato_cambios(writer.sheets[nombre])
    return destino


# ------------------------------------------------------------
# Reporte de la corrida (boton "Generar reporte")
# ------------------------------------------------------------

def generar_reporte_corrida(carpeta_base, aamm, registrar=print, progreso=None):
    """
    Hoja "Reporte" de Log/Control_corrida.xlsx: lo que recibe cada central
    de esta corrida y, si no recibe (o recibe por un solo lado), por que.
    """

    from .escritura import escribir_control
    from .rutas import resolver_rutas, validar_aamm

    rutas = resolver_rutas(carpeta_base, validar_aamm(aamm))
    if not rutas["salida"].is_file():
        raise ErrorEntrada(
            f"No esta {rutas['salida'].name}: primero calcula la planilla."
        )
    resumen = resumen_por_central(rutas["salida"], registrar)
    if progreso:
        progreso(70)
    if resumen.empty:
        raise ErrorEntrada(
            f"{rutas['salida'].name} no tiene centrales calculadas todavia: "
            f"calcula E Costos y RE545 primero."
        )
    lineas, tabla, ficha = reporte_corrida(resumen)
    escribir_control(rutas["control"], df_reporte=(lineas, tabla, ficha))
    for linea in lineas:
        registrar(f"  {linea}")
    con_observacion = tabla[tabla["Observación"] != ""]
    for _, fila in con_observacion.iterrows():
        registrar(f"  {fila[C_CENTRAL]}: {fila['Observación']}")
    registrar(f"Reporte en {rutas['control']} (hojas 'Reporte' y 'Reporte ficha').")
    return rutas["control"]
