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

COLUMNAS_RESUMEN = (
    C_CENTRAL, C_EMPRESA, C_ECOSTOS, C_RE545, C_TOTAL, C_DESCARGA, C_CARGA,
    C_CMG, C_HORAS_ADJ, C_ADJ_CPF, C_ADJ_CSF, C_RESERVA, C_FD_CPF, C_FD_CSF,
    C_OFERTA_DIAS,
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
    return tabla.groupby("c")["v"].agg(como).to_dict()


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
        datos[C_CARGA] = _por_central(ecostos, "Configuracion", "Carga kWh")
        datos[C_CMG] = _por_central(ecostos, "Configuracion", "CMg", "mean")
        adj = _por_central(ecostos, "Configuracion", "Adj SSCC")
        # Adj SSCC va por cuarto de hora: 4 cuartos = 1 hora.
        datos[C_HORAS_ADJ] = {c: v / 4 for c, v in adj.items()}
    elif medidores is not None and _col(medidores, "Gen_Unidad") is not None:
        anotar(medidores, "clave")
        energia = _num(_col(medidores, "Gen_Unidad"))
        clave = _col(medidores, "clave").map(normalizar)
        datos[C_DESCARGA] = energia.clip(lower=0).groupby(clave).sum().to_dict()
        datos[C_CARGA] = energia.clip(upper=0).groupby(clave).sum().to_dict()
    if re545 is not None:
        datos[C_RESERVA] = _por_central(re545, "Configuracion", "Reserva con FD y FMA")

    # --- lo adjudicado en Subastas, con el nombre Balance_BESS ---------------
    if subastas is not None and _col(subastas, "Configuración") is not None:
        subastas = subastas_con_nombre_balance_bess(
            subastas, diccionario, registrar=callado
        )
        anotar(subastas, "Configuración")
        servicio = _col(subastas, "Concepto").map(lambda v: _texto_seguro(v)[:3].upper()) \
            if _col(subastas, "Concepto") is not None else None
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

    centrales_norm = sorted(nombres, key=lambda c: nombres[c].upper())
    filas = []
    for c in centrales_norm:
        fila = {C_CENTRAL: nombres[c], C_EMPRESA: empresa_de.get(c, "")}
        for columna in COLUMNAS_RESUMEN[2:]:
            if columna == C_TOTAL:
                continue
            fila[columna] = datos[columna].get(c, 0.0) if columna in datos else math.nan
        partes = [fila[C_ECOSTOS], fila[C_RE545]]
        fila[C_TOTAL] = (
            math.nan if all(math.isnan(p) for p in partes)
            else sum(0.0 if math.isnan(p) else p for p in partes)
        )
        filas.append(fila)

    return pd.DataFrame(filas, columns=list(COLUMNAS_RESUMEN))


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
        motivo = next(motivos(), None)
        return "Solo recibe por E Costos" + (f": {motivo[0].lower()}{motivo[1:]}" if motivo else ".")
    if _hay(ecostos) and ecostos == 0 and _hay(re545) and re545 > 0:
        return "Solo recibe por RE545."
    return ""


def reporte_corrida(resumen):
    """
    (texto de arriba, tabla por central) para la hoja "Reporte" del
    control: lo que recibe cada central y, si hay algo que decir, por que.
    """

    tabla = resumen[[C_CENTRAL, C_EMPRESA, C_ECOSTOS, C_RE545, C_TOTAL]].copy()
    tabla["Observación"] = [explicar_central(f) for f in resumen.to_dict("records")]

    total = resumen[C_TOTAL].sum(min_count=1)
    sin_ingreso = resumen[resumen[C_TOTAL].fillna(-1) == 0][C_CENTRAL].tolist()
    lineas = []
    if _hay(total):
        lineas.append(f"Total del mes: {plata(total)} entre {len(resumen)} centrales.")
        lineas.append(
            f"E Costos {plata(resumen[C_ECOSTOS].sum())} · "
            f"RE545 {plata(resumen[C_RE545].sum())}."
        )
    if sin_ingreso:
        lineas.append(
            f"{len(sin_ingreso)} central(es) sin ingreso: {', '.join(sin_ingreso)} "
            f"(el motivo va en su fila)."
        )
    elif _hay(total):
        lineas.append("Todas las centrales reciben algo.")
    return lineas, tabla


# ------------------------------------------------------------
# Comparacion entre dos planillas
# ------------------------------------------------------------

# Que mira cada lado para explicar un cambio, y como se dice.
_MOTORES_RE545 = (
    (C_ADJ_CPF, "lo adjudicado en CPF", "MW"),
    (C_ADJ_CSF, "lo adjudicado en CSF", "MW"),
    (C_RESERVA, "la reserva con FD y FMA", "kWh"),
    (C_FD_CPF, "el FD CPF promedio", ""),
    (C_FD_CSF, "el FD CSF promedio", ""),
    (C_OFERTA_DIAS, "los días con oferta completa", "días"),
)
_MOTORES_ECOSTOS = (
    (C_DESCARGA, "la descarga", "kWh"),
    (C_CARGA, "la carga", "kWh"),
    (C_CMG, "el CMg promedio", "CLP/kWh"),
    (C_HORAS_ADJ, "las horas con Adj SSCC", "h"),
)


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
        # La carga es negativa: "mas carga" es un numero mas chico.
        crece = (nuevo < viejo) if columna == C_CARGA else (nuevo > viejo)
        if crece != sube:
            continue
        decimales = 2 if columna in (C_FD_CPF, C_FD_CSF) else 0
        viejo_txt = numero(abs(viejo) if columna == C_CARGA else viejo, decimales)
        nuevo_txt = numero(abs(nuevo) if columna == C_CARGA else nuevo, decimales)
        detalle = f"{viejo_txt} → {nuevo_txt}{(' ' + unidad) if unidad else ''}"
        plural = nombre.startswith(("los ", "las "))
        verbo = ("subieron" if crece else "bajaron") if plural else (
            "subió" if crece else "bajó"
        )
        candidatos.append((abs(pct) if not math.isinf(pct) else 1e9,
                           f"{nombre} {verbo} ({detalle})"))
    candidatos.sort(reverse=True)
    return [texto for _, texto in candidatos[:2]]


def _frase_lado(lado, nuevo, viejo, actual, anterior, motores):
    delta = (nuevo if _hay(nuevo) else 0.0) - (viejo if _hay(viejo) else 0.0)
    if abs(delta) < CAMBIO_MINIMO_CLP:
        return None
    pct = _pct(nuevo, viejo)
    sube = delta > 0
    cabeza = (
        f"{lado} {'sube' if sube else 'baja'} {plata(abs(delta))}"
        + (f" ({_texto_pct(pct)})" if _hay(pct) else "")
    )
    razones = _motores(actual, anterior, motores, sube)
    if razones:
        return f"{cabeza}: {' y '.join(razones)}."
    return f"{cabeza}, sin un cambio grande en sus entradas."


def comparar_resumenes(actual, anterior):
    """
    (lineas de arriba, tabla por central con los dos meses y la
    explicacion). actual/anterior: resumen_por_central() de cada planilla.
    """

    a = actual.set_index(actual[C_CENTRAL].map(normalizar))
    b = anterior.set_index(anterior[C_CENTRAL].map(normalizar))
    claves = sorted(set(a.index) | set(b.index),
                    key=lambda c: (a[C_CENTRAL].get(c) or b[C_CENTRAL].get(c)).upper())

    filas, sin_cambios = [], []
    for clave in claves:
        fila_a = a.loc[clave].to_dict() if clave in a.index else {}
        fila_b = b.loc[clave].to_dict() if clave in b.index else {}
        nombre = fila_a.get(C_CENTRAL) or fila_b.get(C_CENTRAL)
        empresa = fila_a.get(C_EMPRESA) or fila_b.get(C_EMPRESA) or ""
        total_a = fila_a.get(C_TOTAL, math.nan)
        total_b = fila_b.get(C_TOTAL, math.nan)

        if not fila_b:
            explicacion = f"Nueva: no estaba en la planilla anterior. Recibe {plata(total_a)}." \
                if _hay(total_a) else "Nueva: no estaba en la planilla anterior."
        elif not fila_a:
            explicacion = (f"Ya no está en esta planilla (antes recibía "
                           f"{plata(total_b)}).") if _hay(total_b) else "Ya no está en esta planilla."
        else:
            frases = [
                _frase_lado("RE545", fila_a.get(C_RE545), fila_b.get(C_RE545),
                            fila_a, fila_b, _MOTORES_RE545),
                _frase_lado("E Costos", fila_a.get(C_ECOSTOS), fila_b.get(C_ECOSTOS),
                            fila_a, fila_b, _MOTORES_ECOSTOS),
            ]
            frases = [f for f in frases if f]
            if not frases:
                sin_cambios.append(nombre)
            explicacion = " ".join(frases)

        delta = (total_a if _hay(total_a) else 0.0) - (total_b if _hay(total_b) else 0.0)
        filas.append({
            C_CENTRAL: nombre, C_EMPRESA: empresa,
            "Total anterior [CLP]": total_b, "Total actual [CLP]": total_a,
            "Diferencia [CLP]": delta, "Diferencia [%]": _pct(total_a, total_b),
            "Explicación": explicacion,
        })

    tabla = pd.DataFrame(filas)
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
            f"Total del mes: {plata(total_b)} → {plata(total_a)} "
            f"({'+' if delta >= 0 else '-'}{plata(abs(delta))}"
            + (f", {_texto_pct(_pct(total_a, total_b))}" if total_b else "") + ")."
        )
        for lado, col in (("RE545", C_RE545), ("E Costos", C_ECOSTOS)):
            va, vb = actual[col].sum(), anterior[col].sum()
            lineas.append(f"{lado}: {plata(vb)} → {plata(va)}.")
    if len(tabla):
        cambio = tabla[tabla["Diferencia [CLP]"].abs() >= CAMBIO_MINIMO_CLP]
        if len(cambio):
            mayor = cambio.iloc[0]
            lineas.append(
                f"El mayor cambio: {mayor[C_CENTRAL]} "
                f"({'+' if mayor['Diferencia [CLP]'] >= 0 else '-'}"
                f"{plata(abs(mayor['Diferencia [CLP]']))})."
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
    detalle = _detalle_lado_a_lado(actual, anterior)
    escribir_reporte(destino, [
        ("Comparacion", lineas, tabla,
         f"{ruta_actual.name} contra {ruta_anterior.name}"),
        ("Detalle", [], detalle, "Las entradas de cada central, mes anterior y actual"),
    ])
    registrar(f"Reporte: {destino}")
    return destino, lineas, tabla


def _periodo(ruta):
    """'11_PAGOS_BESS_2610.xlsx' -> '2610' (o el nombre sin extension)."""

    import re

    hallado = re.search(r"(\d{4})$", Path(ruta).stem)
    return hallado.group(1) if hallado else Path(ruta).stem


def _detalle_lado_a_lado(actual, anterior):
    """Cada entrada en dos columnas (anterior | actual), central por central."""

    a = actual.set_index(actual[C_CENTRAL].map(normalizar))
    b = anterior.set_index(anterior[C_CENTRAL].map(normalizar))
    claves = sorted(set(a.index) | set(b.index))
    filas = []
    for clave in claves:
        fila = {C_CENTRAL: (a[C_CENTRAL].get(clave) or b[C_CENTRAL].get(clave))}
        for columna in COLUMNAS_RESUMEN[2:]:
            fila[f"{columna} anterior"] = b[columna].get(clave, math.nan) if clave in b.index else math.nan
            fila[f"{columna} actual"] = a[columna].get(clave, math.nan) if clave in a.index else math.nan
        filas.append(fila)
    return pd.DataFrame(filas)


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
    if lineas:
        _, columna = _escribir_tabla_con_titulo(
            writer, hoja, pd.DataFrame({"Resumen": lineas}),
            titulo_de_cuadro(1, titulo), fila_inicio=1,
        )
        numero = 2
    _escribir_tabla_con_titulo(
        writer, hoja, tabla, titulo_de_cuadro(numero, "Por central"),
        fila_inicio=1, columna_inicio=columna,
    )


FILA_ENCABEZADO_REPORTE = 3
FILA_TITULOS_REPORTE = 2
_TEXTOS_LARGOS = ("Resumen", "Observación", "Explicación")


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
            ws.column_dimensions[letra].width = 70 if nombre == "Resumen" else 90
            for celda in celdas:
                celda.alignment = Alignment(
                    wrap_text=True, vertical="top", horizontal="left"
                )
                if nombre == "Observación" and celda.value:
                    celda.fill = amarillo
            continue
        if "[%]" in nombre:
            formato = '0.0" %"'
        elif "FD" in nombre:
            formato = "0.00"
        else:
            formato = "#,##0"
        for celda in celdas:
            if isinstance(celda.value, (int, float)) and not isinstance(celda.value, bool):
                celda.number_format = formato
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
    with pd.ExcelWriter(destino, engine="openpyxl") as writer:
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
    lineas, tabla = reporte_corrida(resumen)
    escribir_control(rutas["control"], df_reporte=(lineas, tabla))
    for linea in lineas:
        registrar(f"  {linea}")
    con_observacion = tabla[tabla["Observación"] != ""]
    for _, fila in con_observacion.iterrows():
        registrar(f"  {fila[C_CENTRAL]}: {fila['Observación']}")
    registrar(f"Reporte en {rutas['control']} (hoja 'Reporte').")
    return rutas["control"]
