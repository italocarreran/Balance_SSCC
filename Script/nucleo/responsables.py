# -*- coding: utf-8 -*-
"""
Quien es el responsable, en pesos, de que una central reciba distinto entre
dos planillas 11_PAGOS_BESS.

Pedido del usuario (2026-11-10): "necesito que se entienda claramente el
principal responsable del aumento o bajada de ingresos por central [...] a
mi me preguntan 'porque en sept eran 1000 palitos y ahora son 1800?'".

Como se arma el monto (y por eso esto es exacto):

    ingreso del lado = suma, por ciclo (E Costos) o ventana (RE545), de
                       max(0, Total C1 - Total C2)
                     = Total C1 - Total C2 + PISO

donde PISO es lo que se recorta en los ciclos/ventanas negativos (el monto
nunca baja de 0). Y C1, C2 se escriben como productos de factores, todos
razones entre sumas que ya estan en la planilla (asi el producto da el
total exacto):

  E Costos, descarga (la carga igual, con su energia):
    C1 = E_total * w * s * F * p1      C2 = E_total * w * s * p2
      E_total  energia medida de la central (las dos hojas de calculo)
      w        parte que va a E Costos (la que no tuvo oferta completa)
      s        parte de esa energia en horas con Adj SSCC
      F        factor FD y prorrata (energia ajustada / energia adjudicada)
      p1       precio de la curva monotona (C1 / energia ajustada)
      p2       CMg de las horas adjudicadas (C2 / energia adjudicada)

  RE545:
    C1 = A * w * phi * a * p1          C2 = A * w * phi * b * p2
      A        lo adjudicado en Subastas (MW por hora -> kWh)
      w        parte de lo adjudicado que cae en RE545 (oferta completa)
      phi      factor FD y FMA (reserva con FD y FMA / reserva)
      a, b     energia asignada (C1) y disponible (C2) por kWh de reserva
      p1, p2   precio de la curva en C1 y en C2

El cambio de cada producto se reparte entre sus factores con Shapley (el
promedio de su efecto en todos los ordenes posibles): la suma de los
efectos da EXACTO la diferencia, sin residuo.

Componentes (pedido del usuario 2026-11-20): cada lado se arma como
  cambio del monto = cambio de C1 - cambio de C2 + reparto
y cada parte se reparte entre sus factores: "Componente 1" da exacto el
cambio de Total C1, "Componente 2" el de Total C2 (con signo contrario) y el
"Reparto" lo que el max(0, C1 - C2) por ciclo/ventana recorta. Si el CMg de
la central no cambio, el precio promedio va junto con la energia que lo
mueve (no sale como "CMg").

Reemplazos: si una planilla trae reemplazos se recalcula sin ellos
(reporte.resumen_sin_reemplazos) y "Reemplazos" es la diferencia, tambien
partida en C1, C2 y reparto.
"""

import itertools
import math

import pandas as pd

from .utiles import normalizar

# --- Las sumas por central que necesita el modelo (columnas internas del
# resumen; no se muestran). E Costos (descarga "d" y carga "c") y RE545.
X_EC_E_TOT = ("x:ec_e_tot_d", "x:ec_e_tot_c")    # energia medida total
X_EC_E = ("x:ec_e_d", "x:ec_e_c")                # energia en E Costos
X_EC_L = ("x:ec_l_d", "x:ec_l_c")                # energia en horas Adj SSCC
X_EC_Q = ("x:ec_q_d", "x:ec_q_c")                # energia ajustada FD-prorrata
X_EC_C1 = ("x:ec_c1_d", "x:ec_c1_c")             # valorizacion C1
X_EC_C2 = ("x:ec_c2_d", "x:ec_c2_c")             # valorizacion C2
X_EC_I = "x:ec_ingreso"

# La energia ajustada solo con la prorrata (FD = 1): para "cuanto dejo de
# recibir por FD" en el reporte de la corrida.
X_EC_QP = ("x:ec_qp_d", "x:ec_qp_c")
X_RE_A = "x:re_adjudicado"                       # A, en kWh
X_RE_R0 = "x:re_reserva"                         # reserva sin FD ni FMA
X_RE_R = "x:re_reserva_fd_fma"
X_RE_RFMA = "x:re_reserva_fma"                   # con FMA, sin FD (FD = 1)
X_RE_EA = "x:re_e_asignada"
X_RE_EB = "x:re_e_disponible"
X_RE_C1 = "x:re_c1"
X_RE_C2 = "x:re_c2"
X_RE_I = "x:re_ingreso"

# Por ciclo (E Costos) o ventana (RE545): las sumas de C1 y C2 de cada uno,
# para aplicar el max(0, C1 - C2) grupo por grupo.
X_EC_GRUPOS = "x:ec_grupos"    # [(C1d, C1c, C2d, C2c), ...]
X_RE_GRUPOS = "x:re_grupos"    # [(C1, C2, con reserva), ...]

AGREGADOS = (
    *X_EC_E_TOT, *X_EC_E, *X_EC_L, *X_EC_Q, *X_EC_C1, *X_EC_C2, X_EC_I,
    *X_EC_QP, X_RE_A, X_RE_R0, X_RE_R, X_RE_RFMA, X_RE_EA, X_RE_EB, X_RE_C1, X_RE_C2, X_RE_I,
    X_EC_GRUPOS, X_RE_GRUPOS,
)

# --- Los responsables, como se muestran (en este orden si empatan).
ENERGIA = "Energía medida"
OFERTAS = "Ofertas completas"
HORAS = "Horas adjudicadas"
REEMPLAZOS = "Reemplazos (7_Reasignaciones)"
FD_PRORRATA = "FD y prorrata"
ADJUDICACION = "Adjudicación en Subastas"
FD_FMA = "FD y FMA"
USO = "Tope de energía disponible (Edisp)"
CMG = "CMg"
REPARTO = "Reparto entre ciclos y ventanas"
NUEVA = "Central nueva"
YA_NO = "Ya no está"
RESPONSABLES = (
    CMG, ADJUDICACION, OFERTAS, FD_PRORRATA, FD_FMA, HORAS, REEMPLAZOS, ENERGIA,
    USO, REPARTO, NUEVA, YA_NO,
)

# --- De que parte del monto sale cada aporte (monto = max(0, C1 - C2) por
# ciclo o ventana).
LADOS = ("E Costos", "RE545")
C1 = "Componente 1"
C2 = "Componente 2"
AMBAS = "C1 y C2"
COMP_REPARTO = "Reparto (C1 − C2 por ciclo/ventana)"
SIN_COMPONENTE = "—"
COMPONENTES = (C1, C2, AMBAS, COMP_REPARTO, SIN_COMPONENTE)


def _num(serie):
    return pd.to_numeric(serie, errors="coerce")


def _col(df, nombre):
    if df is None or nombre not in df:
        return None
    columna = df[nombre]
    if isinstance(columna, pd.DataFrame):
        columna = columna.iloc[:, 0]
    return pd.to_numeric(columna, errors="coerce").reset_index(drop=True)


def _sumar(df, valores):
    """{central normalizada: suma} de la serie 'valores' (alineada a df)."""

    central = df["Configuracion"]
    if isinstance(central, pd.DataFrame):
        central = central.iloc[:, 0]
    tabla = pd.DataFrame({
        "c": central.reset_index(drop=True).map(normalizar),
        "v": valores.fillna(0.0),
    })
    tabla = tabla[tabla["c"] != ""]
    return tabla.groupby("c")["v"].sum().to_dict()


def agregados(ecostos, re545, subastas):
    """
    {columna de AGREGADOS: {central normalizada: suma}} desde las dos hojas
    de calculo (ya con completar_formulas) y Subastas (ya con el nombre
    Balance_BESS). Lo que falta queda fuera (sin dato).

    """

    datos = {}
    descarga_re = carga_re = None
    if re545 is not None and "Configuracion" in re545:
        descarga_re = _col(re545, "Descarga kWh")
        carga_re = _col(re545, "Carga kWh")

    if ecostos is not None and "Configuracion" in ecostos:
        l = _col(ecostos, "Adj SSCC")
        cmg = _col(ecostos, "CMg")
        for i, (energia, ajustada, c1) in enumerate((
            ("Descarga kWh", "Descarga ajustada FD-prorrata kWh", "Valorización Descarga [$]"),
            ("Carga kWh", "Carga ajustada FD-prorrata kWh", "Valorización Carga [$]"),
        )):
            e = _col(ecostos, energia)
            if e is None:
                continue
            datos[X_EC_E[i]] = _sumar(ecostos, e)
            if l is not None:
                datos[X_EC_L[i]] = _sumar(ecostos, l * e)
                if cmg is not None:
                    datos[X_EC_C2[i]] = _sumar(ecostos, l * e * cmg)
            if _col(ecostos, ajustada) is not None:
                datos[X_EC_Q[i]] = _sumar(ecostos, _col(ecostos, ajustada))
            prorratas = [
                _col(ecostos, n) for n in
                ("Prorrata CPF(-)", "Prorrata CSF(-)", "Prorrata CPF(+)", "Prorrata CSF(+)")
            ]
            if all(x is not None for x in prorratas):
                datos[X_EC_QP[i]] = _sumar(ecostos, e * sum(x.fillna(0.0) for x in prorratas))
            if _col(ecostos, c1) is not None:
                datos[X_EC_C1[i]] = _sumar(ecostos, _col(ecostos, c1))
            # Energia medida total: lo de E Costos y lo de RE545.
            total = dict(datos[X_EC_E[i]])
            otra = descarga_re if i == 0 else carga_re
            if otra is not None:
                for c, v in _sumar(re545, otra).items():
                    total[c] = total.get(c, 0.0) + v
            datos[X_EC_E_TOT[i]] = total
        if _col(ecostos, "Monto a compensar") is not None:
            datos[X_EC_I] = _sumar(ecostos, _col(ecostos, "Monto a compensar"))

    if re545 is not None and "Configuracion" in re545:
        reservas = [
            _col(re545, n) for n in
            ("Reserva CPF(-)", "Reserva CSF(-)", "Reserva CPF(+)", "Reserva CSF(+)")
        ]
        if all(r is not None for r in reservas):
            # Misma cuenta que la Reserva con FD y FMA, sin FD ni FMA.
            datos[X_RE_R0] = _sumar(re545, sum(r.fillna(0.0) for r in reservas) / 4 * 1000)
            fmas = [
                _col(re545, n) for n in
                ("FMA CPF(-)", "FMA CSF(-)", "FMA CPF(+)", "FMA CSF(+)")
            ]
            if all(f is not None for f in fmas):
                datos[X_RE_RFMA] = _sumar(re545, sum(
                    r.fillna(0.0) * f.fillna(0.0) for r, f in zip(reservas, fmas)
                ) / 4 * 1000)
        for nombre, clave in (
            ("Reserva con FD y FMA", X_RE_R), ("Energía asignada", X_RE_EA),
            ("Asignacion Edisponible+SSCC ultima hora", X_RE_EB),
            ("Total C1", X_RE_C1), ("Total C2", X_RE_C2), ("Monto a compensar", X_RE_I),
        ):
            if _col(re545, nombre) is not None:
                datos[clave] = _sumar(re545, _col(re545, nombre))

    if subastas is not None and "Configuración" in subastas and "Energía SSCC" in subastas:
        central = subastas["Configuración"]
        if isinstance(central, pd.DataFrame):
            central = central.iloc[:, 0]
        energia = pd.to_numeric(subastas["Energía SSCC"], errors="coerce").fillna(0.0)
        tabla = pd.DataFrame({
            "c": central.reset_index(drop=True).map(normalizar),
            # MW de una hora -> kWh (la reserva de un cuarto es MW / 4 * 1000).
            "v": energia.reset_index(drop=True) * 1000,
        })
        datos[X_RE_A] = tabla[tabla["c"] != ""].groupby("c")["v"].sum().to_dict()

    datos.update(_grupos(ecostos, re545))
    return datos


def _grupos(ecostos, re545):
    """{X_EC_GRUPOS / X_RE_GRUPOS: {central: [tuplas por grupo]}}."""

    datos = {}
    if ecostos is not None and "Configuracion" in ecostos and "Ciclo de Carga del mes" in ecostos:
        l, cmg = _col(ecostos, "Adj SSCC"), _col(ecostos, "CMg")
        partes = {
            "c1d": _col(ecostos, "Valorización Descarga [$]"),
            "c1c": _col(ecostos, "Valorización Carga [$]"),
        }
        for nombre, energia in (("c2d", "Descarga kWh"), ("c2c", "Carga kWh")):
            e = _col(ecostos, energia)
            partes[nombre] = None if e is None or l is None or cmg is None else l * e * cmg
        if all(v is not None for v in partes.values()):
            datos[X_EC_GRUPOS] = _por_grupo(
                ecostos, "Ciclo de Carga del mes",
                {k: v.fillna(0.0) for k, v in partes.items()}, ("c1d", "c1c", "c2d", "c2c"),
            )
    if re545 is not None and "Configuracion" in re545 and "Ventana de valorizacion" in re545:
        c1, c2, au = _col(re545, "Total C1"), _col(re545, "Total C2"), _col(re545, "Reserva con FD y FMA")
        if c1 is not None and c2 is not None and au is not None:
            datos[X_RE_GRUPOS] = _por_grupo(
                re545, "Ventana de valorizacion",
                {"c1": c1.fillna(0.0), "c2": c2.fillna(0.0), "au": au.fillna(0.0)},
                ("c1", "c2", "au"),
            )
            # La ventana paga solo si tiene reserva (SUMIFS(AU) > 0).
            datos[X_RE_GRUPOS] = {
                c: [(a, b, r > 0) for a, b, r in grupos]
                for c, grupos in datos[X_RE_GRUPOS].items()
            }
    return datos


def _por_grupo(df, columna_grupo, partes, orden):
    central = df["Configuracion"]
    if isinstance(central, pd.DataFrame):
        central = central.iloc[:, 0]
    grupo = df[columna_grupo]
    if isinstance(grupo, pd.DataFrame):
        grupo = grupo.iloc[:, 0]
    tabla = pd.DataFrame({
        "c": central.reset_index(drop=True).map(normalizar),
        "g": grupo.reset_index(drop=True),
        **partes,
    })
    tabla = tabla[(tabla["c"] != "") & tabla["g"].notna()]
    sumas = tabla.groupby(["c", "g"])[list(orden)].sum()
    resultado = {}
    for (c, _g), fila in sumas.iterrows():
        resultado.setdefault(c, []).append(tuple(float(fila[k]) for k in orden))
    return resultado


# ------------------------------------------------------------
# Shapley de un producto
# ------------------------------------------------------------

def _razon(numerador, denominador):
    if not _hay(numerador) or not _hay(denominador) or denominador == 0:
        return math.nan
    return numerador / denominador


def _hay(valor):
    return valor is not None and not (isinstance(valor, float) and math.isnan(valor))


def shapley_producto(viejos, nuevos):
    """
    El cambio de x1*x2*...*xk (de 'viejos' a 'nuevos') repartido entre los
    k factores. La suma da exacto producto(nuevos) - producto(viejos).
    """

    k = len(viejos)
    efectos = [0.0] * k
    for i in range(k):
        otros = [j for j in range(k) if j != i]
        for n in range(k):
            peso = math.factorial(n) * math.factorial(k - n - 1) / math.factorial(k)
            for subconjunto in itertools.combinations(otros, n):
                con = [nuevos[j] if (j in subconjunto or j == i) else viejos[j] for j in range(k)]
                sin = [nuevos[j] if j in subconjunto else viejos[j] for j in range(k)]
                efectos[i] += peso * (math.prod(con) - math.prod(sin))
    return efectos


def _repartir(producto_viejo, producto_nuevo, factores_viejos, factores_nuevos, etiquetas):
    """
    {etiqueta: efecto} del cambio de un producto. Si algun factor no existe
    en uno de los dos meses (0/0), el cambio entero va al primer factor que
    falta: lo que "aparece" o "desaparece".
    """

    viejo = producto_viejo if _hay(producto_viejo) else 0.0
    nuevo = producto_nuevo if _hay(producto_nuevo) else 0.0
    efectos = {}
    if viejo == nuevo:
        return efectos
    faltan = [
        i for i in range(len(etiquetas))
        if not _hay(factores_viejos[i]) or not _hay(factores_nuevos[i])
    ]
    if faltan:
        # Lo que aparece o desaparece: el factor que paso de 0 a algo (o al
        # reves) es el que lo explica (sept: FD 0 -> C1 = 0 y su precio
        # 0/0); si ninguno, el primero que falta.
        a_cero = [
            i for i in range(len(etiquetas))
            if _hay(factores_viejos[i]) and _hay(factores_nuevos[i])
            and (factores_viejos[i] == 0) != (factores_nuevos[i] == 0)
        ]
        efectos[etiquetas[(a_cero or faltan)[0]]] = nuevo - viejo
        return efectos
    for etiqueta, efecto in zip(etiquetas, shapley_producto(factores_viejos, factores_nuevos)):
        efectos[etiqueta] = efectos.get(etiqueta, 0.0) + efecto
    # Lo que no cuadre por redondeo, al primero (no deberia haber nada).
    resto = (nuevo - viejo) - sum(efectos.values())
    if abs(resto) > 1e-6 * max(1.0, abs(nuevo - viejo)):
        efectos[etiquetas[0]] = efectos.get(etiquetas[0], 0.0) + resto
    return efectos


def _factores_ec(fila, i):
    """(C1, [E_tot, w, s, F, p1]), (C2, [E_tot, w, s, p2]) de E Costos."""

    g = lambda clave: fila.get(clave, math.nan)  # noqa: E731
    e_tot, e, l, q = g(X_EC_E_TOT[i]), g(X_EC_E[i]), g(X_EC_L[i]), g(X_EC_Q[i])
    c1, c2 = g(X_EC_C1[i]), g(X_EC_C2[i])
    w, s, f = _razon(e, e_tot), _razon(l, e), _razon(q, l)
    return (c1, [e_tot, w, s, f, _razon(c1, q)]), (c2, [e_tot, w, s, _razon(c2, l)])


def _factores_re(fila):
    """(C1, [A, w, phi, a, p1]), (C2, [A, w, phi, b, p2]) de RE545."""

    g = lambda clave: fila.get(clave, math.nan)  # noqa: E731
    adj, r0, r = g(X_RE_A), g(X_RE_R0), g(X_RE_R)
    ea, eb, c1, c2 = g(X_RE_EA), g(X_RE_EB), g(X_RE_C1), g(X_RE_C2)
    w, phi = _razon(r0, adj), _razon(r, r0)
    return (
        (c1, [adj, w, phi, _razon(ea, r), _razon(c1, ea)]),
        (c2, [adj, w, phi, _razon(eb, r), _razon(c2, eb)]),
    )


def _theta_ec(fila):
    """Los factores de E Costos: [E, w, s, F, p1, p2] de descarga y de carga."""

    return [
        [f for f in (lambda c1, c2: [*c1[1], c2[1][3]])(*_factores_ec(fila, i))]
        for i in range(2)
    ]


def _theta_re(fila):
    """Los factores de RE545: [A, w, phi, a, b, p1, p2]."""

    (_, uno), (_, dos) = _factores_re(fila)
    return [uno[0], uno[1], uno[2], uno[3], dos[3], uno[4], dos[4]]


# Que factores entran en cada parte (indices en theta).
_EC_PARTES = ((0, (0, 1, 2, 3, 4)), (0, (0, 1, 2, 5)))      # C1 y C2 (por signo)
_RE_C1, _RE_C2 = (0, 1, 2, 3, 5), (0, 1, 2, 4, 6)


def _escala(theta, theta_base, indices):
    """Producto de theta/theta_base en esos indices; None si no se puede."""

    escala = 1.0
    for i in indices:
        nuevo, base = theta[i], theta_base[i]
        if _hay(nuevo) and _hay(base) and nuevo == base:
            continue
        if not _hay(nuevo) or not _hay(base) or base == 0:
            return None
        escala *= nuevo / base
    return escala


def _ingreso_ec(theta, theta_base, grupos):
    """Lo que recibe por E Costos con los factores theta, sobre esos ciclos."""

    escalas = []
    for signo in range(2):
        for indices in (_EC_PARTES[0][1], _EC_PARTES[1][1]):
            e = _escala(theta[signo], theta_base[signo], indices)
            if e is None:
                return None
            escalas.append(e)
    e1d, e2d, e1c, e2c = escalas
    return sum(
        max(0.0, c1d * e1d + c1c * e1c - c2d * e2d - c2c * e2c)
        for c1d, c1c, c2d, c2c in grupos
    )


def _ingreso_re(theta, theta_base, grupos):
    e1 = _escala(theta, theta_base, _RE_C1)
    e2 = _escala(theta, theta_base, _RE_C2)
    if e1 is None or e2 is None:
        return None
    return sum(max(0.0, c1 * e1 - c2 * e2) for c1, c2, paga in grupos if paga)


def _completar_theta(theta_a, theta_b):
    """
    Un factor que no existe en un mes (0/0: por ejemplo el precio de C1
    cuando C1 era 0 porque el FD CPF era 0) toma el valor del otro mes: no
    se le atribuye un efecto que no tuvo, y el cambio queda en el factor
    que de verdad se movio (el que paso de 0 a algo).
    """

    def completar(a, b):
        if isinstance(a, list):
            pares = [completar(x, y) for x, y in zip(a, b)]
            return [x for x, _ in pares], [y for _, y in pares]
        if not _hay(a) and _hay(b):
            return b, b
        if _hay(a) and not _hay(b):
            return a, a
        if not _hay(a) and not _hay(b):
            return 1.0, 1.0
        return a, b

    return completar(theta_a, theta_b)


_ETIQUETAS_EC_C1 = tuple((r_, C1) for r_ in (ENERGIA, OFERTAS, HORAS, FD_PRORRATA, CMG))
_ETIQUETAS_EC_C2 = tuple((r_, C2) for r_ in (ENERGIA, OFERTAS, HORAS, CMG))
_ETIQUETAS_RE_C1 = tuple((r_, C1) for r_ in (ADJUDICACION, OFERTAS, FD_FMA, USO, CMG))
_ETIQUETAS_RE_C2 = tuple((r_, C2) for r_ in (ADJUDICACION, OFERTAS, FD_FMA, USO, CMG))


def _sumar_efectos(total, parte, signo=1.0):
    for etiqueta, efecto in parte.items():
        total[etiqueta] = total.get(etiqueta, 0.0) + signo * efecto


def _cmg_igual(actual, anterior):
    """El CMg de la central no cambio (mismo promedio en las dos planillas)."""

    from . import reporte as r

    a, b = actual.get(r.C_CMG, math.nan), anterior.get(r.C_CMG, math.nan)
    return _hay(a) and _hay(b) and abs(a - b) <= 1e-9 * max(1.0, abs(a), abs(b))


def _juntar(factores, i, j):
    """factores con el i-esimo y el j-esimo multiplicados en uno (en la posicion i)."""

    x, y = factores[i], factores[j]
    junto = x * y if _hay(x) and _hay(y) else math.nan
    return [junto if k == i else f for k, f in enumerate(factores) if k != j]


def _por_componente(lado, actual, anterior):
    """
    {(responsable, componente): efecto}: el cambio de C1 repartido entre
    sus factores (Componente 1) y el de C2, con signo contrario porque resta
    (Componente 2). Componente 1 da exacto el cambio de "Total C1" y
    Componente 2 el de "Total C2" (los de la hoja Cambios por central).

    Si el CMg de la central es el mismo en las dos planillas, el precio
    promedio de la curva (o de las horas adjudicadas) solo cambia porque
    entran otros bloques u otras horas: va junto con la energia que lo mueve
    y no sale como "CMg" (ANDES3 oct: se partia en -$21 MM de Edisp y +$9,5
    MM de "CMg" con el CMg igual).
    """

    igual = _cmg_igual(actual, anterior)
    efectos = {}
    if lado == "E Costos":
        for i in range(2):
            (c1_a, c2_a), (c1_b, c2_b) = _factores_ec(actual, i), _factores_ec(anterior, i)
            et1, et2 = list(_ETIQUETAS_EC_C1), list(_ETIQUETAS_EC_C2)
            f1_a, f1_b, f2_a, f2_b = c1_a[1], c1_b[1], c2_a[1], c2_b[1]
            if igual:
                # C1: FD x prorrata x precio de la curva; C2: horas x su CMg.
                f1_a, f1_b, et1 = _juntar(f1_a, 3, 4), _juntar(f1_b, 3, 4), et1[:4]
                f2_a, f2_b, et2 = _juntar(f2_a, 2, 3), _juntar(f2_b, 2, 3), et2[:3]
            _sumar_efectos(efectos, _repartir(c1_b[0], c1_a[0], f1_b, f1_a, tuple(et1)))
            _sumar_efectos(efectos, _repartir(c2_b[0], c2_a[0], f2_b, f2_a, tuple(et2)), -1.0)
    else:
        (c1_a, c2_a), (c1_b, c2_b) = _factores_re(actual), _factores_re(anterior)
        et1, et2 = list(_ETIQUETAS_RE_C1), list(_ETIQUETAS_RE_C2)
        f1_a, f1_b, f2_a, f2_b = c1_a[1], c1_b[1], c2_a[1], c2_b[1]
        if igual:
            # Energia asignada x su precio = C1 por kWh de reserva (Edisp).
            f1_a, f1_b, et1 = _juntar(f1_a, 3, 4), _juntar(f1_b, 3, 4), et1[:4]
            f2_a, f2_b, et2 = _juntar(f2_a, 3, 4), _juntar(f2_b, 3, 4), et2[:4]
        _sumar_efectos(efectos, _repartir(c1_b[0], c1_a[0], f1_b, f1_a, tuple(et1)))
        _sumar_efectos(efectos, _repartir(c2_b[0], c2_a[0], f2_b, f2_a, tuple(et2)), -1.0)
    return efectos


def _c1_c2(fila, lado):
    """(Total C1, Total C2) del lado, de las sumas de la descomposicion."""

    if lado == "E Costos":
        return (sum(_cero(fila.get(k, math.nan)) for k in X_EC_C1),
                sum(_cero(fila.get(k, math.nan)) for k in X_EC_C2))
    return _cero(fila.get(X_RE_C1, math.nan)), _cero(fila.get(X_RE_C2, math.nan))


def descomponer_detalle(actual, anterior):
    """
    {lado: {(responsable, componente): pesos}} del cambio de lo que recibe
    la central (actual - anterior), lado "E Costos" y "RE545":
      Componente 1 = cambio de C1 por factor, Componente 2 = - cambio de C2
      por factor, y el Reparto lo que el max(0, C1 - C2) por ciclo/ventana
      recorta (o deja de recortar). Por lado, la suma da exacto el cambio de
      su ingreso.
    """

    resultado = {}
    for lado, ingreso in (("E Costos", X_EC_I), ("RE545", X_RE_I)):
        nuevo = _cero(actual.get(ingreso, math.nan)) if actual else 0.0
        viejo = _cero(anterior.get(ingreso, math.nan)) if anterior else 0.0
        c1_a, c2_a = _c1_c2(actual, lado) if actual else (0.0, 0.0)
        c1_b, c2_b = _c1_c2(anterior, lado) if anterior else (0.0, 0.0)
        if abs(nuevo - viejo) < 0.5 and abs(c1_a - c1_b) < 0.5 and abs(c2_a - c2_b) < 0.5:
            continue
        if not anterior:
            resultado[lado] = {(NUEVA, SIN_COMPONENTE): nuevo - viejo}
            continue
        if not actual:
            resultado[lado] = {(YA_NO, SIN_COMPONENTE): nuevo - viejo}
            continue
        efectos = _por_componente(lado, actual, anterior)
        # Cierre exacto: lo que falte es el reparto por ciclo/ventana.
        resto = (nuevo - viejo) - sum(efectos.values())
        if abs(resto) >= 0.5:
            clave = (REPARTO, COMP_REPARTO)
            efectos[clave] = efectos.get(clave, 0.0) + resto
        resultado[lado] = {k: v for k, v in efectos.items() if abs(v) >= 0.5}
    return resultado


def descomponer(actual, anterior):
    """
    {lado: {responsable: pesos}} del cambio de lo que recibe la central
    (actual - anterior): descomponer_detalle sumando los componentes.
    """

    resultado = {}
    for lado, por in descomponer_detalle(actual, anterior).items():
        suma = {}
        for (responsable, _), aporte in por.items():
            suma[responsable] = suma.get(responsable, 0.0) + aporte
        resultado[lado] = {k: v for k, v in suma.items() if abs(v) >= 0.5}
    return resultado


def _cero(valor):
    return valor if _hay(valor) else 0.0


def total_por_responsable(descomposicion):
    """{responsable: pesos} sumando los dos lados."""

    total = {}
    for efectos in descomposicion.values():
        _sumar_efectos(total, efectos)
    return total


# ------------------------------------------------------------
# Lo que se movio en cada responsable (para que el numero se entienda)
# ------------------------------------------------------------

def _texto_reemplazos(lado, actual, anterior):
    """Cuantas horas se reemplazaron y cuanto movio eso, anterior → actual."""

    from . import reporte as r

    partes = []
    for nombre, columna in (("horas reemplazadas", r.C_H_REEMPLAZADA),
                            ("horas como reemplazante", r.C_H_REEMPLAZANTE)):
        viejo, nuevo = _cero(anterior.get(columna, math.nan)), _cero(actual.get(columna, math.nan))
        if viejo or nuevo:
            partes.append(f"{nombre} {r.numero(viejo)} → {r.numero(nuevo)}")
    return "; ".join(partes)


def _texto_movimiento(responsable, lado, actual, anterior, componente=None):
    """
    'factor FD y prorrata 0,36 → 0,81' (anterior → actual), o ''. Con el
    componente, lo de ese componente (el precio de C1 o el de C2).
    """

    from . import reporte as r

    def flecha(viejo, nuevo, decimales=0, unidad="", pct=False):
        if not _hay(viejo) or not _hay(nuevo):
            return ""
        if pct:
            return f"{r.numero(viejo * 100)} % → {r.numero(nuevo * 100)} %"
        return (f"{r.numero(viejo, decimales)} → {r.numero(nuevo, decimales)}"
                + (f" {unidad}" if unidad else ""))

    if responsable in (NUEVA, YA_NO):
        return ""
    # Con el CMg igual, el precio promedio va junto con la energia (ver
    # _por_componente): se dice, para que no se lea como cambio de CMg.
    nota_cmg = " (el CMg no cambió: entran otras horas o bloques)"
    igual = _cmg_igual(actual, anterior) if actual and anterior else False
    if responsable == REEMPLAZOS:
        return _texto_reemplazos(lado, actual, anterior)
    if lado == "E Costos":
        a, b = _theta_ec(actual)[0], _theta_ec(anterior)[0]
        if responsable == ENERGIA:
            return "descarga medida " + flecha(b[0] / 1000, a[0] / 1000, 0, "MWh")
        if responsable == OFERTAS:
            return "energía que va a E Costos " + flecha(b[1], a[1], pct=True)
        if responsable == HORAS:
            texto = "energía en horas adjudicadas " + flecha(b[2], a[2], pct=True)
            if igual and componente == C2:
                texto += "; CMg de esas horas " + flecha(b[5], a[5], 1, "CLP/kWh") + nota_cmg
            return texto
        if responsable == FD_PRORRATA:
            texto = "factor FD × prorrata " + flecha(b[3], a[3], 2)
            if igual:
                texto += ("; precio de la curva C1 " + flecha(b[4], a[4], 1, "CLP/kWh")
                          + nota_cmg)
            for nombre, columna in (("FD CPF", r.C_FD_CPF), ("FD CSF", r.C_FD_CSF)):
                viejo, nuevo = anterior.get(columna), actual.get(columna)
                if _hay(viejo) and _hay(nuevo) and abs(nuevo - viejo) >= 0.05:
                    texto += f" ({nombre} {r.numero(viejo, 2)} → {r.numero(nuevo, 2)})"
            # La mezcla: un servicio con FD bajo que se adjudico mas o menos
            # (sept: CPF adjudicado con FD CPF 0; oct: sin CPF -> el factor sube).
            for servicio, adj, fd in (("CPF", r.C_ADJ_CPF, r.C_FD_CPF),
                                      ("CSF", r.C_ADJ_CSF, r.C_FD_CSF)):
                adj_b, adj_a = anterior.get(adj), actual.get(adj)
                fd_b, fd_a = anterior.get(fd), actual.get(fd)
                if not (_hay(adj_b) and _hay(adj_a)) or adj_a == adj_b:
                    continue
                cambio = abs(adj_a - adj_b) / max(abs(adj_a), abs(adj_b))
                fd_con_mas = fd_b if adj_b > adj_a else fd_a
                if cambio >= 0.1 and _hay(fd_con_mas) and fd_con_mas < 0.5:
                    texto += (f"; {servicio} adjudicado {r.numero(adj_b)} → {r.numero(adj_a)} MW"
                              f" con FD {servicio} {r.numero(fd_con_mas, 2)}")
            return texto
        if responsable == CMG and componente == C1:
            return "precio de la curva C1 " + flecha(b[4], a[4], 1, "CLP/kWh")
        if responsable == CMG:
            return "CMg en horas adjudicadas " + flecha(b[5], a[5], 1, "CLP/kWh")
    else:
        a, b = _theta_re(actual), _theta_re(anterior)
        if responsable == ADJUDICACION:
            return "adjudicado " + flecha(b[0] / 1000, a[0] / 1000, 0, "MWh")
        if responsable == OFERTAS:
            return "adjudicado que cae en RE545 " + flecha(b[1], a[1], pct=True)
        if responsable == FD_FMA:
            return "factor FD × FMA " + flecha(b[2], a[2], 2)
        # Cada responsable mueve un factor de C1 y otro de C2: se muestran
        # los que cambiaron (antes solo el de C1, y una diferencia que
        # venia toda de C2 se leia "0,58 -> 0,58").
        def los_dos(textos):
            if componente in (C1, C2):
                nombre, i, decimales, unidad = textos[0 if componente == C1 else 1]
                texto = f"{nombre} {flecha(b[i], a[i], decimales, unidad)}"
                if igual and responsable == USO:
                    precio = 5 if componente == C1 else 6
                    texto += (f"; su precio {flecha(b[precio], a[precio], 1, 'CLP/kWh')}"
                              + nota_cmg)
                return texto
            cambiaron = [
                f"{nombre} {flecha(b[i], a[i], decimales, unidad)}"
                for nombre, i, decimales, unidad in textos
                if _hay(b[i]) and _hay(a[i]) and round(b[i], 4) != round(a[i], 4)
            ]
            if cambiaron:
                return "; ".join(cambiaron)
            nombre, i, decimales, unidad = textos[0]
            return f"{nombre} {flecha(b[i], a[i], decimales, unidad)}"

        if responsable == USO:
            return los_dos((
                ("energía asignada C1 / reserva con FD y FMA", 3, 2, ""),
                ("energía contada en C2 / reserva con FD y FMA", 4, 2, ""),
            ))
        if responsable == CMG:
            return los_dos((
                ("precio de la curva C1", 5, 1, "CLP/kWh"),
                ("precio de lo contado en C2", 6, 1, "CLP/kWh"),
            ))
    if responsable == REPARTO:
        con, total = (
            (r.C_EC_CICLOS, r.C_EC_CICLOS_MES) if lado == "E Costos"
            else (r.C_RE_VENTANAS, r.C_RE_VENTANAS_MES)
        )
        que = "ciclos" if lado == "E Costos" else "ventanas"
        viejo, nuevo = anterior.get(con), actual.get(con)
        if _hay(viejo) and _hay(nuevo):
            return (f"{que} con monto {r.numero(viejo)} → {r.numero(nuevo)}"
                    + (f" de {r.numero(actual.get(total))}" if _hay(actual.get(total)) else ""))
    return ""


# El peso de cada aporte es sobre TODO lo que se movio (la suma de los aportes
# sin signo), no sobre la diferencia neta: cuando lo que sube y lo que baja
# casi se compensan, la diferencia es chica y los pesos salian en miles de %
# (GABRIELA oct: +$1,2 MM netos con +$44 MM de CMg y -$42 MM de energia
# daba 3.603 %). Asi cada uno queda entre -100 % y 100 % y, sin signo,
# suman 100 %. Pedido del usuario (2026-11-20).
PESO = "Peso sobre lo movido [%]"
# Si la diferencia neta es menos que esto de lo movido, se avisa que se
# compensan.
_COMPENSAN = 0.25


def _peso(aporte, bruto):
    return aporte / bruto * 100 if bruto else math.nan


def _efecto_reemplazos(fila, fila_sin):
    """
    {(lado, componente): pesos} que los reemplazos le suman a la central en
    su planilla: en C1, en C2 (con signo contrario) y lo que mueven en el
    reparto por ciclo/ventana.
    """

    if not fila or not fila_sin:
        return {}
    efectos = {}
    for lado, col in (("E Costos", X_EC_I), ("RE545", X_RE_I)):
        monto = _cero(fila.get(col, math.nan)) - _cero(fila_sin.get(col, math.nan))
        (c1, c2), (c1_s, c2_s) = _c1_c2(fila, lado), _c1_c2(fila_sin, lado)
        efectos[(lado, C1)] = c1 - c1_s
        efectos[(lado, C2)] = -(c2 - c2_s)
        efectos[(lado, COMP_REPARTO)] = monto - (c1 - c1_s) + (c2 - c2_s)
    return efectos


def tablas_responsables(actual, anterior, orden=None, actual_sin=None, anterior_sin=None):
    """
    (del mes, por central, {central: (principal, desglose)}).

    actual_sin / anterior_sin: el resumen de la misma planilla recalculada
    SIN los reemplazos (reporte.resumen_sin_reemplazos), si la planilla los
    trae. Entonces el cambio se arma en dos partes, exacto: lo de siempre
    entre las dos planillas sin reemplazos, y "Reemplazos" = lo que los
    reemplazos suman en la actual menos lo que sumaban en la anterior.
    Sin ellos, los reemplazos se separan solo de las horas adjudicadas / lo
    adjudicado (aproximado: lo que mueven en la prorrata o la reserva queda
    en esos responsables).

    Devuelve:
      - del mes: cada responsable con su aporte al cambio del total del mes
        y su % (la suma da la diferencia del mes);
      - por central: central, lado, responsable, aporte, % del cambio de la
        central y lo que se movio;
      - por central, el texto para la hoja Comparacion.
    """

    from . import reporte as r

    a = actual.set_index(actual[r.C_CENTRAL].map(normalizar))
    b = anterior.set_index(anterior[r.C_CENTRAL].map(normalizar))
    a_sin = None if actual_sin is None else actual_sin.set_index(
        actual_sin[r.C_CENTRAL].map(normalizar))
    b_sin = None if anterior_sin is None else anterior_sin.set_index(
        anterior_sin[r.C_CENTRAL].map(normalizar))

    def fila_de(tabla, clave):
        if tabla is None or clave not in tabla.index:
            return None
        return tabla.loc[clave].to_dict()

    claves = list(dict.fromkeys(
        [normalizar(c) for c in (orden or [])] + sorted(set(a.index) | set(b.index))
    ))
    filas, del_mes, textos = [], {}, {}
    for clave in claves:
        en_a, en_b = clave in a.index, clave in b.index
        if not en_a and not en_b:
            continue
        fila_a = a.loc[clave].to_dict() if en_a else {}
        fila_b = b.loc[clave].to_dict() if en_b else {}
        nombre = fila_a.get(r.C_CENTRAL) or fila_b.get(r.C_CENTRAL)
        delta = _cero(fila_a.get(r.C_TOTAL)) - _cero(fila_b.get(r.C_TOTAL))
        sin_a, sin_b = fila_de(a_sin, clave), fila_de(b_sin, clave)
        detalle = descomponer_detalle(
            sin_a if sin_a is not None and fila_a else fila_a,
            sin_b if sin_b is not None and fila_b else fila_b,
        )
        reemplazos = _efecto_reemplazos(fila_a, sin_a)
        for clave_r, valor in _efecto_reemplazos(fila_b, sin_b).items():
            reemplazos[clave_r] = reemplazos.get(clave_r, 0.0) - valor
        for (lado, comp), valor in reemplazos.items():
            if abs(valor) >= 0.5:
                por = detalle.setdefault(lado, {})
                por[(REEMPLAZOS, comp)] = por.get((REEMPLAZOS, comp), 0.0) + valor
        por_lado = {}
        for lado, por in detalle.items():
            for (resp, comp), aporte in por.items():
                por_lado[(lado, resp)] = por_lado.get((lado, resp), 0.0) + aporte
                del_mes[resp] = del_mes.get(resp, 0.0) + aporte
        lista = sorted(
            ((lado, resp, aporte) for (lado, resp), aporte in por_lado.items()
             if abs(aporte) >= 0.5),
            key=lambda x: -abs(x[2]),
        )
        # Una fila por lado, componente y responsable (la hoja la muestra en
        # arbol: Lado > Componente > Responsable).
        bruto = sum(abs(v) for por in detalle.values() for v in por.values() if abs(v) >= 0.5)
        for lado in LADOS:
            por = detalle.get(lado, {})
            for (resp, comp), aporte in sorted(
                por.items(),
                key=lambda kv: (COMPONENTES.index(kv[0][1]), -abs(kv[1])),
            ):
                filas.append({
                    r.C_CENTRAL: nombre, "Lado": lado, "Componente": comp,
                    "Responsable": resp, "Aporte [CLP]": aporte,
                    PESO: _peso(aporte, bruto),
                    "Qué se movió": _texto_movimiento(resp, lado, fila_a, fila_b, comp),
                })
        textos[nombre] = _textos_central(lista, delta, fila_a, fila_b)
    columnas = [r.C_CENTRAL, "Lado", "Componente", "Responsable", "Aporte [CLP]",
                PESO, "Qué se movió"]
    delta_mes = sum(del_mes.values())
    bruto_mes = sum(abs(v) for v in del_mes.values())
    mes = pd.DataFrame([
        {"Responsable": resp, "Aporte [CLP]": aporte, PESO: _peso(aporte, bruto_mes)}
        for resp, aporte in sorted(del_mes.items(), key=lambda kv: -abs(kv[1]))
    ], columns=["Responsable", "Aporte [CLP]", PESO])
    if len(mes):
        mes.loc[len(mes)] = ["Total (= diferencia del mes)", delta_mes, _peso(delta_mes, bruto_mes)]
    return mes, pd.DataFrame(filas, columns=columnas), textos


# Las hojas en arbol (pedido del usuario 2026-11-20: "que no sea desplegable,
# que sea jerarquico"): cada fila con su nivel (df.attrs["niveles"]), los
# niveles de arriba con el subtotal de lo que tienen abajo.
COLUMNA_ARBOL_MES = "Lado / componente / responsable"
COLUMNA_ARBOL_CENTRAL = "Central / lado / componente / responsable"
TOTAL_MES = "Total (= diferencia del mes)"


def _orden(valor, lista):
    return lista.index(valor) if valor in lista else len(lista)


def _ramas(clave, bruto, texto=False):
    """
    El arbol Lado > Componente > Responsable de 'clave' (filas de
    por_central de una central o del mes): [(nombre, aporte, peso sobre
    lo movido 'bruto', que se movio, nivel 1..3)].
    """

    filas = []

    def peso(aporte):
        return _peso(aporte, bruto)

    for lado in sorted(clave["Lado"].unique(), key=lambda v: _orden(v, LADOS)):
        del_lado = clave[clave["Lado"] == lado]
        filas.append((lado, del_lado["Aporte [CLP]"].sum(), "", 1))
        for comp in sorted(del_lado["Componente"].unique(), key=lambda v: _orden(v, COMPONENTES)):
            del_comp = del_lado[del_lado["Componente"] == comp]
            filas.append((comp, del_comp["Aporte [CLP]"].sum(), "", 2))
            juntos = del_comp.groupby("Responsable", sort=False).agg(
                aporte=("Aporte [CLP]", "sum"),
                texto=("Qué se movió", lambda t: "; ".join(x for x in dict.fromkeys(t) if x)),
            )
            for resp, fila in juntos.sort_values("aporte", key=abs, ascending=False).iterrows():
                filas.append((resp, fila["aporte"], fila["texto"] if texto else "", 3))
    return [(n, a, peso(a), t, nivel) for n, a, t, nivel in filas]


def _a_tabla(filas, columnas, con_texto):
    datos = [
        dict(zip(columnas, (n, a, p, t) if con_texto else (n, a, p)))
        for n, a, p, t, _ in filas
    ]
    tabla = pd.DataFrame(datos, columns=columnas)
    tabla.attrs["niveles"] = [nivel for *_, nivel in filas]
    return tabla


def _con_total(nombre, totales):
    """
    'SAE-X  ·  recibe $138 MM → $139 MM (+$1 MM)': lo que recibe en total
    al lado del nombre (pedido del usuario 2026-11-20: "que se vea el total
    al lado del nombre, no como columna"). totales: (anterior, actual) o None.
    """

    from . import reporte as r

    if not totales:
        return nombre
    antes, ahora = (_cero(v) for v in totales)
    dif = ahora - antes
    return (f"{nombre}  ·  recibe {r.plata_corta(antes)} → {r.plata_corta(ahora)} "
            f"({'+' if dif >= 0 else '-'}{r.plata_corta(abs(dif))})")


def arbol_mes(por_central, totales=None):
    """
    El mes en arbol: Lado > Componente > Responsable, cada nivel con el
    subtotal de abajo, y al final el total (= diferencia del mes).
    totales: {"total" | lado: (anterior, actual)} de lo que recibe, para
    ponerlo al lado del nombre del lado y del total.
    """

    columnas = [COLUMNA_ARBOL_MES, "Aporte [CLP]", PESO]
    if por_central is None or not len(por_central):
        return _a_tabla([], columnas, False)
    delta = por_central["Aporte [CLP]"].sum()
    bruto = por_central["Aporte [CLP]"].abs().sum()
    totales = totales or {}
    filas = [
        (_con_total(n, totales.get(n)) if nivel == 1 else n, a, p, t, nivel - 1)
        for n, a, p, t, nivel in _ramas(por_central, bruto)
    ]
    filas.append((_con_total(TOTAL_MES, totales.get("total")), delta,
                  _peso(delta, bruto), "", -1))
    return _a_tabla(filas, columnas, False)


def arbol_por_central(por_central, totales=None):
    """
    Cada central en arbol: Central (su cambio) > Lado > Componente >
    Responsable (con lo que se movio), cada nivel con su subtotal y el peso
    sobre el cambio de la central. totales: {central: (anterior, actual)}
    de lo que recibe, al lado del nombre.
    """

    columnas = [COLUMNA_ARBOL_CENTRAL, "Aporte [CLP]", PESO, "Qué se movió"]
    if por_central is None or not len(por_central):
        return _a_tabla([], columnas, True)
    filas = []
    for central in dict.fromkeys(por_central["Central"]):
        de_central = por_central[por_central["Central"] == central]
        delta = de_central["Aporte [CLP]"].sum()
        bruto = de_central["Aporte [CLP]"].abs().sum()
        filas.append((_con_total(central, (totales or {}).get(central)), delta,
                      _peso(delta, bruto), "", 0))
        filas.extend(_ramas(de_central, bruto, texto=True))
    return _a_tabla(filas, columnas, True)


# Un responsable que explica menos que esto de lo que se movio en la central
# (y menos de $1 MM) no se nombra en el desglose: va en "otros".
_MINIMO_PCT_DESGLOSE = 5.0
_MINIMO_CLP_DESGLOSE = 1e6


_MAXIMO_DESGLOSE = 5


def _textos_central(lista, delta, actual, anterior):
    """
    (principal responsable, desglose en lineas) para Comparacion. El mismo
    responsable en los dos lados va junto (lo que se movio, el del lado que
    mas aporta); hasta 5 lineas, el resto en "otros".
    """

    from . import reporte as r

    if not lista or abs(delta) < r.CAMBIO_MINIMO_CLP:
        return "", ""

    juntos = {}
    for lado, resp, aporte in lista:
        total, lado_mayor, mayor = juntos.get(resp, (0.0, lado, 0.0))
        if abs(aporte) > abs(mayor):
            lado_mayor, mayor = lado, aporte
        juntos[resp] = (total + aporte, lado_mayor, mayor)
    ordenados = sorted(
        ((resp, total, lado) for resp, (total, lado, _) in juntos.items()),
        key=lambda x: -abs(x[1]),
    )
    bruto = sum(abs(total) for _, total, _ in ordenados)

    def linea(resp, aporte, lado):
        signo = "+" if aporte >= 0 else "-"
        peso = abs(_peso(aporte, bruto))
        movimiento = _texto_movimiento(resp, lado, actual, anterior)
        return (f"{signo}{r.plata_corta(abs(aporte))} ({r.numero(peso)} %) {resp}"
                + (f": {movimiento}" if movimiento else ""))

    mismo_sentido = [x for x in ordenados if (x[1] > 0) == (delta > 0)]
    principal = linea(*(mismo_sentido or ordenados)[0])
    desglose, otros = [], 0.0
    for i, (resp, aporte, lado) in enumerate(ordenados):
        importa = (abs(aporte) >= _MINIMO_CLP_DESGLOSE
                   or abs(_peso(aporte, bruto)) >= _MINIMO_PCT_DESGLOSE)
        if importa and len(desglose) < _MAXIMO_DESGLOSE:
            desglose.append("• " + linea(resp, aporte, lado))
        else:
            otros += aporte
    if abs(otros) >= 0.5:
        desglose.append(f"• {'+' if otros >= 0 else '-'}{r.plata_corta(abs(otros))} otros")
    if bruto and abs(delta) < _COMPENSAN * bruto:
        desglose.append(
            f"(Se compensan: se movieron {r.plata_corta(bruto)} entre lo que sube y "
            f"lo que baja, y la diferencia neta es {'+' if delta >= 0 else '-'}"
            f"{r.plata_corta(abs(delta))}. El % es sobre lo movido.)"
        )
    return principal, "\n".join(desglose)


# ------------------------------------------------------------
# Un solo mes: cuanto dejo de recibir por FD (reporte de la corrida)
# ------------------------------------------------------------

# Bajo este factor (FD x prorrata, FD x FMA) no se estima "con FD = 1": el
# C1 de hoy es casi 0 y escalarlo seria inventar.
_FACTOR_MINIMO = 0.1


def perdida_por_fd(fila):
    """
    {"E Costos": pesos, "RE545": pesos}: cuanto MAS habria recibido la
    central este mes con FD = 1 (desempeno perfecto), con los mismos
    escenarios que la comparacion: C1 (y en RE545 tambien C2, que sale de
    la misma reserva) escalados por el factor sin FD, y max(0, C1 - C2)
    ciclo por ciclo. NaN si no se puede armar (por ejemplo un FD = 0 en
    todo el lado: sin C1 no hay precio de la curva con que valorizar).
    Las horas "No participó" vienen con FD = 1: esto no las ve (van
    aparte, como horas).
    """

    resultado = {}
    grupos = fila.get(X_EC_GRUPOS)
    sin_horas_ec = all(_cero(fila.get(X_EC_L[i], math.nan)) == 0 for i in range(2))
    sin_reserva = _cero(fila.get(X_RE_R0, math.nan)) == 0
    if sin_horas_ec:
        resultado["E Costos"] = 0.0      # sin horas adjudicadas: no aplica
    elif isinstance(grupos, list) and _hay(fila.get(X_EC_I)):
        theta = _theta_ec(fila)
        ideal = [list(t) for t in theta]
        posible = True
        for i in range(2):
            f_ideal = _razon(fila.get(X_EC_QP[i], math.nan), fila.get(X_EC_L[i], math.nan))
            if _hay(f_ideal) and _hay(theta[i][3]) and theta[i][3] >= _FACTOR_MINIMO:
                ideal[i][3] = f_ideal
            elif _hay(f_ideal) and f_ideal != 0 and not (
                _hay(theta[i][3]) and theta[i][3] >= _FACTOR_MINIMO
            ):
                # FD (casi) 0 en todo el lado: no hay C1 con que valorizar,
                # o seria multiplicarlo por mucho (SAE-LILE oct: 0,013 -> x77).
                posible = False
        # Un lado sin carga (o sin descarga) tiene sus factores en 0/0.
        ideal, base = _completar_theta(ideal, theta)
        valor = _ingreso_ec(ideal, base, grupos) if posible else None
        resultado["E Costos"] = (valor - fila[X_EC_I]) if valor is not None else math.nan
    grupos = fila.get(X_RE_GRUPOS)
    if sin_reserva:
        resultado["RE545"] = 0.0         # sin reserva en RE545: no aplica
    elif isinstance(grupos, list) and _hay(fila.get(X_RE_I)):
        theta = _theta_re(fila)
        ideal = list(theta)
        phi_ideal = _razon(fila.get(X_RE_RFMA, math.nan), fila.get(X_RE_R0, math.nan))
        valor = None
        if _hay(phi_ideal) and _hay(theta[2]) and theta[2] >= _FACTOR_MINIMO:
            ideal[2] = phi_ideal
            ideal, base = _completar_theta(ideal, theta)
            valor = _ingreso_re(ideal, base, grupos)
        elif _hay(phi_ideal) and phi_ideal == 0:
            valor = fila[X_RE_I]          # el FMA ya es 0: con FD = 1 igual da 0
        if valor is None and _cero(fila.get(X_RE_EA, math.nan)) == 0:
            # Sin energia asignada (sin oferta completa no hay Edisp): con
            # FD = 1 tampoco recibiria por RE545.
            valor = fila[X_RE_I]
        resultado["RE545"] = (valor - fila[X_RE_I]) if valor is not None else math.nan
    # Algunas horas pueden tener un factor sobre 1: lo que "pierde" no baja
    # de 0 (no es una ganancia por FD, es ruido de esas horas).
    return {k: (max(0.0, v) if _hay(v) else v) for k, v in resultado.items()}
