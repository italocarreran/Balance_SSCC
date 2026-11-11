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
efectos da EXACTO la diferencia, sin residuo. Si un factor no existe en un
mes (por ejemplo la central no tenia energia en E Costos), todo el cambio
de ese producto va al factor que aparece o desaparece.
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
    *X_EC_E_TOT, *X_EC_E, *X_EC_L, *X_EC_Q, *X_EC_C1, *X_EC_C2, X_EC_I, *X_EC_QP,
    X_RE_A, X_RE_R0, X_RE_R, X_RE_RFMA, X_RE_EA, X_RE_EB, X_RE_C1, X_RE_C2, X_RE_I,
    X_EC_GRUPOS, X_RE_GRUPOS,
)

# --- Los responsables, como se muestran (en este orden si empatan).
ENERGIA = "Energía medida"
OFERTAS = "Ofertas completas"
HORAS = "Horas adjudicadas"
FD_PRORRATA = "FD y prorrata"
ADJUDICACION = "Adjudicación en Subastas"
FD_FMA = "FD y FMA"
USO = "Tope de energía disponible (Edisp)"
CMG = "CMg"
REPARTO = "Reparto entre ciclos y ventanas"
NUEVA = "Central nueva"
YA_NO = "Ya no está"
RESPONSABLES = (
    CMG, ADJUDICACION, OFERTAS, FD_PRORRATA, FD_FMA, HORAS, ENERGIA, USO,
    REPARTO, NUEVA, YA_NO,
)


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
        efectos[etiquetas[faltan[0]]] = nuevo - viejo
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


# Que factores entran en cada parte (indices en theta) y que jugador mueve
# cada factor.
_EC_PARTES = ((0, (0, 1, 2, 3, 4)), (0, (0, 1, 2, 5)))      # C1 y C2 (por signo)
_EC_JUGADOR = {0: ENERGIA, 1: OFERTAS, 2: HORAS, 3: FD_PRORRATA, 4: CMG, 5: CMG}
_RE_C1, _RE_C2 = (0, 1, 2, 3, 5), (0, 1, 2, 4, 6)
_RE_JUGADOR = {0: ADJUDICACION, 1: OFERTAS, 2: FD_FMA, 3: USO, 4: USO, 5: CMG, 6: CMG}


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


def _shapley(jugadores, valor):
    """{jugador: efecto} con valor(conjunto de jugadores en "nuevo")."""

    k = len(jugadores)
    cache = {}

    def v(conjunto):
        clave = frozenset(conjunto)
        if clave not in cache:
            cache[clave] = valor(clave)
        return cache[clave]

    efectos = {}
    for jugador in jugadores:
        otros = [j for j in jugadores if j != jugador]
        total = 0.0
        for n in range(k):
            peso = math.factorial(n) * math.factorial(k - n - 1) / math.factorial(k)
            for subconjunto in itertools.combinations(otros, n):
                total += peso * (v(set(subconjunto) | {jugador}) - v(set(subconjunto)))
        efectos[jugador] = total
    return efectos


def _descomponer_lado(theta_a, theta_b, grupos_a, grupos_b, ingreso, jugador_de,
                      mezclar, ingreso_a, ingreso_b):
    """
    Shapley entre los factores (cada uno pasa de anterior a actual) y el
    REPARTO entre ciclos/ventanas (los grupos de un mes o del otro). Cada
    escenario aplica max(0, C1 - C2) grupo por grupo: no hay "piso".
    Exacto: con todo en "actual" da lo que recibe ahora; con todo en
    "anterior", lo de antes.
    """

    jugadores = sorted(set(jugador_de.values()), key=RESPONSABLES.index) + [REPARTO]

    def valor(nuevos):
        if nuevos == frozenset(jugadores):
            return ingreso_a
        if not nuevos:
            return ingreso_b
        theta = mezclar(theta_a, theta_b, nuevos)
        primero, segundo = (
            ((grupos_a, theta_a), (grupos_b, theta_b)) if REPARTO in nuevos
            else ((grupos_b, theta_b), (grupos_a, theta_a))
        )
        for grupos, base in (primero, segundo):
            resultado = ingreso(theta, base, grupos)
            if resultado is not None:
                return resultado
        return None

    # Si algun escenario no se puede armar (un factor que no existe en un
    # mes), ese lado se reparte con el modelo simple (ver _repartir).
    efectos = _shapley(jugadores, lambda c: _o_falla(valor(c)))
    return {k: v for k, v in efectos.items() if abs(v) >= 0.5}


class _NoSePuede(Exception):
    pass


def _o_falla(valor):
    if valor is None:
        raise _NoSePuede
    return valor


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


def _mezclar_ec(theta_a, theta_b, nuevos):
    return [
        [theta_a[s][i] if _EC_JUGADOR[i] in nuevos else theta_b[s][i] for i in range(6)]
        for s in range(2)
    ]


def _mezclar_re(theta_a, theta_b, nuevos):
    return [theta_a[i] if _RE_JUGADOR[i] in nuevos else theta_b[i] for i in range(7)]


_ETIQUETAS_EC_C1 = (ENERGIA, OFERTAS, HORAS, FD_PRORRATA, CMG)
_ETIQUETAS_EC_C2 = (ENERGIA, OFERTAS, HORAS, CMG)
_ETIQUETAS_RE = (ADJUDICACION, OFERTAS, FD_FMA, USO, CMG)


def _sumar_efectos(total, parte, signo=1.0):
    for etiqueta, efecto in parte.items():
        total[etiqueta] = total.get(etiqueta, 0.0) + signo * efecto


def _simple(lado, actual, anterior):
    """
    Respaldo, cuando no se pueden armar los escenarios: el cambio de C1 y
    de C2 por factores (producto) y lo que no se explica (el max por
    grupo) como REPARTO.
    """

    efectos = {}
    if lado == "E Costos":
        for i in range(2):
            (c1_a, c2_a), (c1_b, c2_b) = _factores_ec(actual, i), _factores_ec(anterior, i)
            _sumar_efectos(efectos, _repartir(c1_b[0], c1_a[0], c1_b[1], c1_a[1], _ETIQUETAS_EC_C1))
            _sumar_efectos(efectos, _repartir(c2_b[0], c2_a[0], c2_b[1], c2_a[1], _ETIQUETAS_EC_C2), -1.0)
    else:
        (c1_a, c2_a), (c1_b, c2_b) = _factores_re(actual), _factores_re(anterior)
        _sumar_efectos(efectos, _repartir(c1_b[0], c1_a[0], c1_b[1], c1_a[1], _ETIQUETAS_RE))
        _sumar_efectos(efectos, _repartir(c2_b[0], c2_a[0], c2_b[1], c2_a[1], _ETIQUETAS_RE), -1.0)
    return efectos


def descomponer(actual, anterior):
    """
    {lado: {responsable: pesos}} del cambio de lo que recibe la central
    (actual - anterior), lado "E Costos" y "RE545". Por lado, la suma de
    los responsables da exacto el cambio de su ingreso.
    """

    resultado = {}
    for lado, ingreso in (("E Costos", X_EC_I), ("RE545", X_RE_I)):
        nuevo = _cero(actual.get(ingreso, math.nan)) if actual else 0.0
        viejo = _cero(anterior.get(ingreso, math.nan)) if anterior else 0.0
        if abs(nuevo - viejo) < 0.5:
            continue
        if not anterior:
            resultado[lado] = {NUEVA: nuevo - viejo}
            continue
        if not actual:
            resultado[lado] = {YA_NO: nuevo - viejo}
            continue
        if lado == "E Costos":
            grupos_a, grupos_b = actual.get(X_EC_GRUPOS), anterior.get(X_EC_GRUPOS)
            args = (_theta_ec(actual), _theta_ec(anterior), _ingreso_ec, _EC_JUGADOR, _mezclar_ec)
        else:
            grupos_a, grupos_b = actual.get(X_RE_GRUPOS), anterior.get(X_RE_GRUPOS)
            args = (_theta_re(actual), _theta_re(anterior), _ingreso_re, _RE_JUGADOR, _mezclar_re)
        efectos = None
        if isinstance(grupos_a, list) and isinstance(grupos_b, list):
            theta_a, theta_b, funcion, jugador_de, mezclar = args
            theta_a, theta_b = _completar_theta(theta_a, theta_b)
            try:
                efectos = _descomponer_lado(
                    theta_a, theta_b, grupos_a, grupos_b, funcion, jugador_de,
                    mezclar, nuevo, viejo,
                )
            except _NoSePuede:
                efectos = None
        if efectos is None:
            efectos = _simple(lado, actual, anterior)
        # Cierre exacto: lo que falte va al reparto entre ciclos/ventanas.
        resto = (nuevo - viejo) - sum(efectos.values())
        if abs(resto) >= 0.5:
            efectos[REPARTO] = efectos.get(REPARTO, 0.0) + resto
        resultado[lado] = {k: v for k, v in efectos.items() if abs(v) >= 0.5}
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

def _texto_movimiento(responsable, lado, actual, anterior):
    """'factor FD y prorrata 0,36 → 0,81' (anterior → actual), o ''."""

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
    if lado == "E Costos":
        a, b = _theta_ec(actual)[0], _theta_ec(anterior)[0]
        if responsable == ENERGIA:
            return "descarga medida " + flecha(b[0] / 1000, a[0] / 1000, 0, "MWh")
        if responsable == OFERTAS:
            return "energía que va a E Costos " + flecha(b[1], a[1], pct=True)
        if responsable == HORAS:
            return "energía en horas adjudicadas " + flecha(b[2], a[2], pct=True)
        if responsable == FD_PRORRATA:
            texto = "factor FD × prorrata " + flecha(b[3], a[3], 2)
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
        if responsable == USO:
            return "energía asignada / reserva con FD y FMA " + flecha(b[3], a[3], 2)
        if responsable == CMG:
            return "precio de la curva " + flecha(b[5], a[5], 1, "CLP/kWh")
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


def tablas_responsables(actual, anterior, orden=None):
    """
    (del mes, por central, {central: (principal, desglose)}):
      - del mes: cada responsable con su aporte al cambio del total del mes
        y su % (la suma da la diferencia del mes);
      - por central: central, lado, responsable, aporte, % del cambio de la
        central y lo que se movio;
      - por central, el texto para la hoja Comparacion.
    """

    from . import reporte as r

    a = actual.set_index(actual[r.C_CENTRAL].map(normalizar))
    b = anterior.set_index(anterior[r.C_CENTRAL].map(normalizar))
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
        efectos = descomponer(fila_a, fila_b)
        lista = sorted(
            ((lado, resp, aporte) for lado, por in efectos.items() for resp, aporte in por.items()),
            key=lambda x: -abs(x[2]),
        )
        for lado, resp, aporte in lista:
            del_mes[resp] = del_mes.get(resp, 0.0) + aporte
            filas.append({
                r.C_CENTRAL: nombre, "Lado": lado, "Responsable": resp,
                "Aporte [CLP]": aporte,
                "Peso en el cambio [%]": (aporte / abs(delta) * 100) if delta else math.nan,
                "Qué se movió": _texto_movimiento(resp, lado, fila_a, fila_b),
            })
        textos[nombre] = _textos_central(lista, delta, fila_a, fila_b)
    columnas = [r.C_CENTRAL, "Lado", "Responsable", "Aporte [CLP]",
                "Peso en el cambio [%]", "Qué se movió"]
    delta_mes = sum(del_mes.values())
    mes = pd.DataFrame([
        {"Responsable": resp, "Aporte [CLP]": aporte,
         "Peso en el cambio [%]": (aporte / abs(delta_mes) * 100) if delta_mes else math.nan}
        for resp, aporte in sorted(del_mes.items(), key=lambda kv: -abs(kv[1]))
    ], columns=["Responsable", "Aporte [CLP]", "Peso en el cambio [%]"])
    if len(mes):
        mes.loc[len(mes)] = ["Total (= diferencia del mes)", delta_mes, 100.0 if delta_mes else math.nan]
    return mes, pd.DataFrame(filas, columns=columnas), textos


# Un responsable que explica menos que esto del cambio de la central (y
# menos de $1 MM) no se nombra en el desglose: va en "otros".
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

    def linea(resp, aporte, lado):
        signo = "+" if aporte >= 0 else "-"
        peso = abs(aporte) / abs(delta) * 100
        movimiento = _texto_movimiento(resp, lado, actual, anterior)
        return (f"{signo}{r.plata_corta(abs(aporte))} ({r.numero(peso)} %) {resp}"
                + (f": {movimiento}" if movimiento else ""))

    mismo_sentido = [x for x in ordenados if (x[1] > 0) == (delta > 0)]
    principal = linea(*(mismo_sentido or ordenados)[0])
    desglose, otros = [], 0.0
    for i, (resp, aporte, lado) in enumerate(ordenados):
        importa = (abs(aporte) >= _MINIMO_CLP_DESGLOSE
                   or abs(aporte) / abs(delta) * 100 >= _MINIMO_PCT_DESGLOSE)
        if importa and len(desglose) < _MAXIMO_DESGLOSE:
            desglose.append("• " + linea(resp, aporte, lado))
        else:
            otros += aporte
    if abs(otros) >= 0.5:
        desglose.append(f"• {'+' if otros >= 0 else '-'}{r.plata_corta(abs(otros))} otros")
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
