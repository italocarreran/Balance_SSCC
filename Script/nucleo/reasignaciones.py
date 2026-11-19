# -*- coding: utf-8 -*-
"""
Reemplazos de SSCC (asignaciones directas del Coordinador) en Subastas.

Pedido del usuario (2026-11-19, observacion de SAE-CRCA-PFV-ELENA en
2610): las horas en que una BESS presta SSCC como unidad REEMPLAZANTE
de otra (asignacion directa del Coordinador, archivo
7_Reasignaciones_SSCC_AAMM*.xlsm, reporte RIO) no estaban en la hoja
Subastas -que sale solo de los Access de las subastas adjudicadas-, asi
que quedaban con Adj SSCC = 0 y su energia no entraba en C1/C2 ni en
RE545. Y al reves: las horas en que una BESS adjudicada fue
REEMPLAZADA por otra unidad seguian en Subastas como si las hubiera
prestado.

El 7_ va en la carpeta Subastas/ del mes. Con el:

  - Cuadro N° 1 de Subastas (lo adjudicado) lleva la columna
    "Reemplazado": 1 si esa central, servicio y hora fue reemplazada
    por otra unidad (segun el 7_), si no 0.
  - Cuadro N° 2 de Subastas: los reemplazos que involucran una BESS
    (como reemplazante o como reemplazada), con las mismas columnas del
    Cuadro N° 1 -la Configuración es la unidad REEMPLAZANTE- mas
    "Reemplaza a", "Minutos" y "Rol BESS".

Para el calculo (E Costos y RE545) se usa la Subastas "efectiva"
(subastas_para_calculo): el Cuadro N° 1 sin las filas con Reemplazado =
1 (QUITAR_REEMPLAZADAS) mas las filas del Cuadro N° 2 cuya reemplazante
es una BESS (SUMAR_REEMPLAZANTES). Las hojas no cambian por eso: la
foto de lo que dicen los archivos queda tal cual.

Lectura del 7_: hoja "Consolidado SSCC", una fila por tramo (Fecha,
Inicio Periodo, Fin Periodo, Instrucción, Configuración, Disponibilidad
[MW], Reemplazado 1..8). Cada tramo se reparte en las horas que toca
(periodo 1..24 = la hora que termina, como Subastas: 05:00-06:00 es la
6). El calculo es horario (pedido del usuario 2026-11-19: "no hagais
ponderaciones, da lo mismo si es parcial"): un tramo de 09:07 a 10:00
con 75 MW da 75 MW en la hora 10. Si en la misma hora hay varios tramos
de la misma reemplazante y servicio, los que van al mismo tiempo se
suman (reemplaza a dos unidades a la vez) y los que van uno despues del
otro no (gana el mayor). "Minutos" solo informa cuanto de la hora
cubrio el reemplazo.

Nombres: tal cual. "PFV-ELENA" (el parque fotovoltaico) NO es
"SAE-CRCA-PFV-ELENA" (la bateria): tienen medidores distintos y se
pagan por separado. Para una BESS que el Coordinador nombre distinto
esta la columna opcional "Reasignaciones" del Diccionario.
"""

from pathlib import Path

import pandas as pd

from .cambio_hora import hora_mes_correlativa
from .hojas_entrada import _bloques_de_columnas, _contiene_bess_o_sae_sin_bat
from .lector import leer_excel
from .lectura import (
    ROL_BALANCE_BESS, ROL_REASIGNACIONES, ROL_SUBASTAS, fila_de_encabezado,
    filas_diccionario, leer_tabla_bajo_titulo,
)
from .unidades import sin_unidad, sin_unidades
from .utiles import _texto_seguro, _tiene_valor, normalizar


# Para el calculo: sacar de Subastas lo que fue reemplazado y sumar lo
# que una BESS presto como reemplazante. Con los dos en False el
# calculo queda como antes (las hojas igual muestran la bandera y el
# Cuadro N° 2).
QUITAR_REEMPLAZADAS = True
SUMAR_REEMPLAZANTES = True

# El archivo del Coordinador, en la carpeta Subastas/ del mes.
PATRON_ARCHIVO = "*Reasignaciones_SSCC_{aamm}*.xls*"
HOJA_7 = "Consolidado SSCC"

COLUMNA_REEMPLAZADO = "Reemplazado"
COLUMNA_REEMPLAZA_A = "Reemplaza a"
COLUMNA_MINUTOS = "Minutos"
COLUMNA_ROL_BESS = "Rol BESS"
# Separador de las reemplazadas en "Reemplaza a".
SEPARADOR = ", "

ROL_REEMPLAZANTE = "Reemplazante"
ROL_REEMPLAZADA = "Reemplazada"
ROL_AMBAS = "Reemplazante y reemplazada"

# Columnas del Cuadro N° 2: las del Cuadro N° 1 (como se escriben) y
# las propias del reemplazo.
COLUMNAS_CUADRO_1 = (
    "Concepto", "Fecha", "Año", "Mes", "Dia", "Hora_dia", "Hora_mes",
    "Configuración", "Energía SSCC",
)
COLUMNAS_REEMPLAZOS = COLUMNAS_CUADRO_1 + (
    COLUMNA_REEMPLAZA_A, COLUMNA_MINUTOS, COLUMNA_ROL_BESS,
)

TITULO_CUADRO_1 = "subastas adjudicadas del periodo"
TITULO_CUADRO_2 = "reemplazos con BESS (asignación directa, 7_Reasignaciones)"


# ------------------------------------------------------------
# El archivo
# ------------------------------------------------------------

def buscar_archivo_reasignaciones(carpeta_subastas, aamm):
    """
    El 7_Reasignaciones_SSCC_AAMM*.xls* de la carpeta Subastas/, o None.
    Si hay varios (p. ej. "_pre" y el definitivo) gana el mas nuevo.
    """

    carpeta = Path(carpeta_subastas)
    if not carpeta.is_dir():
        return None
    candidatos = [
        ruta for ruta in carpeta.glob(PATRON_ARCHIVO.format(aamm=aamm))
        if ruta.is_file() and not ruta.name.startswith("~$")
    ]
    if not candidatos:
        return None
    return max(candidatos, key=lambda ruta: ruta.stat().st_mtime)


def _columna(df, *nombres):
    """La primera columna de df cuyo nombre (normalizado) empieza con uno de nombres."""

    for nombre in nombres:
        buscado = normalizar(nombre)
        for col in df.columns:
            if normalizar(col).startswith(buscado):
                return col
    return None


def leer_tramos(ruta):
    """
    Los tramos de la hoja "Consolidado SSCC" del 7_: una fila por tramo
    con Inicio, Fin (datetime), Concepto, Reemplazante, MW, Reemplazados
    (lista) y Comentario. Se saltan las filas sin fecha o sin
    reemplazado.
    """

    try:
        crudo = leer_excel(ruta, sheet_name=HOJA_7)
    except ValueError as error:
        raise ValueError(
            f"{Path(ruta).name} no tiene la hoja '{HOJA_7}'."
        ) from error

    c_fecha = _columna(crudo, "Fecha")
    c_ini = _columna(crudo, "Inicio Periodo")
    c_fin = _columna(crudo, "Fin Periodo")
    c_serv = _columna(crudo, "Instrucción")
    c_conf = _columna(crudo, "Configuración")
    c_unidad = _columna(crudo, "Central-Unidad")
    c_mw = _columna(crudo, "Disponibilidad")
    c_com = _columna(crudo, "Comentario")
    c_reemplazados = [
        col for col in crudo.columns
        if normalizar(col).startswith("reemplazado ")
    ]

    faltan = [
        nombre for nombre, col in (
            ("Inicio Periodo", c_ini), ("Fin Periodo", c_fin),
            ("Instrucción", c_serv), ("Configuración", c_conf),
            ("Disponibilidad", c_mw),
        ) if col is None
    ]
    if faltan or not c_reemplazados:
        raise ValueError(
            f"La hoja '{HOJA_7}' de {Path(ruta).name} no trae "
            f"{faltan or ['Reemplazado 1..8']}."
        )

    inicio = pd.to_datetime(crudo[c_ini], errors="coerce")
    fin = pd.to_datetime(crudo[c_fin], errors="coerce")
    if c_fecha is not None:
        # Si Inicio/Fin vinieran como hora sola, se les pone la fecha.
        fecha = pd.to_datetime(crudo[c_fecha], errors="coerce").dt.normalize()
        inicio = inicio.where(inicio.dt.year > 1900, fecha + (inicio - inicio.dt.normalize()))
        fin = fin.where(fin.dt.year > 1900, fecha + (fin - fin.dt.normalize()))

    tramos = []
    for i in range(len(crudo)):
        ini, fi = inicio.iloc[i], fin.iloc[i]
        if pd.isna(ini) or pd.isna(fi) or ini.year < 2000:
            continue
        # El 7_ cierra el dia en 23:59: es hasta las 24:00.
        if fi.hour == 23 and fi.minute == 59:
            fi = fi.normalize() + pd.Timedelta(days=1)
        if fi <= ini:
            continue
        reemplazados = []
        for col in c_reemplazados:
            nombre = _texto_seguro(crudo[col].iloc[i])
            if nombre and nombre not in reemplazados:
                reemplazados.append(nombre)
        if not reemplazados:
            continue
        reemplazante = _texto_seguro(crudo[c_conf].iloc[i])
        if not reemplazante and c_unidad is not None:
            reemplazante = _texto_seguro(crudo[c_unidad].iloc[i])
        tramos.append(dict(
            Inicio=ini, Fin=fi,
            Concepto=_texto_seguro(crudo[c_serv].iloc[i]).replace(" ", ""),
            Reemplazante=reemplazante,
            MW=pd.to_numeric(crudo[c_mw].iloc[i], errors="coerce"),
            Reemplazados=reemplazados,
            Comentario=_texto_seguro(crudo[c_com].iloc[i]) if c_com else "",
        ))

    return pd.DataFrame(
        tramos,
        columns=["Inicio", "Fin", "Concepto", "Reemplazante", "MW",
                 "Reemplazados", "Comentario"],
    )


def tramos_por_hora(tramos):
    """
    Cada tramo repartido en las horas que toca: Fecha (dia de la hora),
    Hora_dia (periodo 1..24, la hora que termina), Minutos del tramo en
    esa hora y el resto de las columnas del tramo.
    """

    filas = []
    for tramo in tramos.itertuples(index=False):
        hora = tramo.Inicio.floor("h")
        while hora < tramo.Fin:
            siguiente = hora + pd.Timedelta(hours=1)
            desde, hasta = max(hora, tramo.Inicio), min(siguiente, tramo.Fin)
            minutos = (hasta - desde).total_seconds() / 60
            if minutos > 0:
                filas.append(dict(
                    Fecha=hora.normalize(), Hora_dia=hora.hour + 1,
                    Minutos=minutos, Desde=desde, Hasta=hasta,
                    Concepto=tramo.Concepto,
                    Reemplazante=tramo.Reemplazante, MW=tramo.MW,
                    Reemplazados=tramo.Reemplazados,
                    Comentario=tramo.Comentario,
                ))
            hora = siguiente
    return pd.DataFrame(
        filas,
        columns=["Fecha", "Hora_dia", "Minutos", "Desde", "Hasta", "Concepto",
                 "Reemplazante", "MW", "Reemplazados", "Comentario"],
    )


# ------------------------------------------------------------
# Nombres: del 7_ al de Subastas
# ------------------------------------------------------------

class NombresBess:
    """
    El nombre del 7_ como esta en Subastas, y si es una BESS.

    Siempre por el nombre EXACTO (sin distinguir mayusculas ni espacios
    repetidos): "PFV-ELENA" es el parque fotovoltaico y
    "SAE-CRCA-PFV-ELENA" la bateria; no se confunden. Lo unico que
    traduce es la columna opcional "Reasignaciones" del Diccionario, para
    una BESS que el Coordinador nombre distinto en el 7_.
    """

    def __init__(self, diccionario, nombres_subastas, registrar=print):
        self.registrar = registrar
        self._de = {}
        self._bess = set()
        filas = filas_diccionario(diccionario) if diccionario is not None else []
        for fila in filas:
            destino = fila.get(ROL_SUBASTAS) or fila.get(ROL_BALANCE_BESS)
            if not _tiene_valor(destino):
                continue
            self._bess.add(normalizar(destino))
            clave = normalizar(fila.get(ROL_REASIGNACIONES))
            if clave and clave not in self._de:
                self._de[clave] = destino
        self._subastas = {}
        for nombre in nombres_subastas:
            texto = _texto_seguro(nombre)
            if texto:
                self._subastas.setdefault(normalizar(texto), texto)
                self._bess.add(normalizar(texto))

    def resolver(self, nombre):
        clave = normalizar(nombre)
        if clave in self._de:
            return self._de[clave]
        if clave in self._subastas:
            return self._subastas[clave]
        return nombre

    def es_bess(self, nombre_resuelto):
        return (
            normalizar(nombre_resuelto) in self._bess
            or _contiene_bess_o_sae_sin_bat(_texto_seguro(nombre_resuelto))
        )


# ------------------------------------------------------------
# Cuadro N° 2 y la bandera del Cuadro N° 1
# ------------------------------------------------------------

def _mw_de_la_hora(tramos):
    """
    El MW de la hora, sin ponderar por minutos: el mayor MW que hay en
    algun momento de la hora, sumando los tramos que van al mismo tiempo.
    """

    mejor = 0.0
    for momento in {desde for desde, _, _ in tramos}:
        activos = sum(mw for desde, hasta, mw in tramos if desde <= momento < hasta)
        mejor = max(mejor, activos)
    return round(mejor, 6)


def _minutos_cubiertos(tramos):
    """Minutos de la hora con algun tramo (la union, sin contar dos veces)."""

    total, fin_actual = 0.0, None
    for desde, hasta, _ in sorted(tramos):
        if fin_actual is not None and desde < fin_actual:
            desde = fin_actual
        if hasta > desde:
            total += (hasta - desde).total_seconds() / 60
            fin_actual = hasta if fin_actual is None else max(fin_actual, hasta)
    return round(total, 2)


def construir_reemplazos(horas, aamm, nombres, registrar=print):
    """
    El Cuadro N° 2 desde los tramos por hora: una fila por servicio,
    hora y unidad reemplazante, solo si la reemplazante o alguna
    reemplazada es BESS. Los nombres ya traducidos al de Subastas.
    """

    anio, mes = 2000 + int(str(aamm)[:2]), int(str(aamm)[2:])
    horas = horas[
        (horas["Fecha"].dt.year == anio) & (horas["Fecha"].dt.month == mes)
    ]

    filas = {}
    propias = 0
    for fila in horas.itertuples(index=False):
        reemplazante = nombres.resolver(fila.Reemplazante)
        reemplazadas = []
        for nombre in fila.Reemplazados:
            resuelto = nombres.resolver(nombre)
            if normalizar(resuelto) == normalizar(reemplazante):
                propias += 1
                continue
            if resuelto not in reemplazadas:
                reemplazadas.append(resuelto)
        if not reemplazadas:
            continue
        clave = (fila.Concepto, fila.Fecha, fila.Hora_dia, reemplazante)
        actual = filas.setdefault(clave, dict(tramos=[], reemplazadas=[]))
        actual["tramos"].append((
            fila.Desde, fila.Hasta,
            float(fila.MW) if pd.notna(fila.MW) else 0.0,
        ))
        for nombre in reemplazadas:
            if nombre not in actual["reemplazadas"]:
                actual["reemplazadas"].append(nombre)

    if propias:
        registrar(
            f"  [AVISO] Reasignaciones: {propias} tramo-hora(s) donde la "
            f"unidad se reemplaza a si misma; no se cuentan."
        )

    salida = []
    for (concepto, fecha, hora_dia, reemplazante), datos in filas.items():
        bess_reemplazante = nombres.es_bess(reemplazante)
        bess_reemplazada = any(nombres.es_bess(n) for n in datos["reemplazadas"])
        if not (bess_reemplazante or bess_reemplazada):
            continue
        rol = (
            ROL_AMBAS if bess_reemplazante and bess_reemplazada
            else ROL_REEMPLAZANTE if bess_reemplazante else ROL_REEMPLAZADA
        )
        salida.append({
            "Concepto": concepto,
            "Fecha": fecha,
            "Año": fecha.year,
            "Mes": fecha.month,
            "Dia": fecha.day,
            "Hora_dia": hora_dia,
            "Configuración": reemplazante,
            "Energía SSCC": _mw_de_la_hora(datos["tramos"]),
            COLUMNA_REEMPLAZA_A: SEPARADOR.join(datos["reemplazadas"]),
            COLUMNA_MINUTOS: _minutos_cubiertos(datos["tramos"]),
            COLUMNA_ROL_BESS: rol,
        })

    df = pd.DataFrame(salida, columns=[c for c in COLUMNAS_REEMPLAZOS if c != "Hora_mes"])
    if df.empty:
        df["Hora_mes"] = pd.Series(dtype="Int64")
    else:
        df["Hora_mes"] = hora_mes_correlativa(df["Fecha"], df["Hora_dia"])
        df = df.sort_values(
            ["Hora_mes", "Configuración", "Concepto"], kind="mergesort"
        )
    return df[list(COLUMNAS_REEMPLAZOS)].reset_index(drop=True)


def reemplazos_desde_archivo(ruta, aamm, diccionario, df_subastas, registrar=print):
    """El Cuadro N° 2 leyendo el 7_ de 'ruta'."""

    registrar(f"Leyendo {Path(ruta).name} (reemplazos de SSCC)...")
    tramos = leer_tramos(ruta)
    nombres = NombresBess(
        diccionario,
        df_subastas["Configuración"].dropna().unique()
        if df_subastas is not None and "Configuración" in df_subastas else [],
        registrar=registrar,
    )
    df = construir_reemplazos(tramos_por_hora(tramos), aamm, nombres, registrar)
    cuenta = df[COLUMNA_ROL_BESS].value_counts().to_dict() if len(df) else {}
    registrar(
        f"  {len(tramos):,} tramo(s) de reemplazo; con BESS: {len(df):,} "
        f"fila(s) por hora ({', '.join(f'{k}: {v}' for k, v in cuenta.items()) or 'ninguna'})."
    )
    return df


def _clave(configuracion, concepto, hora_mes):
    return (
        normalizar(configuracion),
        _texto_seguro(concepto).replace(" ", "").upper(),
        int(hora_mes) if pd.notna(hora_mes) else None,
    )


def claves_reemplazadas(df_reemplazos):
    """(central, servicio, hora_mes) de cada reemplazada del Cuadro N° 2."""

    claves = set()
    if df_reemplazos is None or df_reemplazos.empty:
        return claves
    for concepto, hora_mes, texto in zip(
        df_reemplazos["Concepto"], df_reemplazos["Hora_mes"],
        df_reemplazos[COLUMNA_REEMPLAZA_A],
    ):
        for nombre in _texto_seguro(texto).split(SEPARADOR):
            if nombre.strip():
                claves.add(_clave(nombre.strip(), concepto, hora_mes))
    return claves


def marcar_reemplazadas(df_subastas, df_reemplazos, registrar=print):
    """
    df_subastas con la columna "Reemplazado" (1/0): 1 si esa central,
    servicio y hora del Cuadro N° 1 fue reemplazada segun el Cuadro N° 2.
    """

    if df_subastas is None:
        return None
    claves = claves_reemplazadas(df_reemplazos)
    marca = [
        int(_clave(c, s, h) in claves)
        for c, s, h in zip(
            df_subastas["Configuración"], df_subastas["Concepto"],
            df_subastas["Hora_mes"],
        )
    ]
    df = df_subastas.copy()
    df[COLUMNA_REEMPLAZADO] = pd.Series(marca, index=df.index, dtype="int64")
    marcadas = int(df[COLUMNA_REEMPLAZADO].sum())
    registrar(
        f"  Subastas: {marcadas:,} fila(s) adjudicada(s) marcada(s) como "
        f"reemplazada(s) (columna '{COLUMNA_REEMPLAZADO}')."
    )
    # Reemplazadas BESS del 7_ que no estan en lo adjudicado: se avisa.
    en_subastas = {
        _clave(c, s, h) for c, s, h in zip(
            df_subastas["Configuración"], df_subastas["Concepto"],
            df_subastas["Hora_mes"],
        )
    }
    bess = {normalizar(c) for c in df_subastas["Configuración"].dropna()}
    sueltas = sorted(
        k for k in claves - en_subastas if k[0] in bess
    )
    if sueltas:
        registrar(
            f"  [AVISO] {len(sueltas)} reemplazo(s) de una BESS en una hora "
            f"en que Subastas no la tiene adjudicada (p. ej. "
            f"{', '.join(f'{k[0]} {k[1]} h{k[2]}' for k in sueltas[:5])})."
        )
    return df


# ------------------------------------------------------------
# Leer la hoja (los dos cuadros)
# ------------------------------------------------------------

def leer_subastas_consolidado(fuente, hoja="Subastas"):
    """
    (cuadro_1, cuadro_2) de la hoja Subastas de la planilla. cuadro_2 es
    None si la hoja no lo trae (hoja de antes o mes sin 7_). Una hoja de
    antes, con una sola tabla, se lee igual que con
    leer_hoja_con_fila_vacia. ValueError si no esta la hoja.
    """

    crudo = leer_excel(fuente, sheet_name=hoja, header=None)
    if crudo.empty:
        return pd.DataFrame(), None
    fila = fila_de_encabezado(crudo)
    desde_encabezado = crudo.iloc[fila:].reset_index(drop=True)
    bloques = [b for b in _bloques_de_columnas(desde_encabezado) if len(b) > 1]
    if not bloques:
        return pd.DataFrame(), None

    def leer_bloque(columnas):
        df = leer_tabla_bajo_titulo(
            fuente, hoja, crudo, columnas[0], len(columnas),
            fila_titulo=fila - 1,
        )
        df = sin_unidades(df, hoja)
        # Un cuadro mas corto que el de al lado vuelve con NaN abajo y
        # sus enteros como float (2026.0): las claves de L y la prorrata
        # los comparan como texto, asi que vuelven a ser enteros.
        for col in ("Año", "Mes", "Dia", "Hora_dia", "Hora_mes"):
            if col in df:
                numeros = pd.to_numeric(df[col], errors="coerce")
                if (numeros.dropna() % 1 == 0).all():
                    df[col] = numeros.astype("Int64")
        return df

    cuadro_1 = leer_bloque(bloques[0])
    cuadro_2 = None
    for columnas in bloques[1:]:
        nombres = {
            sin_unidad(_texto_seguro(desde_encabezado.iloc[0, c]), hoja)
            for c in columnas
        }
        if {"Concepto", "Configuración", COLUMNA_REEMPLAZA_A} <= nombres:
            cuadro_2 = leer_bloque(columnas)
            break
    return cuadro_1, cuadro_2


# ------------------------------------------------------------
# Para el calculo
# ------------------------------------------------------------

def subastas_para_calculo(df_subastas, df_reemplazos, registrar=print):
    """
    La Subastas con que calculan E Costos y RE545: el Cuadro N° 1 sin
    lo reemplazado (QUITAR_REEMPLAZADAS) mas lo que una BESS presto como
    reemplazante (SUMAR_REEMPLAZANTES). Si la reemplazante ya tenia esa
    hora y servicio adjudicados, el MW se suma a esa fila (no se repite
    la fila: la prorrata SSCC cuenta filas). Sin Cuadro N° 2 devuelve
    df_subastas tal cual.
    """

    if df_subastas is None:
        return None
    hay_bandera = COLUMNA_REEMPLAZADO in df_subastas
    hay_reemplazos = df_reemplazos is not None and not df_reemplazos.empty
    if not hay_bandera and not hay_reemplazos:
        return df_subastas

    df = df_subastas
    quitadas = 0
    if QUITAR_REEMPLAZADAS and hay_bandera:
        marca = pd.to_numeric(df[COLUMNA_REEMPLAZADO], errors="coerce").fillna(0) > 0
        quitadas = int(marca.sum())
        df = df[~marca]
    df = df.reset_index(drop=True)

    sumadas = agregadas = 0
    if SUMAR_REEMPLAZANTES and hay_reemplazos:
        propias = df_reemplazos[
            df_reemplazos[COLUMNA_ROL_BESS].isin([ROL_REEMPLAZANTE, ROL_AMBAS])
        ]
        posicion = {}
        for i, clave in enumerate(
            _clave(c, s, h) for c, s, h in zip(
                df["Configuración"], df["Concepto"], df["Hora_mes"]
            )
        ):
            posicion.setdefault(clave, i)
        energia = pd.to_numeric(df["Energía SSCC"], errors="coerce").astype(float)
        nuevas = []
        for fila in propias.to_dict("records"):
            clave = _clave(fila["Configuración"], fila["Concepto"], fila["Hora_mes"])
            if clave in posicion:
                i = posicion[clave]
                energia.iloc[i] = (0.0 if pd.isna(energia.iloc[i]) else energia.iloc[i]) + float(fila["Energía SSCC"] or 0)
                sumadas += 1
            else:
                nuevas.append({c: fila[c] for c in COLUMNAS_CUADRO_1})
                agregadas += 1
        df = df.assign(**{"Energía SSCC": energia.values})
        if nuevas:
            extra = pd.DataFrame(nuevas)
            for col in df.columns:
                if col not in extra:
                    extra[col] = 0 if col == COLUMNA_REEMPLAZADO else pd.NA
            df = pd.concat([df, extra[df.columns]], ignore_index=True)

    registrar(
        f"  Subastas para el calculo: {quitadas:,} fila(s) reemplazada(s) "
        f"fuera, {agregadas:,} hora(s) de BESS reemplazante agregada(s) y "
        f"{sumadas:,} sumada(s) a una ya adjudicada."
    )
    return df
