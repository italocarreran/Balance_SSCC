# -*- coding: utf-8 -*-
"""
Graficos interactivos del Calculo RE545 de una planilla 11_PAGOS_BESS.

Pedido del usuario (2026-11-18): "ya habiendo corrido las hojas de calculo
del bess, que yo seleccione el archivo 11 y que con eso genere graficos para
cada central [...] lo mas interactivo posible, que los graficos los pueda ver
de cerca [...] Partamos con RE545. Que se vean las variables como afectan,
Energia disponible, reserva etc."

Que hace:
  - lee la hoja "Calculo RE545" (las dos tablas: la de bloques y el resumen
    por central + ventana) y "Centrales" (Pmax, capacidad, eficiencia);
  - arma, por central y ventana, las series que dibuja la pagina;
  - escribe UN archivo HTML (Log/Graficos_RE545_AAMM.html, al lado de la
    planilla) que se abre con doble clic. No recalcula nada en Python: usa lo
    que ya esta en la planilla. La pagina SI vuelve a hacer el calculo en el
    navegador (mismas reglas que re545_componentes.py) para el simulador de
    "que pasa si", y avisa si con los valores de la planilla no da lo mismo.

Las filas de cada grupo cumplen dos papeles a la vez (igual que en la hoja):
la fila i es el bloque i EN EL TIEMPO (CMg, descarga, carga, SoC, reserva,
monto) y tambien el puesto i+1 DE LA CURVA de CMg descendente (Curva Cmg
Decendente, inyeccion del puesto, energia asignada C1 y C2). Por eso las
series se guardan en el orden de la hoja, sin reordenar.

Graficos con Plotly (zoom, arrastre, leyenda que prende y apaga series,
descarga PNG). La libreria se embebe en el HTML para que funcione sin
internet: se busca una copia local (ver _plotly_local) y, si no hay, se
intenta bajar una vez y se guarda; si tampoco se puede, la pagina la pide
al CDN al abrirse.
"""

import json
import math
import re
import urllib.request
from pathlib import Path

import pandas as pd

from .lector import abrir_excel, leer_excel
from .lectura import fila_de_encabezado, leer_centrales_planilla
from .unidades import sin_unidades
from .parametros import HOJA_CALCULO_RE545, INICIO_VENTANA
from .utiles import normalizar

PLANTILLA = Path(__file__).with_name("graficos_re545.html")

VERSION_PLOTLY = "2.35.2"
URLS_PLOTLY = (
    f"https://cdnjs.cloudflare.com/ajax/libs/plotly.js/{VERSION_PLOTLY}/plotly.min.js",
    f"https://cdn.jsdelivr.net/npm/plotly.js-dist-min@{VERSION_PLOTLY}/plotly.min.js",
)
NOMBRE_PLOTLY = f"plotly-{VERSION_PLOTLY}.min.js"

# Columnas de la tabla de bloques que usa la pagina -> clave corta.
COLUMNAS_BLOQUES = {
    "Dia": "dia",
    "Hora": "hora",
    "Minuto": "min",
    "CMg": "cmg",
    "Descarga kWh": "desc",
    "Carga kWh": "carga",
    "SoC %": "soc",
    "Reserva con FD y FMA": "au",
    "Curva Cmg Decendente": "bm",
    "inyeccion en el periodo del Cmg Descendente": "bs",
    "Energía asignada": "bn",
    "Total C1": "bo",
    "Energia Total": "bq",
    "Asignacion Edisponible": "bu",
    "SSCC ultima hora": "bv",
    "inyeccion orden cronologico": "bw",
    "Energia Asignada Ultima hora": "bz",
    "Total C2": "cc",
    "Monto a compensar": "ce",
}
# Las reservas por servicio, para ver de que se compone la reserva.
COLUMNAS_SERVICIOS = (
    ("CPF(+)", "Reserva CPF(+)", "FD CPF(+)", "FMA CPF(+)"),
    ("CSF(+)", "Reserva CSF(+)", "FD CSF(+)", "FMA CSF(+)"),
    ("CPF(−)", "Reserva CPF(-)", "FD CPF(-)", "FMA CPF(-)"),
    ("CSF(−)", "Reserva CSF(-)", "FD CSF(-)", "FMA CSF(-)"),
)
# Nombres de la tabla resumen (cuadro final de la hoja), en su orden.
COLUMNAS_RESUMEN = (
    "Central", "Ventana", "Oferta Completa", "EiniT", "EalmT",
    "Total Reservas* FD *FMA", "Edisp_T", "check 1", "check 2",
    "Margen ultima hora", "flag ultima hora",
)


class ErrorGraficos(Exception):
    """Algo de la planilla impide armar los graficos (se muestra tal cual)."""


# ------------------------------------------------------------
# Lectura
# ------------------------------------------------------------

def _num(serie):
    return pd.to_numeric(serie, errors="coerce")


def _leer_hoja(libro):
    """
    (bloques, resumen) de Calculo RE545. Se lee una sola vez con
    header=None y se separan las dos tablas por la posicion de "Central"
    (el resumen es el ultimo cuadro, a la derecha, y repite nombres como
    EiniT, que pandas renombraria).
    """

    if HOJA_CALCULO_RE545 not in libro.sheet_names:
        raise ErrorGraficos(
            f"La planilla no tiene la hoja '{HOJA_CALCULO_RE545}'. "
            "Corre primero el Calculo RE545 en Balance_BESS."
        )
    crudo = leer_excel(libro, sheet_name=HOJA_CALCULO_RE545, header=None)
    if crudo.empty:
        raise ErrorGraficos(f"La hoja '{HOJA_CALCULO_RE545}' esta vacia.")
    fila = fila_de_encabezado(crudo.head(4))
    # Los nombres de la hoja llevan la unidad ("Descarga [kWh]"); se pasan
    # a los nombres de siempre ("Descarga kWh") con sin_unidades y, lo que
    # quede, sin el corchete.
    crudos = [str(v).strip() if pd.notna(v) else "" for v in crudo.iloc[fila]]
    internos = sin_unidades(pd.DataFrame(columns=crudos), HOJA_CALCULO_RE545).columns
    nombres = [
        re.sub(r"\s*\[[^\]]*\]\s*$", "", str(n)).strip() if c else ""
        for c, n in zip(crudos, internos)
    ]
    datos = crudo.iloc[fila + 1:].reset_index(drop=True)

    try:
        corte = nombres.index("Central")
    except ValueError:
        corte = None

    def tabla(desde, hasta):
        columnas = {}
        for posicion in range(desde, hasta):
            nombre = nombres[posicion]
            if nombre and nombre not in columnas:
                columnas[nombre] = datos.iloc[:, posicion]
        return pd.DataFrame(columnas)

    bloques = tabla(0, corte if corte is not None else len(nombres))
    bloques = bloques[bloques.get("Configuracion", pd.Series(dtype=object)).notna()]
    resumen = None
    if corte is not None:
        resumen = tabla(corte, len(nombres))
        resumen = resumen[resumen["Central"].notna()].reset_index(drop=True)
    return bloques.reset_index(drop=True), resumen


def _completar(bloques):
    """
    Las columnas con formula que vengan sin valor guardado (libro escrito
    por Python y no abierto en Excel) se calculan con las mismas reglas
    de la hoja.
    """

    from .re545 import FORMULAS_RE545, NOMBRES_CALCULO_RE545
    from .reporte import completar_formulas

    bloques = completar_formulas(bloques, FORMULAS_RE545, NOMBRES_CALCULO_RE545)
    if "SSCC ultima hora" not in bloques or _num(bloques["SSCC ultima hora"]).isna().all():
        au = _num(bloques["Reserva con FD y FMA"]).fillna(0.0)
        hora = _num(bloques["Hora"])
        bloques["SSCC ultima hora"] = au.where((hora == INICIO_VENTANA - 1) & (au != 0), 0.0)
    return bloques


def _redondo(valor, decimales=4):
    if valor is None:
        return None
    try:
        numero = float(valor)
    except (TypeError, ValueError):
        return None
    if math.isnan(numero) or math.isinf(numero):
        return None
    numero = round(numero, decimales)
    return int(numero) if numero == int(numero) and abs(numero) < 1e15 else numero


def _lista(serie, decimales=4):
    return [_redondo(v, decimales) for v in _num(serie).tolist()]


def armar_datos(ruta_planilla, registrar=print):
    """
    El diccionario que se embebe en la pagina:
      {archivo, periodo, inicio_ventana, centrales: [
          {nombre, pmax, capacidad, eficiencia, ventanas: [
              {v, oferta, eini, ealm, reserva, edisp, margen_uh, flag,
               c1, c2, monto, n, t: {series por bloque}, srv: {...}}
          ]}
      ]}
    """

    ruta_planilla = Path(ruta_planilla)
    if not ruta_planilla.is_file():
        raise ErrorGraficos(f"No esta {ruta_planilla}.")

    registrar(f"Leyendo {ruta_planilla.name} (puede tardar un poco)...")
    with abrir_excel(ruta_planilla) as libro:
        bloques, resumen = _leer_hoja(libro)
        try:
            centrales = leer_centrales_planilla(ruta_planilla, libro=libro)
        except Exception:  # noqa: BLE001 - sin Centrales, sin Pmax/capacidad
            centrales = None

    if bloques.empty:
        raise ErrorGraficos(f"La hoja '{HOJA_CALCULO_RE545}' no tiene filas.")
    faltan = [c for c in ("Configuracion", "Ventana de valorizacion", "Orden", "Periodo")
              if c not in bloques]
    if faltan:
        raise ErrorGraficos(
            f"A la hoja '{HOJA_CALCULO_RE545}' le faltan columnas: {', '.join(faltan)}."
        )
    bloques = _completar(bloques)

    maestro = {}
    if centrales is not None:
        tabla_bess = centrales[0]
        for _, fila in tabla_bess.iterrows():
            nombre = normalizar(fila.get("Nombre activo"))
            if nombre:
                maestro[nombre] = {
                    "pmax": _redondo(fila.get("Pmax (MW)")),
                    "capacidad": _redondo(fila.get("Capacidad (MWh)")),
                    "eficiencia": _redondo(fila.get("Eficiencia")),
                    "propietario": str(fila.get("Propietario") or "").strip() or None,
                }

    por_grupo = {}
    if resumen is not None:
        for _, fila in resumen.iterrows():
            clave = (normalizar(fila["Central"]), _redondo(fila["Ventana"]))
            por_grupo[clave] = fila

    salida = []
    bloques["_central"] = bloques["Configuracion"].map(normalizar)
    bloques["_ventana"] = _num(bloques["Ventana de valorizacion"])
    for central, de_central in bloques.groupby("_central", sort=True):
        nombre = str(de_central["Configuracion"].iloc[0]).strip()
        info = maestro.get(central, {})
        ventanas = []
        for ventana, grupo in de_central.groupby("_ventana", sort=True):
            fila_res = por_grupo.get((central, _redondo(ventana)))
            serie = {
                corta: _lista(grupo[larga]) if larga in grupo else None
                for larga, corta in COLUMNAS_BLOQUES.items()
            }
            servicios = {}
            for etiqueta, reserva, fd, fma in COLUMNAS_SERVICIOS:
                if reserva in grupo:
                    mw = _num(grupo[reserva]).fillna(0.0)
                    if mw.abs().sum() == 0:
                        continue
                    factor = (_num(grupo[fd]).fillna(0.0) if fd in grupo else 1.0) * (
                        _num(grupo[fma]).fillna(0.0) if fma in grupo else 1.0)
                    servicios[etiqueta] = {
                        "mw": _lista(mw),
                        "kwh": _lista(mw * factor / 4 * 1000),
                        "fd": _lista(grupo[fd]) if fd in grupo else None,
                        "fma": _lista(grupo[fma]) if fma in grupo else None,
                    }

            def de_resumen(nombre_col, defecto=None):
                if fila_res is None or nombre_col not in fila_res:
                    return defecto
                valor = _redondo(fila_res[nombre_col])
                return defecto if valor is None else valor

            suma = lambda clave: sum(v for v in (serie.get(clave) or []) if v is not None)  # noqa: E731
            c1, c2 = suma("bo"), suma("cc")
            ventanas.append({
                "v": _redondo(ventana),
                "n": len(grupo),
                "oferta": de_resumen("Oferta Completa", 1),
                "eini": de_resumen("EiniT", _redondo(_num(grupo["EiniT"]).iloc[0])
                                   if "EiniT" in grupo else None),
                "ealm": de_resumen("EalmT", _redondo(_num(grupo["EalmT"]).iloc[0])
                                   if "EalmT" in grupo else None),
                "reserva": de_resumen("Total Reservas* FD *FMA", suma("au")),
                "edisp": de_resumen("Edisp_T"),
                "margen_uh": de_resumen("Margen ultima hora", suma("bv")),
                "flag": de_resumen("flag ultima hora", 0),
                "c1": _redondo(c1, 2),
                "c2": _redondo(c2, 2),
                "monto": _redondo(suma("ce"), 2),
                "t": serie,
                "srv": servicios,
            })
        pmax, capacidad, estimado = _maestro_o_estimado(info, ventanas)
        for datos_v in ventanas:
            _llenar_faltantes(datos_v, pmax)
        salida.append({
            "nombre": nombre,
            "pmax": pmax,
            "capacidad": capacidad,
            "estimado": estimado,
            "eficiencia": info.get("eficiencia"),
            "propietario": info.get("propietario"),
            "ventanas": ventanas,
        })
        registrar(f"  {nombre}: {len(ventanas)} ventana(s)")

    hallado = re.search(r"(\d{4})$", ruta_planilla.stem)
    return {
        "archivo": ruta_planilla.name,
        "periodo": hallado.group(1) if hallado else "",
        "inicio_ventana": INICIO_VENTANA,
        "centrales": salida,
    }


def _maestro_o_estimado(info, ventanas):
    """
    (pmax MW, capacidad MWh, [lo estimado]) de la central. Si Centrales no
    los trae, se estiman de la propia hoja: la caja mas llena de C1 es
    Pmax/4 (Energia Total = inyeccion + asignada) y la energia inicial es
    SoC del primer bloque x capacidad.
    """

    pmax, capacidad, estimado = info.get("pmax"), info.get("capacidad"), []
    if pmax is None:
        llena = max((v for w in ventanas for v in (w["t"].get("bq") or []) if v), default=0)
        if llena:
            pmax, estimado = _redondo(llena * 4 / 1000, 3), estimado + ["Pmax"]
    if capacidad is None:
        for w in ventanas:
            soc = (w["t"].get("soc") or [None])[0]
            if soc and w.get("eini"):
                capacidad = _redondo(w["eini"] / soc / 1000, 3)
                estimado.append("capacidad")
                break
    return pmax, capacidad, estimado


def _llenar_faltantes(ventana, pmax):
    """
    Los valores que falten en las columnas del Componente 1 y 2 (una
    planilla escrita por Python y no guardada en Excel deja la formula de
    prueba de la primera fila sin valor) se completan con las mismas
    reglas de re545_componentes.py. Lo que trae la planilla no se toca.
    Recalcula tambien C1, C2 y el monto de la ventana.
    """

    t = ventana["t"]
    claves = ("bn", "bo", "bq", "bu", "bz", "cc", "ce")
    if pmax is None or not t.get("bs") or not t.get("bm"):
        return
    if all(t.get(k) and None not in t[k] for k in claves):
        return
    n = ventana["n"]
    edisp = ventana.get("edisp") or 0.0
    au = [v or 0.0 for v in t.get("au") or [0.0] * n]
    suma_bv = sum(v or 0.0 for v in t.get("bv") or [])
    caja = pmax * 1000 / 4
    bn, acumulado = [None] * n, 0.0
    for i in range(n):
        if t["bs"][i] is None:
            continue
        bn[i] = max(0.0, min(max(0.0, caja - t["bs"][i]), edisp - acumulado))
        acumulado += bn[i]
    bq = [(t["bs"][i] or 0.0) + (bn[i] or 0.0) for i in range(n)]
    bu, acumulado = [0.0] * n, 0.0
    for i in range(n - 1, -1, -1):
        if t["bs"][i]:
            bu[i] = max(0.0, min(bq[i], edisp - acumulado - suma_bv))
        acumulado += bu[i]
    bw = [(t["desc"][i] or 0.0) + (t["carga"][i] or 0.0) for i in range(n)] \
        if t.get("desc") and t.get("carga") else [0.0] * n
    bz = [(0.0 if bw[i] < 0 else 8.0) * (ventana.get("flag") or 0) for i in range(n)]
    precio = lambda i: t["bm"][i]  # noqa: E731
    bo = [None if bn[i] is None or precio(i) is None else bn[i] * precio(i) for i in range(n)]
    cc = [None if precio(i) is None else (bu[i] + bz[i]) * precio(i) for i in range(n)]
    monto = max(sum(v or 0 for v in bo) - sum(v or 0 for v in cc), 0.0)
    total_au = sum(au)
    ce = [monto * a / total_au if total_au else 0.0 for a in au]
    calculado = {"bn": bn, "bo": bo, "bq": bq, "bu": bu, "bz": bz, "cc": cc, "ce": ce}
    for clave, valores in calculado.items():
        actual = t.get(clave) or [None] * n
        t[clave] = [_redondo(valores[i]) if actual[i] is None else actual[i] for i in range(n)]
    ventana["c1"] = _redondo(sum(v or 0 for v in t["bo"]), 2)
    ventana["c2"] = _redondo(sum(v or 0 for v in t["cc"]), 2)
    ventana["monto"] = _redondo(sum(v or 0 for v in t["ce"]), 2)


# ------------------------------------------------------------
# Plotly embebido
# ------------------------------------------------------------

def _plotly_local(carpetas_extra=()):
    """
    El texto de plotly.min.js si hay una copia local: junto a la
    herramienta (Herramientas/), en __config__ o en las carpetas que se
    pasen. Si no hay, se intenta bajar una vez y guardar en la primera
    carpeta donde se pueda escribir. None si no hay forma.
    """

    raiz = Path(__file__).resolve().parents[2]
    candidatas = [*map(Path, carpetas_extra), raiz / "Herramientas", raiz / "__config__"]
    for carpeta in candidatas:
        for nombre in (NOMBRE_PLOTLY, "plotly.min.js"):
            ruta = carpeta / nombre
            if ruta.is_file() and ruta.stat().st_size > 1_000_000:
                return ruta.read_text(encoding="utf-8")
    for url in URLS_PLOTLY:
        try:
            with urllib.request.urlopen(url, timeout=20) as respuesta:  # noqa: S310
                texto = respuesta.read().decode("utf-8")
        except Exception:  # noqa: BLE001 - sin internet se usa el CDN al abrir
            continue
        if len(texto) < 1_000_000:
            continue
        for carpeta in candidatas:
            try:
                carpeta.mkdir(parents=True, exist_ok=True)
                (carpeta / NOMBRE_PLOTLY).write_text(texto, encoding="utf-8")
                break
            except OSError:
                continue
        return texto
    return None


def generar_html(ruta_planilla, destino=None, registrar=print, plotly_js=None,
                 carpetas_plotly=()):
    """
    Escribe la pagina y devuelve su ruta. destino: por defecto
    <carpeta de la planilla>/Log/Graficos_RE545_<AAMM>.html.
    plotly_js: el texto de plotly.min.js (para pruebas); si es None se
    busca con _plotly_local.
    """

    datos = armar_datos(ruta_planilla, registrar=registrar)
    ruta_planilla = Path(ruta_planilla)
    if destino is None:
        sufijo = f"_{datos['periodo']}" if datos["periodo"] else ""
        destino = ruta_planilla.parent / "Log" / f"Graficos_RE545{sufijo}.html"
    destino = Path(destino)
    destino.parent.mkdir(parents=True, exist_ok=True)

    if plotly_js is None:
        registrar("Buscando la libreria de graficos (Plotly)...")
        plotly_js = _plotly_local(carpetas_plotly)
    if plotly_js:
        etiqueta = "<script>" + plotly_js.replace("</script", "<\\/script") + "</script>"
    else:
        registrar("  [AVISO] Plotly no esta en el equipo ni se pudo bajar: la pagina "
                  "lo pedira a internet al abrirse.")
        etiqueta = f'<script src="{URLS_PLOTLY[0]}"></script>'

    json_datos = json.dumps(datos, ensure_ascii=False, separators=(",", ":"))
    json_datos = json_datos.replace("</", "<\\/")
    html = PLANTILLA.read_text(encoding="utf-8")
    html = html.replace("/*__PLOTLY__*/", etiqueta, 1)
    html = html.replace("/*__DATOS__*/null", json_datos, 1)
    destino.write_text(html, encoding="utf-8")
    total = sum(v["monto"] or 0 for c in datos["centrales"] for v in c["ventanas"])
    registrar(f"Listo: {len(datos['centrales'])} central(es); monto RE545 total "
              + f"{total:,.0f}".replace(",", "."))
    registrar(f"Pagina: {destino}")
    return destino
