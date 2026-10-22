# -*- coding: utf-8 -*-
"""
Genera explicacion/datos/ejemplo.js: el caso de ejemplo de las páginas de
explicación (SAE-DEMO, una ventana reducida a 3 horas = 12 bloques), calculado
con las funciones REALES de Script/nucleo. Si cambia el cálculo, se vuelve a
correr este script y las páginas muestran los números nuevos.

Uso (desde la raíz del repositorio):
    python explicacion/generar_ejemplo.py
"""

import json
import sys
from pathlib import Path

import pandas as pd

RAIZ = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(RAIZ))

from Script.nucleo.diccionarios import _normaliza_cuarto  # noqa: E402
from Script.nucleo.ecostos import (  # noqa: E402
    completar_calculo_e_costos_grupos, construir_calculo_e_costos,
)
from Script.nucleo.re545 import (  # noqa: E402
    completar_calculo_re545, completar_checks_resumen_re545,
    construir_calculo_re545,
)
from Script.nucleo.re545_componentes import calcular_componentes_re545  # noqa: E402
from Script.nucleo.re545_resumen import construir_resumen_ventanas_re545  # noqa: E402
from Script.nucleo.utiles import normalizar  # noqa: E402

SALIDA = Path(__file__).resolve().parent / "datos" / "ejemplo.js"

# ---- el caso ----
CENTRAL, BARRA, UNIDAD = "SAE-DEMO", "B-DEMO", "DEMO_U"
PMAX, CAPACIDAD, EFICIENCIA, UMBRAL_SOC = 10.0, 40.0, 0.85, 0.10
CMG = {1: [22, 20, 18, 21], 2: [55, 60, 58, 62], 3: [120, 135, 128, 115]}
GEN = {1: [-1500] * 4, 2: [500] * 4, 3: [1000] * 4}
SOC = {1: [0.20, 0.23, 0.26, 0.29], 2: [0.32, 0.31, 0.30, 0.29],
       3: [0.28, 0.25, 0.22, 0.19]}
ADJUDICADO = (("CPF(+)", "CPF", "SUBIDA", 3.0), ("CSF(-)", "CSF", "BAJADA", 2.0))
FD, FMA = 0.9, 0.95

silencio = lambda *a, **kw: None  # noqa: E731


def _base():
    filas, dic_cmg = [], {}
    for h in (1, 2, 3):
        prom = sum(CMG[h]) / 4
        for q in range(4):
            hora_mes = 24 + h
            cuarto = (hora_mes - 1) * 4 + q + 1
            filas.append({
                "Mes": 7, "Dia": 2, "Hora": h, "Minutos": q * 15,
                "Hora Mes": hora_mes, "Cuarto de Hora": cuarto,
                "clave": CENTRAL, "Gen_Unidad": GEN[h][q], "SoC": SOC[h][q],
                "Ventana": 1, "Copia_Ventana": 1,
            })
            dic_cmg[BARRA.upper() + "|" + _normaliza_cuarto(cuarto)] = (CMG[h][q], prom)
    return pd.DataFrame(filas), dic_cmg


def _subastas(horas):
    return pd.DataFrame([
        {"Concepto": c, "Control": ctl, "Sub_Baj": t, "Mes": 7, "Dia": 2,
         "Hora_dia": h, "Hora_mes": 24 + h, "Configuración": CENTRAL,
         "Energía SSCC": mw, "FD": FD, "FMA": FMA}
        for h in horas for c, ctl, t, mw in ADJUDICADO
    ])


def _fd(tipo):
    return pd.DataFrame([
        {"id": f"{24 + h}{UNIDAD}", "Unidad": UNIDAD, "Central": CENTRAL,
         f"{tipo}(+)": FD, f"{tipo}(-)": FD}
        for h in (1, 2, 3)
    ])


def calcular_ecostos(horas_subasta=(1, 2, 3)):
    base, dic_cmg = _base()
    k = normalizar(CENTRAL)
    ec = construir_calculo_e_costos(
        base.assign(Ventana_No_Completa=1), {k: BARRA}, dic_cmg, registrar=silencio)
    return completar_calculo_e_costos_grupos(
        ec, _subastas(horas_subasta), {k: PMAX}, UMBRAL_SOC, None,
        _fd("CSF"), _fd("CPF"), registrar=silencio)


def calcular_re545():
    base, dic_cmg = _base()
    k = normalizar(CENTRAL)
    re = construir_calculo_re545(
        base.assign(Ventana_No_Completa=0), {k: BARRA}, dic_cmg, registrar=silencio)
    re = completar_calculo_re545(
        re, _subastas((1, 2, 3)), UMBRAL_SOC, {k: CAPACIDAD}, {k: EFICIENCIA},
        registrar=silencio)
    oferta = pd.DataFrame([{"Central": CENTRAL, "Ventana T": 1, "Oferta": 96, "Completa": 1}])
    res = construir_resumen_ventanas_re545(re, oferta, {k: CAPACIDAD}, registrar=silencio)
    for col, serie in calcular_componentes_re545(re, res, {k: PMAX}, registrar=silencio).items():
        re[col] = serie
    res = completar_checks_resumen_re545(res, re)
    return re, res


def _columnas(df, cols):
    out = {}
    for c in cols:
        out[c] = [None if pd.isna(v) else round(float(v), 6) for v in df[c].tolist()]
    return out


def _prom(valores):
    valores = [v for v in valores if v is not None]
    return sum(valores) / len(valores) if valores else 0.0


def main():
    ec = calcular_ecostos()
    sin1 = calcular_ecostos((2, 3))
    re, res = calcular_re545()

    cols_ec = ["Hora", "Minutos", "Cuarto de Hora", "CMg", "Energia_Positiva",
               "Energia_Negativa", "L", "W", "Y", "AB", "AC", "AD", "AE", "AF",
               "AG", "AH", "AM", "AN", "AP", "AQ", "AS", "AT", "AU", "AV",
               "AW", "AX", "S", "T", "U", "AZ"]
    cols_re = ["Hora", "Minutos", "CMg", "R", "S", "SoC", "U", "V",
               "Energia_Positiva", "Energia_Negativa", "AU", "BI", "BJ", "BM",
               "BS", "BN", "BO", "BU", "CC", "CE"]
    E = _columnas(ec, cols_ec)
    R = _columnas(re, cols_re)
    fila_res = res.iloc[0]

    # ---- valores que citan los textos (placeholders {nombre}) ----
    usados_d = [p for p, e in zip(E["AB"], E["AE"]) if p is not None and e]
    usados_c = [p for p, e in zip(E["AD"], E["AF"]) if p is not None and e]
    horas_sub = 3
    lim = horas_sub * 4
    no_resp_d = sum(e - s for e, s in zip(E["AE"], E["AS"]))
    no_resp_c = sum(s - e for e, s in zip(E["AF"], E["AT"]))
    n = len(E["CMg"])
    v = {
        "pmax_mw": PMAX, "capacidad_mwh": CAPACIDAD, "eficiencia": EFICIENCIA,
        "fd": FD, "fma": FMA, "n_bloques": n,
        "pmax_bloque": PMAX * 1000 / 4,
        "mw_cpf": ADJUDICADO[0][3], "mw_csf": ADJUDICADO[1][3],
        "mw_cpf_fd": ADJUDICADO[0][3] * FD, "mw_csf_fd_fma": ADJUDICADO[1][3] * FD * FMA,
        "mw_reserva": ADJUDICADO[0][3] * FD + ADJUDICADO[1][3] * FD * FMA,
        # E Costos
        "ec_tot_desc": sum(e for e, l in zip(E["Energia_Positiva"], E["L"]) if l == 1),
        "ec_tot_carga": -sum(e for e, l in zip(E["Energia_Negativa"], E["L"]) if l == 1),
        "ec_precio_desc": _prom(usados_d), "ec_precio_carga": _prom(usados_c),
        "ec_n_usados_desc": len(usados_d), "ec_n_usados_carga": len(usados_c),
        "ec_ingreso": sum(E["AU"]), "ec_costo": sum(E["AV"]),
        "ec_no_resp_desc": no_resp_d, "ec_no_resp_carga": no_resp_c,
        "ec_bloques_desc": sum(e for e, l in zip(E["Energia_Positiva"], E["L"]) if l == 1) / (PMAX * 250),
        "ec_bloques_carga": -sum(e for e, l in zip(E["Energia_Negativa"], E["L"]) if l == 1) / (PMAX * 250),
        "ec_horas_subasta": horas_sub, "ec_limite_puestos": lim,
        "ec_precio_descuento_desc": _prom([p for p in E["AB"][:lim] if p is not None]),
        "ec_precio_descuento_carga": _prom([p for p in E["AD"][:lim] if p is not None]),
        "ec_descuento": sum(E["AW"]), "ec_c1": sum(E["AX"]),
        "ec_c2_desc": sum(E["S"]), "ec_c2_carga": sum(E["T"]), "ec_c2": sum(E["U"]),
        "ec_monto": sum(E["AZ"]), "ec_monto_bloque": sum(E["AZ"]) / n,
        "ec_monto_sin1": sum(sin1["AZ"]),
        # RE545
        "re_soc0": R["SoC"][0], "re_eini": float(fila_res["EiniT"]),
        "re_carga": -sum(R["Energia_Negativa"]), "re_ealm": float(fila_res["EalmT"]),
        "re_res_bloque": R["AU"][0], "re_res_total": float(fila_res["Total Reservas* FD *FMA"]),
        "re_edisp": float(fila_res["Edisp_T"]), "re_c1": sum(R["BO"]),
        "re_c2": sum(R["CC"]), "re_monto": sum(R["CE"]),
        "re_monto_bloque": R["CE"][0], "re_capacidad_kwh": CAPACIDAD * 1000,
        "re_eini_mas_ealm": float(fila_res["EiniT"]) + float(fila_res["EalmT"]),
    }
    datos = {"ec": E, "ec_sin_subasta_hora1": _columnas(sin1, cols_ec), "re": R, "valores": v}
    SALIDA.parent.mkdir(parents=True, exist_ok=True)
    SALIDA.write_text(
        "// Generado por explicacion/generar_ejemplo.py. No editar a mano.\n"
        "window.EJEMPLO = " + json.dumps(datos, ensure_ascii=False, separators=(",", ":")) + ";\n",
        encoding="utf-8")
    print(f"Escrito {SALIDA.relative_to(RAIZ)}: monto E Costos {v['ec_monto']:,.0f}, "
          f"monto RE545 {v['re_monto']:,.0f}")


if __name__ == "__main__":
    main()
