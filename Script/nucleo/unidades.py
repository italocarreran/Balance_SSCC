# -*- coding: utf-8 -*-
"""
La unidad de medida de cada columna, entre corchetes, en las hojas.

Pedido del usuario: "agregar unidad de medida [...] donde corresponda
entre corchetes []. Si es CLP o USD igual".

Solo cambia lo que se VE: el nombre de la columna en la fila de
encabezados de la hoja. Dentro del programa las columnas siguen con su
nombre de siempre (todo el calculo las busca por ese nombre): la
escritura le pone la unidad al escribir la hoja (con_unidades) y cada
lectura se la saca al leerla (sin_unidades). Un libro de antes, sin
unidades, se lee igual.

De donde sale cada unidad (revisado en el codigo, no supuesto):
  - Gen_Unidad: energia del intervalo en kWh (docs/Estructura_Archivos_Reales.md).
  - CMg: CMg[CLP/KWh] de cmg_15min (Script/Cmg/Extrae_CMG_barras.py);
    la valorizacion kWh x CMg da CLP ("Total [CLP]" de E Costos).
  - Energía SSCC de Subastas: CANTIDAD PONDERADA MW; RE545 la pasa a
    kWh del cuarto con /4*1000 (re545_reservas.py, AU).
  - SoC: fraccion (0 a 1), RE545 hace SoC x Capacidad (MWh) x 1000.
  - Pmax (MW) y Capacidad (MWh) de Resumen BESS.
Los factores (FD, FMA, prorratas, Eficiencia) son adimensionales y
quedan sin corchetes.
"""

from .parametros import (
    HOJA_CALCULO_ECOSTOS, HOJA_CALCULO_RE545, HOJA_CENTRALES, HOJA_CMG,
    HOJA_COMPENSACION_CENTRAL, HOJA_MEDIDORES, HOJA_PRORRATA_RETIROS,
    HOJA_RESUMEN, HOJA_SUBASTAS,
)


KWH = "[kWh]"
CLP = "[CLP]"
CLP_KWH = "[CLP/kWh]"
PU = "[p.u.]"

# Calculo E Costos y Calculo RE545 comparten buena parte de los nombres.
_COMUNES_CALCULO = {
    "Descarga kWh": f"Descarga {KWH}",
    "Carga kWh": f"Carga {KWH}",
    "SoC %": f"SoC {PU}",
    "CMg": f"CMg {CLP_KWH}",
    "Energía SSCC (-) por remunerar": f"Energía SSCC (-) por remunerar {KWH}",
    "Energía SSCC (+) por remunerar": f"Energía SSCC (+) por remunerar {KWH}",
    "Monto a compensar": f"Monto a compensar {CLP}",
}

# {hoja: {nombre interno: nombre visible}}
NOMBRES_CON_UNIDAD = {
    HOJA_MEDIDORES: {
        "Gen_Unidad": f"Gen_Unidad {KWH}",
        # La fecha y hora (inicio) del intervalo de 15 minutos.
        "intervalo": "intervalo [fecha hh:mm]",
        "SoC": f"SoC {PU}",
    },
    HOJA_SUBASTAS: {
        "Energía SSCC": "Energía SSCC [MW]",
    },
    HOJA_CMG: {
        "F_CMg": f"F_CMg {CLP_KWH}",
        "I_CMgProm": f"I_CMgProm {CLP_KWH}",
    },
    HOJA_CENTRALES: {
        "Pmax (MW)": "Pmax [MW]",
        "Capacidad (MWh)": "Capacidad [MWh]",
    },
    HOJA_CALCULO_ECOSTOS: {
        **_COMUNES_CALCULO,
        "Valorizacion Descarga": f"Valorizacion Descarga {CLP}",
        "Valorizacion Carga": f"Valorizacion Carga {CLP}",
        "Curva monotona CMg Descarga":
            f"Curva monotona CMg Descarga {CLP_KWH}",
        "Curva monotona CMg Carga": f"Curva monotona CMg Carga {CLP_KWH}",
        "Energía descargada": f"Energía descargada {KWH}",
        "Energía cargada": f"Energía cargada {KWH}",
        "Descarga ajustada FD-prorrata kWh":
            f"Descarga ajustada FD-prorrata {KWH}",
        "Carga ajustada FD-prorrata kWh": f"Carga ajustada FD-prorrata {KWH}",
        "Energía SSCC (descargada con FD-Prorrata) por remunerar":
            f"Energía SSCC (descargada con FD-Prorrata) por remunerar {KWH}",
        "Energía SSCC (cargada con FD-Prorrata) por remunerar":
            f"Energía SSCC (cargada con FD-Prorrata) por remunerar {KWH}",
        "Asignación Descarga kWh": f"Asignación Descarga {KWH}",
        "Asignación Carga kWh": f"Asignación Carga {KWH}",
        "Valorización Descarga [$]": f"Valorización Descarga {CLP}",
        "Valorización Carga [$]": f"Valorización Carga {CLP}",
        # U (grupo "Componente 2": lo ganado bloque a bloque a su CMg)
        # y BA (AY + AZ: el Componente 1).
        "Total": f"Total C2 {CLP}",
        "Total [CLP]": f"Total C1 {CLP}",
    },
    HOJA_CALCULO_RE545: {
        **_COMUNES_CALCULO,
        "CMg Promedio": f"CMg Promedio {CLP_KWH}",
        "EiniT": f"EiniT {KWH}",
        "EalmT": f"EalmT {KWH}",
        "SUMA Reservas*FMA*FD": f"SUMA Reservas*FMA*FD {KWH}",
        "Curva Cmg Decendente promedio horario":
            f"Curva Cmg Decendente promedio horario {CLP_KWH}",
        "Curva Cmg Decendente": f"Curva Cmg Decendente {CLP_KWH}",
        "Edisp_Asig": f"Edisp_Asig {KWH}",
        "Total  C1_545": f"Total C1_545 {CLP}",
        "Energia Total": f"Energia Total {KWH}",
        "inyeccion en el periodo del Cmg Descendente":
            f"inyeccion en el periodo del Cmg Descendente {KWH}",
        "Energía ya Asignada": f"Energía ya Asignada {KWH}",
        "Asignacion Edisponible": f"Asignacion Edisponible {KWH}",
        "SSCC ultima hora": f"SSCC ultima hora {KWH}",
        "inyeccion orden cronologico": f"inyeccion orden cronologico {KWH}",
        "Energia Ultima hora": f"Energia Ultima hora {KWH}",
        "Energía ya Asignada ultima hora":
            f"Energía ya Asignada ultima hora {KWH}",
        "Energia Asignada Ultima hora": f"Energia Asignada Ultima hora {KWH}",
        "Asignacion Edisponible+SSCC ultima hora":
            f"Asignacion Edisponible+SSCC ultima hora {KWH}",
        "Total C2_545": f"Total C2_545 {CLP}",
        "Total Reservas* FD *FMA": f"Total Reservas* FD *FMA {KWH}",
        "Edisp_T": f"Edisp_T {KWH}",
    },
    HOJA_COMPENSACION_CENTRAL: {
        "Compensación Ecostos [$]": f"Compensación Ecostos {CLP}",
        "Compensación RE545 [$]": f"Compensación RE545 {CLP}",
        "Compensación Total [$]": f"Compensación Total {CLP}",
    },
    HOJA_PRORRATA_RETIROS: {
        "Monto a compensar [$]": f"Monto a compensar {CLP}",
        "Monto a pagar [$]": f"Monto a pagar {CLP}",
        "Total a pagar [$]": f"Total a pagar {CLP}",
    },
    HOJA_RESUMEN: {
        "RECIBE": f"RECIBE {CLP}",
        "PAGA": f"PAGA {CLP}",
        "NETO": f"NETO {CLP}",
    },
}

_NOMBRES_INTERNOS = {
    hoja: {visible: interno for interno, visible in nombres.items()}
    for hoja, nombres in NOMBRES_CON_UNIDAD.items()
}


def con_unidades(ws, hoja, fila_encabezado):
    """Pone la unidad en los nombres de columna de la hoja ya escrita."""

    nombres = NOMBRES_CON_UNIDAD.get(hoja)
    if not nombres or ws.max_row < fila_encabezado:
        return
    for celda in ws[fila_encabezado]:
        visible = nombres.get(celda.value)
        if visible is not None:
            celda.value = visible


def sin_unidad(nombre, hoja):
    """El nombre interno de una columna leida de la hoja."""

    return _NOMBRES_INTERNOS.get(hoja, {}).get(nombre, nombre)


def sin_unidades(df, hoja):
    """El DataFrame leido de la hoja, con los nombres internos."""

    internos = _NOMBRES_INTERNOS.get(hoja)
    if not internos or df is None:
        return df
    return df.rename(columns=lambda nombre: internos.get(nombre, nombre))
