# -*- coding: utf-8 -*-
"""Quien RECIBE la compensacion: por central, por ciclo/ventana y por empresa.

La contraparte de ``prorrata_retiros.py`` (quien PAGA). Las dos hojas
de calculo dejan el monto en la misma columna ("Monto a compensar"),
pero no lo agrupan igual:

- "Calculo E Costos" reparte `AZ` por (central + ``Ciclo de Carga del
  mes``): el total del grupo dividido entre sus filas, repetido en
  todas ellas.
- "Calculo RE545" reparte `CE` por (central + ``Ventana de
  valorizacion``), en proporcion a `AU`.

En los dos casos la suma de las filas del grupo devuelve el total del
grupo, asi que agrupar y sumar es lo correcto para las dos.
"""

import pandas as pd

from .subastas_accdb import construir_mapa_propietario
from .utiles import normalizar


COL_CENTRAL = "Central"
COL_EMPRESA = "Empresa"
COL_CICLO = "Ciclo / Ventana"
COL_COMPENSACION = "Compensación [$]"
COL_COMPENSACION_TOTAL = "Compensación Total [$]"
COL_MONTO = "Monto a compensar"

# De donde sale el ciclo/ventana en cada hoja de calculo. El orden es
# el de preferencia: la primera columna que exista es la que se usa.
CICLO_ECOSTOS = ("Ciclo de Carga del mes", "Ciclo")
CICLO_RE545 = ("Ventana de valorizacion", "Ciclo de Carga del mes")


def _columna_ciclo(df, candidatas):
    for nombre in candidatas:
        if nombre in df.columns:
            return nombre
    return None


def _con_empresa(df, propietarios, registrar):
    """Agrega la columna Empresa (propietario) a partir de la central."""

    empresas = df[COL_CENTRAL].map(
        lambda central: propietarios.get(normalizar(central), "")
    )
    sin_duenio = sorted(set(df.loc[empresas.eq(""), COL_CENTRAL]) - {""})
    if sin_duenio:
        registrar(
            f"  {len(sin_duenio):,} central(es) sin Propietario en "
            f"'Resumen BESS'; quedan a su propio nombre: {sin_duenio[:5]}"
        )
    # Sin propietario, la central se representa a si misma: es mejor
    # que perder la plata en una fila con la empresa vacia.
    return empresas.where(empresas.ne(""), df[COL_CENTRAL])


def _detalle_por_central(df, candidatas_ciclo, propietarios, registrar):
    """Central + ciclo/ventana -> compensacion, para una hoja de calculo."""

    columnas = [COL_CENTRAL, COL_EMPRESA, COL_CICLO, COL_COMPENSACION]
    if df is None or df.empty or COL_MONTO not in df.columns:
        return pd.DataFrame(columns=columnas)
    columna_ciclo = _columna_ciclo(df, candidatas_ciclo)
    detalle = pd.DataFrame({
        COL_CENTRAL: df["Configuracion"].fillna("").astype(str).str.strip(),
        COL_CICLO: (
            df[columna_ciclo] if columna_ciclo
            else pd.Series("", index=df.index)
        ),
        COL_COMPENSACION: pd.to_numeric(
            df[COL_MONTO], errors="coerce"
        ).fillna(0.0),
    })
    if columna_ciclo is None:
        registrar(
            "  La hoja de calculo no trae columna de ciclo/ventana; el "
            "resumen queda solo por central."
        )
    detalle[COL_EMPRESA] = _con_empresa(detalle, propietarios, registrar)
    return (
        detalle.groupby([COL_CENTRAL, COL_EMPRESA, COL_CICLO], as_index=False,
                        dropna=False)[COL_COMPENSACION].sum()
        .sort_values([COL_CENTRAL, COL_CICLO], kind="stable")
        .loc[:, columnas].reset_index(drop=True)
    )


def construir_compensacion_central(df_ecostos, df_re545, resumen_bess,
                                   registrar=None):
    """Los tres cuadros de la hoja COMPENSACION_CENTRAL.

    Devuelve (E Costos por central/ciclo, RE545 por central/ventana,
    total recibido por empresa).
    """

    registrar = registrar or (lambda _mensaje: None)
    propietarios = construir_mapa_propietario(resumen_bess)
    por_ecostos = _detalle_por_central(
        df_ecostos, CICLO_ECOSTOS, propietarios, registrar
    )
    por_re545 = _detalle_por_central(
        df_re545, CICLO_RE545, propietarios, registrar
    )
    total = (
        pd.concat([por_ecostos, por_re545], ignore_index=True)
        .groupby(COL_EMPRESA, as_index=False)[COL_COMPENSACION].sum()
        .rename(columns={COL_COMPENSACION: COL_COMPENSACION_TOTAL})
        .sort_values(COL_EMPRESA, kind="stable").reset_index(drop=True)
    )
    return por_ecostos, por_re545, total


def construir_compensacion_total(df_ecostos, df_re545, resumen_bess,
                                 registrar=None):
    """Solo el total recibido por empresa (lo que consume 'Resumen')."""

    return construir_compensacion_central(
        df_ecostos, df_re545, resumen_bess, registrar=registrar
    )[2]


def construir_resumen(compensacion_total, pagos, registrar=None):
    """Une acreedores y deudores: RECIBE, PAGA y NETO por empresa.

    Los dos lados vienen de fuentes distintas -RECIBE del Propietario
    de 'Resumen BESS', PAGA del Suministrador de la prorrata del CEN-,
    asi que el cruce se hace por el nombre NORMALIZADO (minuscula, sin
    tildes, sin espacios de mas). Cruzando por el texto crudo, una
    empresa que recibe y paga aparecia dos veces, cada una con la
    mitad de la historia y un NETO que no era su neto.
    """

    from .prorrata_retiros import COL_PAGO, COL_SUMINISTRADOR, COL_TOTAL_PAGO

    registrar = registrar or (lambda _mensaje: None)

    recibe = compensacion_total.rename(
        columns={COL_EMPRESA: "NOMBRE", COL_COMPENSACION_TOTAL: "RECIBE"}
    ).loc[:, ["NOMBRE", "RECIBE"]].copy()
    columna_pago = COL_TOTAL_PAGO if COL_TOTAL_PAGO in pagos.columns else COL_PAGO
    paga = pagos.rename(
        columns={COL_SUMINISTRADOR: "NOMBRE", columna_pago: "PAGA"}
    ).loc[:, ["NOMBRE", "PAGA"]].copy()

    for tabla in (recibe, paga):
        tabla["clave"] = tabla["NOMBRE"].map(normalizar)
    recibe = recibe.groupby("clave", as_index=False).agg(
        NOMBRE=("NOMBRE", "first"), RECIBE=("RECIBE", "sum")
    )
    paga = paga.groupby("clave", as_index=False).agg(
        NOMBRE_PAGA=("NOMBRE", "first"), PAGA=("PAGA", "sum")
    )

    df = recibe.merge(paga, on="clave", how="outer")
    cruzadas = int((df["NOMBRE"].notna() & df["NOMBRE_PAGA"].notna()).sum())
    registrar(
        f"  Resumen: {len(df):,} empresa(s); {cruzadas:,} reciben y pagan."
    )
    # El nombre visible es el del Propietario cuando la empresa recibe;
    # si solo paga, el del Suministrador de la prorrata.
    df["NOMBRE"] = df["NOMBRE"].fillna(df["NOMBRE_PAGA"])
    df = df.fillna({"RECIBE": 0.0, "PAGA": 0.0})
    df["NETO"] = df["RECIBE"] - df["PAGA"]
    # Lo que se recibe y lo que se paga son la MISMA plata mirada desde
    # los dos lados: el total tiene que cerrar. Si no cierra, es que
    # quedaron cuartos con monto y sin prorrata (la prorrata ya lo
    # informa) -- se deja dicho aca tambien, que es donde se ve.
    total_recibe = float(df["RECIBE"].sum())
    total_paga = float(df["PAGA"].sum())
    if abs(total_recibe - total_paga) > 1.0:
        registrar(
            f"  Resumen: se recibe ${total_recibe:,.0f} y se paga "
            f"${total_paga:,.0f} (diferencia ${total_recibe - total_paga:,.0f}); "
            "son la misma plata, la diferencia es lo que no se repartio."
        )
    return (
        df.loc[:, ["NOMBRE", "RECIBE", "PAGA", "NETO"]]
        .sort_values("NOMBRE", kind="stable").reset_index(drop=True)
    )
