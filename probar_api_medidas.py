# -*- coding: utf-8 -*-
"""
probar_api_medidas.py — diagnostico aparte de la API de medidas del
Coordinador (la misma que usa "Traer Medidas_SAE.xlsx").

Consulta la API EN VIVO para UN punto de medida (canales 1 y 3), sin
pasar por la descarga guardada en <CARPETA_BASE>/Medidas/_trabajo/, y
dice hasta que hora trae valores. Sirve para saber si un mes incompleto
es porque la API todavia no lo publico o porque la descarga guardada es
vieja.

Uso, desde la carpeta del programa:

    python probar_api_medidas.py 2609
    python probar_api_medidas.py 2609 ARENA_220_JT1_ARE
    python probar_api_medidas.py 2609 ARENA_220_JT1_ARE "T:\\...\\Medidas\\_trabajo"

  1er argumento: periodo AAMM (2609 = septiembre 2026).
  2do (opcional): idPuntoMedida (por defecto ARENA_220_JT1_ARE).
  3ro (opcional): carpeta _trabajo de Medidas, para comparar contra lo
     que tiene guardado la descarga.

Deja al lado un CSV con lo que devolvio la API, para abrirlo en Excel.
No modifica nada del caso.
"""

import sys
from pathlib import Path

import pandas as pd

from Script.Medidas import Claves_Balance
from Script.Medidas.Descarga_PRMTE import URL_MEDIDAS, CANALES, TIMEOUT
from Script.Medidas.comun import leer_clave_api, CLAVE_PRMTE


PUNTO_POR_DEFECTO = "ARENA_220_JT1_ARE"


def consultar(sesion, periodo, punto, canal, user_key):
    """Una llamada, sin reintentos ni silencios: muestra lo que pasa."""

    respuesta = sesion.get(
        URL_MEDIDAS.format(periodo=periodo),
        params={"idCanal": canal, "idPuntoMedida": punto,
                "user_key": user_key},
        timeout=TIMEOUT,
    )

    print(f"\n=== Canal {canal}: HTTP {respuesta.status_code}")

    if respuesta.status_code != 200:
        print(f"  respuesta: {respuesta.text[:500]}")
        return pd.DataFrame()

    datos = respuesta.json()

    if not datos or not isinstance(datos, list):
        print(f"  la API no devolvio registros: {str(datos)[:300]}")
        return pd.DataFrame()

    registro = datos[0]
    for campo in ("idPuntoMedida", "periodo", "fechaUltimaLectura"):
        print(f"  {campo}: {registro.get(campo)}")

    df = pd.DataFrame(registro.get("mediciones", []))
    print(f"  mediciones: {len(df):,} fila(s); columnas: {list(df.columns)}")

    if not df.empty:
        print("  primeras filas tal como vienen:")
        print(df.head(3).to_string(index=False))

    df["canal"] = canal
    return df


def resumir(df, titulo):
    print(f"\n##### {titulo}")

    if df.empty:
        print("  (sin filas)")
        return

    df = Claves_Balance.normalizar_fechas(df)

    local = Claves_Balance.COL_INTERVALO_LOCAL

    if "canal" in df:
        grupos = [
            (f"canal {c}", df[df["canal"] == c],
             "canalVal1" if c == 1 else "canalVal3")
            for c in sorted(df["canal"].unique())
        ]
    else:
        # Lo guardado no dice de que canal es cada fila: solo se puede
        # decir hasta donde llega cada columna.
        print(f"  filas del punto: {len(df):,}; el mes va del "
              f"{df[local].min()} al {df[local].max()}")
        for columna in ("canalVal1", "canalVal3"):
            if columna in df:
                con_valor = df[df[columna].notna()]
                print(f"  {columna}: {len(con_valor):,} valor(es)"
                      + (f", hasta {con_valor[local].max()}"
                         if not con_valor.empty else ""))
        grupos = []

    for etiqueta, parte, columna in grupos:
        if columna not in parte:
            print(f"  {etiqueta}: no trae {columna}")
            continue

        con_valor = parte[parte[columna].notna()]
        sin_valor = parte[parte[columna].isna()]

        print(f"\n  {etiqueta} ({columna}):")
        print(f"    filas: {len(parte):,}  con valor: {len(con_valor):,}"
              f"  vacias: {len(sin_valor):,}")
        print(f"    el mes va del {parte[local].min()} al {parte[local].max()}")
        if not con_valor.empty:
            print(f"    valores desde {con_valor[local].min()} "
                  f"hasta {con_valor[local].max()}")

        if not sin_valor.empty:
            # tramos seguidos de cuartos de hora sin valor
            orden = parte.sort_values(local).reset_index(drop=True)
            vacia = orden[columna].isna()
            tramo = (vacia != vacia.shift()).cumsum()
            print("    tramos sin valor:")
            for _, t in orden[vacia].groupby(tramo[vacia]):
                print(f"      {t[local].min()} -> {t[local].max()} "
                      f"({len(t):,} cuartos de hora)")

    # cambio de hora: intervaloUtc con mas de una hora local
    utc = Claves_Balance.COL_INTERVALO_UTC
    pares = df[[Claves_Balance.COL_INTERVALO_LOCAL, utc]].drop_duplicates()
    repetidos = pares[pares.duplicated(utc, keep=False)].sort_values(utc)
    if not repetidos.empty:
        print("\n  intervaloUtc con dos horas locales (cambio de hora):")
        print(repetidos.to_string(index=False))


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)

    aamm = sys.argv[1].strip()
    if len(aamm) != 4 or not aamm.isdigit():
        sys.exit(f"Periodo invalido: {aamm!r} (se espera AAMM, p. ej. 2609)")
    periodo = f"20{aamm[:2]}{aamm[2:]}"
    punto = sys.argv[2] if len(sys.argv) > 2 else PUNTO_POR_DEFECTO
    trabajo = Path(sys.argv[3]) if len(sys.argv) > 3 else None

    import requests

    user_key = leer_clave_api(CLAVE_PRMTE)
    print(f"Periodo {periodo}, punto de medida {punto}")
    print(f"URL: {URL_MEDIDAS.format(periodo=periodo)}")

    sesion = requests.Session()
    partes = [consultar(sesion, periodo, punto, c, user_key) for c in CANALES]
    api = pd.concat([p for p in partes if not p.empty], ignore_index=True) \
        if any(not p.empty for p in partes) else pd.DataFrame()

    resumir(api, "LO QUE LA API ENTREGA HOY")

    if not api.empty:
        salida = Path(__file__).with_name(f"probar_api_{punto}_{periodo}.csv")
        api.to_csv(salida, index=False, sep=";", decimal=",",
                   encoding="utf-8-sig")
        print(f"\nCSV con lo que devolvio la API: {salida}")

    if trabajo:
        lotes = sorted(trabajo.glob(f"medidas_batch_{periodo}_*.parquet"))
        if not lotes:
            print(f"\nNo hay lotes de {periodo} en {trabajo}")
        else:
            guardado = pd.concat(
                [pd.read_parquet(l) for l in lotes], ignore_index=True
            )
            guardado = guardado[guardado["idPuntoMedida"] == punto]
            resumir(guardado, f"LO QUE TIENE GUARDADO LA DESCARGA ({trabajo})")


if __name__ == "__main__":
    main()
