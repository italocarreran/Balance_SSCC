# -*- coding: utf-8 -*-
"""
La matriz de control: cada central y si se encontro en cada origen.

Pedido del usuario (2026-11-03): una matriz en el control de corrida con
las centrales en las filas y los origenes en las columnas.
"""

import tempfile
import unittest
from pathlib import Path

import openpyxl
import pandas as pd

from Script import nucleo
from Script.nucleo.matriz_centrales import construir_matriz_centrales

DICCIONARIO = pd.DataFrame([
    [None] * 5,
    ["Balance_BESS", "FD", "Subastas", "Ofertas", "FMA_CPF"],
    ["SAE-UNO", "SAE-UNO", "BAT_UNO", "BESS UNO OF", "Uno - BESS"],
])
RESUMEN = pd.DataFrame({
    "Nombre activo": ["SAE-UNO", "SAE-DOS"], "Barra inyección": ["B1", ""],
})


def matriz():
    return construir_matriz_centrales(
        resumen=RESUMEN, diccionario=DICCIONARIO,
        df_medidores=pd.DataFrame({
            "clave": ["SAE-UNO"] * 3 + ["SAE-DOS"], "SoC": [0.5, 0.5, None, None],
        }),
        df_wxy=pd.DataFrame({
            "Nombre": ["BESS UNO OF", "BESS UNO OF", "SAE-DOS", "SAE-DOS"],
            "Central": ["SAE-UNO", "SAE-UNO", "SAE-DOS", "SAE-DOS"],
            "Dia": [1, 2, 1, 2], "Oferta completa": [1, 0, 0, 0],
        }),
        df_cmg=pd.DataFrame({"A": [""], "B": [""], "C": [""], "D": ["B1"]}),
        df_subastas=pd.DataFrame({"Configuración": ["BAT_UNO", "BAT_UNO"]}),
        # La segunda fila es la hora que no existe (cambio de hora): no cuenta.
        df_fd_csf=pd.DataFrame({"Central": ["SAE-UNO"] * 2, "Hora Mes": [1, None]}),
        df_fd_cpf=pd.DataFrame({"Central": []}),
        tablas_fma={"cpf": pd.DataFrame({"Central": ["Uno - BESS"]})},
    ).set_index("Central")


class TestMatrizCentrales(unittest.TestCase):
    def test_marca_lo_encontrado_por_los_mismos_cruces(self):
        m = matriz()
        uno, dos = m.loc["SAE-UNO"], m.loc["SAE-DOS"]
        self.assertEqual(uno["Diccionario"], "Sí")
        self.assertEqual(dos["Diccionario"], "No")
        self.assertEqual(uno["Medidores (Medidas_SAE)"], "Sí (3 cuartos)")
        self.assertEqual(uno["SoC"], "Sí (2 cuartos)")
        self.assertEqual(dos["SoC"], "No")
        self.assertEqual(uno["CMg (por barra)"], "Sí (B1)")
        self.assertEqual(dos["CMg (por barra)"], "No (sin barra)")
        # Subastas y FMA CPF se encuentran por el nombre del Diccionario.
        self.assertEqual(uno["Subastas"], "Sí (2 filas)")
        self.assertEqual(uno["FMA CPF"], "Sí (1 horas)")
        self.assertEqual(uno["FD CSF"], "Sí (1 horas)")
        # Sin servicio en Subastas (no trae Concepto): puede no estar.
        self.assertEqual(uno["FD CPF"], "No (sin CPF en Subastas)")
        self.assertEqual(uno["Ofertas SSCC"], "Sí (1 de 2 días completos)")
        self.assertTrue(dos["Ofertas SSCC"].startswith("No"))

    def test_va_al_control_pintada(self):
        with tempfile.TemporaryDirectory() as carpeta:
            ruta = Path(carpeta) / nucleo.CARPETA_LOG / nucleo.ARCHIVO_CONTROL
            nucleo.escribir_control(
                ruta, df_matriz_centrales=matriz().reset_index()
            )
            ws = openpyxl.load_workbook(ruta)[nucleo.HOJA_MATRIZ_CENTRALES]
        nombres = {c.value: c.column_letter for c in ws[3] if c.value}
        fila = next(r for r in range(4, 6) if ws[f"B{r}"].value == "SAE-DOS")
        celda = ws[f"{nombres['Diccionario']}{fila}"]
        self.assertEqual(celda.value, "No")
        self.assertEqual(celda.fill.start_color.rgb[-6:], "F8CBAD")
        self.assertTrue(str(ws["B2"].value).startswith("Cuadro N° 1"))


class TestMatrizParcialYColores(unittest.TestCase):
    """
    Pedido del usuario (2026-11-10): la matriz sale tambien en una
    corrida parcial (lo que no se corrio, "Aún no se corre" en gris);
    lo adjudicado en Subastas manda: si se adjudico CPF tiene que tener
    FD CPF y FMA CPF (si no, rojo); no tener FD CSF cuando solo presta
    CPF es normal (amarillo).
    """

    def subastas(self):
        return pd.DataFrame({
            "Concepto": ["CPF(+)", "CPF(-)", "CSF(+)"],
            "Configuración": ["BAT_UNO", "BAT_UNO", "SAE-DOS"],
        })

    def test_parcial_lo_no_corrido_queda_pendiente(self):
        from Script.nucleo import matriz_centrales as mc

        m = construir_matriz_centrales(
            resumen=RESUMEN, diccionario=DICCIONARIO,
            df_subastas=self.subastas(),
        )
        estados = m.attrs["estados"]
        self.assertEqual(list(m.columns[1:]), list(mc.COLUMNAS_MATRIZ))
        m = m.set_index("Central")
        for columna in ("FD CSF", "FD CPF", "FMA CPF", "Medidores (Medidas_SAE)",
                        "Ofertas SSCC", "CMg (por barra)"):
            self.assertEqual(m.loc["SAE-UNO", columna], mc.TEXTO_PENDIENTE)
            self.assertEqual(estados[("SAE-UNO", columna)], mc.PENDIENTE)
        self.assertEqual(m.loc["SAE-UNO", "Subastas"], "Sí (2 filas; CPF)")
        self.assertEqual(estados[("SAE-UNO", "Subastas")], mc.OK)

    def test_lo_adjudicado_manda_sobre_fd_fma_y_ofertas(self):
        from Script.nucleo import matriz_centrales as mc

        m = construir_matriz_centrales(
            resumen=RESUMEN, diccionario=DICCIONARIO,
            df_subastas=self.subastas(),
            df_fd_csf=pd.DataFrame({"Central": ["SAE-DOS"], "Hora Mes": [1]}),
            df_fd_cpf=pd.DataFrame({"Central": []}),
            tablas_fma={"cpf": pd.DataFrame({"Central": []})},
            df_wxy=pd.DataFrame({
                "Nombre": ["SAE-UNO", "SAE-DOS"],
                "Central": ["SAE-UNO", "SAE-DOS"],
                "Dia": [1, 1], "Oferta completa": [0, 1],
            }),
        )
        e = m.attrs["estados"]
        m = m.set_index("Central")
        # SAE-UNO se adjudico CPF: sin FD CPF ni FMA CPF -> rojo.
        self.assertEqual(e[("SAE-UNO", "FD CPF")], mc.FALTA)
        self.assertIn("adjudicada CPF", m.loc["SAE-UNO", "FD CPF"])
        self.assertEqual(e[("SAE-UNO", "FMA CPF")], mc.FALTA)
        # ...pero no CSF: no estar en el FD de CSF es normal.
        self.assertEqual(e[("SAE-UNO", "FD CSF")], mc.AVISO)
        # SAE-DOS solo CSF: esta en FD CSF (verde), CPF amarillo.
        self.assertEqual(e[("SAE-DOS", "FD CSF")], mc.OK)
        self.assertEqual(e[("SAE-DOS", "FD CPF")], mc.AVISO)
        # Adjudicada y sin oferta en el archivo: rojo.
        self.assertEqual(e[("SAE-UNO", "Ofertas SSCC")], mc.FALTA)
        self.assertEqual(e[("SAE-DOS", "Ofertas SSCC")], mc.OK)

    def test_sin_subastas_lo_que_falta_es_aviso(self):
        from Script.nucleo import matriz_centrales as mc

        m = construir_matriz_centrales(
            resumen=RESUMEN, diccionario=DICCIONARIO,
            df_fd_cpf=pd.DataFrame({"Central": []}),
        )
        e = m.attrs["estados"]
        self.assertEqual(e[("SAE-UNO", "FD CPF")], mc.AVISO)
        self.assertEqual(
            m.set_index("Central").loc["SAE-UNO", "FD CPF"],
            "No (Subastas aún no se corre)",
        )

    def test_el_control_pinta_los_cuatro_colores(self):
        m = construir_matriz_centrales(
            resumen=RESUMEN, diccionario=DICCIONARIO,
            df_subastas=self.subastas(),
            df_fd_cpf=pd.DataFrame({"Central": ["SAE-DOS"], "Hora Mes": [1]}),
        )
        with tempfile.TemporaryDirectory() as carpeta:
            ruta = Path(carpeta) / nucleo.ARCHIVO_CONTROL
            nucleo.escribir_control(ruta, df_matriz_centrales=m)
            ws = openpyxl.load_workbook(ruta)[nucleo.HOJA_MATRIZ_CENTRALES]
        nombres = {c.value: c.column_letter for c in ws[3] if c.value}
        fila = {ws[f"B{r}"].value: r for r in range(4, 6)}

        def color(central, columna):
            return ws[f"{nombres[columna]}{fila[central]}"].fill.start_color.rgb[-6:]

        self.assertEqual(color("SAE-UNO", "Subastas"), "E2EFDA")   # verde
        self.assertEqual(color("SAE-UNO", "FD CPF"), "F8CBAD")     # rojo
        self.assertEqual(color("SAE-DOS", "FD CPF"), "E2EFDA")
        self.assertEqual(color("SAE-DOS", "FMA CPF"), "D9D9D9")    # gris
        m2 = construir_matriz_centrales(
            resumen=RESUMEN, diccionario=DICCIONARIO,
            df_subastas=self.subastas(),
            df_fd_csf=pd.DataFrame({"Central": []}),
        )
        with tempfile.TemporaryDirectory() as carpeta:
            ruta = Path(carpeta) / nucleo.ARCHIVO_CONTROL
            nucleo.escribir_control(ruta, df_matriz_centrales=m2)
            ws = openpyxl.load_workbook(ruta)[nucleo.HOJA_MATRIZ_CENTRALES]
        nombres = {c.value: c.column_letter for c in ws[3] if c.value}
        fila = {ws[f"B{r}"].value: r for r in range(4, 6)}
        self.assertEqual(
            ws[f"{nombres['FD CSF']}{fila['SAE-UNO']}"].fill.start_color.rgb[-6:],
            "FFF2CC",  # amarillo: solo presta CPF
        )
        textos = [ws.cell(r, c).value for r in range(2, 8) for c in range(1, ws.max_column + 1)]
        self.assertIn("No está y debería estar", textos)

    def test_desde_una_planilla_que_solo_tiene_subastas(self):
        """Corrida parcial real: el libro tiene Centrales y Subastas."""

        from Script.nucleo import matriz_centrales as mc
        from Script.nucleo.proceso import matriz_desde_consolidado

        subastas = pd.DataFrame({
            "Concepto": ["CPF(+)"], "Fecha": ["2026-10-01"], "Año": [2026],
            "Mes": [10], "Dia": [1], "Hora_dia": [1], "Hora_mes": [1],
            "Configuración": ["BAT_UNO"], "Propietario": [""],
            "Energía SSCC": [5.0],
        })
        with tempfile.TemporaryDirectory() as carpeta:
            ruta = Path(carpeta) / "Balance_BESS.xlsx"
            nucleo.escribir_salida(
                None, ruta, [], [], df_subastas=subastas,
                centrales=(RESUMEN, DICCIONARIO),
                hojas_regenerar={nucleo.HOJA_SUBASTAS, nucleo.HOJA_CENTRALES},
                registrar=lambda *a: None,
            )
            m = matriz_desde_consolidado(ruta, registrar=lambda *a: None)
        e = m.attrs["estados"]
        self.assertEqual(e[("SAE-UNO", "Subastas")], mc.OK)
        self.assertEqual(e[("SAE-UNO", "FD CPF")], mc.PENDIENTE)
        self.assertEqual(e[("SAE-UNO", "Medidores (Medidas_SAE)")], mc.PENDIENTE)
        self.assertEqual(e[("SAE-UNO", "Diccionario")], mc.OK)


if __name__ == "__main__":
    unittest.main()
