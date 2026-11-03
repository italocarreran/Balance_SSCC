"""
escritura.escribir_tabla_rapida (Medidas_SAE.xlsx, cmg.xlsx,
subastas_AAMM.xlsx): celda por celda igual que df.to_excel(index=False).
"""

import tempfile
import unittest
from pathlib import Path

import numpy as np
import openpyxl
import pandas as pd

from Script.nucleo import escritura


def _celdas(ruta):
    ws = openpyxl.load_workbook(ruta)["H"]
    return [
        (c.coordinate, c.value, c.number_format, c.font.b,
         c.border.left.style, c.alignment.horizontal)
        for fila in ws.iter_rows() for c in fila
        if c.value is not None or c.has_style
    ], ws.max_row, ws.max_column


class TestTablaRapida(unittest.TestCase):

    def test_igual_que_to_excel(self):
        n = 1500
        rng = np.random.default_rng(3)
        df = pd.DataFrame({
            "entero": np.arange(n),
            "real": rng.uniform(-1e6, 1e6, n),
            "texto": [None if i % 7 == 0 else f"C & <{i % 4}>" for i in range(n)],
            "vacio": "",
            "fecha": pd.Timestamp("2026-09-01") + pd.to_timedelta(np.arange(n) * 15, "min"),
            "dia": [pd.Timestamp("2026-09-01").date()] * n,
            "logico": [i % 2 == 0 for i in range(n)],
        })
        df.loc[::5, "real"] = np.nan
        with tempfile.TemporaryDirectory() as carpeta:
            a = Path(carpeta) / "a.xlsx"
            b = Path(carpeta) / "b.xlsx"
            df.to_excel(a, sheet_name="H", index=False)
            b.write_bytes(b"viejo")  # se reemplaza entero
            escritura.escribir_tabla_rapida(df, b, "H")
            self.assertEqual(_celdas(a), _celdas(b))
            pd.testing.assert_frame_equal(
                pd.read_excel(a, sheet_name="H"), pd.read_excel(b, sheet_name="H")
            )
            self.assertEqual(sorted(p.name for p in Path(carpeta).iterdir()),
                             ["a.xlsx", "b.xlsx"])


if __name__ == "__main__":
    unittest.main()
