import unittest

import pandas as pd

from Script import nucleo


class AlertasMapeosTest(unittest.TestCase):
    def test_agrupa_claves_ausentes_y_datos_vacios(self):
        mensajes = []

        faltantes = nucleo._avisar_claves_sin_mapeo(
            ["SAE-A", "SAE-A", "SAE-B", "SAE-C"],
            {"sae-a": "BARRA", "sae-b": ""},
            "Central sin barra",
            "Resumen BESS",
            registrar=mensajes.append,
        )

        self.assertEqual(faltantes, ["SAE-B", "SAE-C"])
        self.assertEqual(len(mensajes), 1)
        self.assertIn("[ALTA]", mensajes[0])
        self.assertIn("pueden quedar vacios o en 0", mensajes[0])

    def test_fd_avisa_diccionario_y_matches_que_se_rellenan_con_cero(self):
        mensajes = []
        df = pd.DataFrame({
            "clave": ["SAE-SIN-DIC", "SAE-CON-DIC"],
            "Y": [1, 1],
            "AC": [1, 5],
        })

        am, an, ap, aq = nucleo.calcular_fd_prorrateado(
            df,
            {"sae-con-dic": "Unidad X"},
            {},
            {},
            registrar=mensajes.append,
        )

        self.assertTrue(pd.isna(am.iloc[0]))
        self.assertEqual((am.iloc[1], an.iloc[1], ap.iloc[1], aq.iloc[1]),
                         (0.0, 0.0, 0.0, 0.0))
        self.assertTrue(any("no presentes en el diccionario" in m
                            for m in mensajes))
        self.assertTrue(any("CPF homologadas" in m for m in mensajes))
        self.assertTrue(any("CSF homologadas" in m for m in mensajes))


if __name__ == "__main__":
    unittest.main()
