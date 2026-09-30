"""
Dos correcciones de la segunda tanda de la corrida real:

  - la generacion que viene de la API de operacion real llega en MWh y
    el balance entero trabaja en kWh;
  - el FD de "Calculo E Costos" separa "a esta unidad le falta una
    hora" de "esta unidad no esta en la hoja FD" (2.232 alertas FD-005
    que en realidad eran 3 hechos).
"""

import tempfile
import unittest
from pathlib import Path

import pandas as pd

from Script import nucleo
from Script.Medidas import Generacion_Real, Homologacion


def respuesta_de_una_hora(mwh):
    """Una hora de la API: se abre en 4 cuartos de hora."""

    return pd.DataFrame({
        "fecha_consulta": ["2026-07-01"],
        "topologyName": ["SAE PFV Andes Solar III (Inyección)"],
        "topologyTimeHour": [1],
        "measure": [mwh],
    })


class UnidadDeGeneracionRealTest(unittest.TestCase):

    def central(self, unidad=None, factor=1):
        ficha = {
            "topologyName": "SAE PFV Andes Solar III (Inyección)",
            "clave": "SAE-CRCA-PFV-ANDES3",
            "factor": factor,
        }
        if unidad is not None:
            ficha["unidad"] = unidad
        return [ficha]

    def test_mwh_se_pasa_a_kwh(self):
        """170,78 MW durante una hora = 42,7 MWh por cuarto = 42.700 kWh."""

        salida = Generacion_Real.expandir_a_cuartos(
            respuesta_de_una_hora(170.78), self.central("mwh"),
            registrar=lambda _: None,
        )

        self.assertEqual(len(salida), 4)
        for valor in salida["Gen_Unidad"]:
            self.assertAlmostEqual(valor, 170.78 / 4 * 1000)

    def test_sin_unidad_se_asume_mwh(self):
        """Es lo que devuelve la API: el default no puede ser 'no convertir'."""

        salida = Generacion_Real.expandir_a_cuartos(
            respuesta_de_una_hora(40.0), self.central(),
            registrar=lambda _: None,
        )

        self.assertAlmostEqual(salida["Gen_Unidad"].iloc[0], 10_000.0)

    def test_kwh_no_se_convierte(self):
        salida = Generacion_Real.expandir_a_cuartos(
            respuesta_de_una_hora(40.0), self.central("kwh"),
            registrar=lambda _: None,
        )

        self.assertAlmostEqual(salida["Gen_Unidad"].iloc[0], 10.0)

    def test_la_unidad_y_el_signo_se_multiplican(self):
        salida = Generacion_Real.expandir_a_cuartos(
            respuesta_de_una_hora(40.0), self.central("mwh", factor=-1),
            registrar=lambda _: None,
        )

        self.assertAlmostEqual(salida["Gen_Unidad"].iloc[0], -10_000.0)

    def test_la_unidad_sale_de_la_columna_canal(self):
        with tempfile.TemporaryDirectory() as carpeta:
            ruta = Path(carpeta) / "Homologacion.xlsx"

            gen_real = pd.DataFrame({
                "clave": ["SAE-UNO", "SAE-DOS", "SAE-TRES"],
                "Punto de Medida": ["Uno", "Dos", "Tres"],
                "Canal": ["MWh", "kWh", None],
                "Flujo": [1, 1, -1],
            })

            with pd.ExcelWriter(ruta) as w:
                pd.DataFrame({
                    "clave": [], "Punto de Medida": [], "Canal": [],
                    "Flujo": [],
                }).to_excel(w, sheet_name="homol", index=False)
                gen_real.to_excel(w, sheet_name="Gen real", index=False)

            centrales = Homologacion.leer_gen_real(ruta)

        self.assertEqual(
            [c["unidad"] for c in centrales], ["mwh", "kwh", "mwh"]
        )
        self.assertEqual([c["factor"] for c in centrales], [1.0, 1.0, -1.0])

    def test_el_canal_trae_la_unidad_con_sufijo(self):
        """
        En el archivo real el canal viene "MWhD"/"MWhR" (la unidad con
        el tipo de medida pegado atras). Se lee el PRINCIPIO: lo que
        decide es la unidad, no la D ni la R del final.
        """

        for canal, esperada in (
            ("MWhD", "mwh"), ("MWhR", "mwh"),
            ("kWhD", "kwh"), ("kWhR", "kwh"),
            ("MWh", "mwh"), ("kwh", "kwh"),
            (" kWh D ", "kwh"),
            # Vacio o cualquier otra cosa: MWh, que es lo que devuelve
            # la API de operacion real.
            ("", "mwh"), (None, "mwh"), ("R", "mwh"), ("Potencia", "mwh"),
        ):
            with self.subTest(canal=canal):
                self.assertEqual(
                    Homologacion.unidad_desde_canal(canal), esperada
                )

    def test_un_canal_en_kwh_con_sufijo_no_se_multiplica_por_mil(self):
        """
        La prueba de que el cambio se nota: la misma medida de la API
        (MWh) entra x1000 con "MWhR" y tal cual con "kWhR".
        """

        respuesta = respuesta_de_una_hora(2.0)

        en_mwh = Generacion_Real.expandir_a_cuartos(
            respuesta,
            self.central(Homologacion.unidad_desde_canal("MWhR")),
            registrar=lambda _: None,
        )
        en_kwh = Generacion_Real.expandir_a_cuartos(
            respuesta,
            self.central(Homologacion.unidad_desde_canal("kWhR")),
            registrar=lambda _: None,
        )

        self.assertEqual(
            en_mwh["Gen_Unidad"].sum(), en_kwh["Gen_Unidad"].sum() * 1000
        )


def bloque_fd(unidades, horas, columna_mas, columna_menos):
    """Un bloque de la hoja FD con esas unidades y esas horas."""

    filas = [
        {
            "id": f"{hora}{unidad}",
            "Hora Mes": hora,
            "Unidad": unidad,
            columna_mas: 1.0,
            columna_menos: 1.0,
        }
        for unidad in unidades for hora in horas
    ]
    return pd.DataFrame(filas)


class FdFaltanteTest(unittest.TestCase):
    """
    El caso real: una central homologada a una unidad que la hoja FD no
    tiene. Antes salía una alerta por hora; ahora sale una por central,
    con la lista de unidades entre las que elegir.
    """

    def setUp(self):
        self.horas = list(range(1, 25))
        self.cpf = bloque_fd(
            ["BESS PFV MANZANO", "BESS PFV MARIA ELENA"], self.horas,
            "CPF(+)", "CPF(-)",
        )
        self.csf = bloque_fd(
            ["BESS PFV MANZANO", "BESS PFV MARIA ELENA"], self.horas,
            "CSF(+)", "CSF(-)",
        )
        # Una central homologada a una unidad que NO está en la hoja FD.
        self.mapeo = {
            "sae-crca-pfv-manzano": "BESS PFV MANZANO",
            "sae-crca-pfv-nuevo-quillagua-2": "SAE-CRCA-PFV-NUEVO-QUILLAGUA-2",
        }
        # Una fila por central y hora (Y/AC en cuartos: bloque = hora).
        self.df = pd.DataFrame([
            {"clave": central, "Y": hora * 4, "AC": hora * 4}
            for central in ("SAE-CRCA-PFV-MANZANO",
                            "SAE-CRCA-PFV-NUEVO-QUILLAGUA-2")
            for hora in self.horas
        ])

    def calcular(self, **extra):
        registro = nucleo.Registro(salida=lambda _: None)
        nucleo.calcular_fd_prorrateado(
            self.df, self.mapeo,
            nucleo.construir_dic_fd_bloque(self.csf, "id", "CSF(+)", "CSF(-)"),
            nucleo.construir_dic_fd_bloque(self.cpf, "id", "CPF(+)", "CPF(-)"),
            registrar=registro, **extra,
        )
        return registro

    def test_una_alerta_por_central_y_no_una_por_hora(self):
        registro = self.calcular(
            unidades_fd_csf=nucleo.unidades_bloque_fd(self.csf),
            unidades_fd_cpf=nucleo.unidades_bloque_fd(self.cpf),
        )

        ids = [a.id_alerta for a in registro.alertas]
        self.assertEqual(ids, ["FD-007", "FD-007"])   # una CPF, una CSF
        self.assertNotIn("FD-005", ids)

        # Y dice de qué central se trata y qué unidades hay para elegir.
        texto = registro.lineas[0]
        self.assertIn("SAE-CRCA-PFV-NUEVO-QUILLAGUA-2", texto)
        self.assertIn("BESS PFV MARIA ELENA", texto)

    def test_sin_la_lista_de_unidades_se_comporta_como_antes(self):
        """Compatibilidad: sin los dos sets nuevos, todo sigue en FD-005."""

        registro = self.calcular()

        ids = {a.id_alerta for a in registro.alertas}
        self.assertEqual(ids, {"FD-005"})
        self.assertEqual(len(registro.alertas), 24 + 24)

    def test_a_una_unidad_presente_le_falta_una_hora_sola(self):
        """Ese sí es FD-005: la unidad está, la hora no."""

        self.cpf = self.cpf[
            ~((self.cpf["Unidad"] == "BESS PFV MANZANO")
              & (self.cpf["Hora Mes"] == 7))
        ].reset_index(drop=True)

        registro = self.calcular(
            unidades_fd_csf=nucleo.unidades_bloque_fd(self.csf),
            unidades_fd_cpf=nucleo.unidades_bloque_fd(self.cpf),
        )

        faltantes = [a for a in registro.alertas if a.id_alerta == "FD-005"]
        self.assertEqual(len(faltantes), 1)
        self.assertIn("bess pfv manzano", faltantes[0].clave)


if __name__ == "__main__":
    unittest.main()
