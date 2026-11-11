# -*- coding: utf-8 -*-
"""Importar Balance_BESS no importa lo pesado (pandas, openpyxl).

Plan de la ventana (desarrollo/PLAN_OPTIMIZACION_VENTANAS_BESS.md, Fase
2): la ventana aparece con un cartel "Cargando..." y Script.nucleo se
carga con ella ya a la vista. Se mira en un subproceso para no depender
de lo que ya importo la corrida de pruebas.
"""

import subprocess
import sys
import unittest
from pathlib import Path

from desarrollo.tests.test_ventana_rueda import _hay_ventana

RAIZ = Path(__file__).resolve().parents[2]


@unittest.skipUnless(_hay_ventana(), "sin tkinter o sin display")
class TestArranqueLiviano(unittest.TestCase):
    def _correr(self, codigo):
        salida = subprocess.run(
            [sys.executable, "-c", codigo], cwd=RAIZ,
            capture_output=True, text=True, timeout=120,
        )
        self.assertEqual(salida.returncode, 0, salida.stderr)
        return salida.stdout.strip()

    def test_importar_no_trae_pandas_ni_openpyxl(self):
        codigo = (
            "import sys, Balance_BESS; "
            "print('pandas' in sys.modules, 'openpyxl' in sys.modules, "
            "'Script.nucleo' in sys.modules)"
        )
        self.assertEqual(self._correr(codigo), "False False False")

    def test_archivo_registro_se_sigue_leyendo_del_modulo(self):
        codigo = (
            "import Balance_BESS as app; from Script import nucleo; "
            "print(app.ARCHIVO_REGISTRO == nucleo.ARCHIVO_REGISTRO)"
        )
        self.assertEqual(self._correr(codigo), "True")

    def test_cerrar_la_ventana_mientras_carga(self):
        # La ventana se cierra en el update() numero N: el primero (recien
        # aparecida) o uno del medio de la carga. main() vuelve sin
        # traceback y sin llegar a mainloop().
        for numero in (1, 3):
            with self.subTest(update=numero):
                codigo = (
                    "import tkinter as tk, time, Balance_BESS as app\n"
                    "lento = app._cargar_nucleo\n"
                    "def cargar():\n"
                    "    time.sleep(0.3)\n"
                    "    return lento()\n"
                    "app._cargar_nucleo = cargar\n"
                    "veces = [0]\n"
                    "original = tk.Tk.update\n"
                    "def update(self):\n"
                    "    veces[0] += 1\n"
                    f"    if veces[0] == {numero}:\n"
                    "        self.destroy()\n"
                    "    original(self)\n"
                    "tk.Tk.update = update\n"
                    "tk.Tk.mainloop = lambda self: print('mainloop')\n"
                    "app.main()\n"
                    "print('fin')\n"
                )
                self.assertEqual(self._correr(codigo), "fin")

    def test_si_la_carga_falla_se_muestra_el_error(self):
        codigo = (
            "import tkinter as tk, Balance_BESS as app\n"
            "def falla():\n"
            "    raise ImportError('falta un modulo')\n"
            "app._cargar_nucleo = falla\n"
            "app.mostrar_error = lambda titulo, texto, parent=None: "
            "print(titulo, '|', 'falta un modulo' in texto)\n"
            "tk.Tk.mainloop = lambda self: print('mainloop')\n"
            "app.main()\n"
            "print('fin')\n"
        )
        self.assertEqual(
            self._correr(codigo), "No se pudo cargar el programa | True\nfin"
        )

if __name__ == "__main__":
    unittest.main()
