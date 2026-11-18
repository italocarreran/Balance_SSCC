# -*- coding: utf-8 -*-
"""Herramientas/Actualizar_desde_GitHub.py (rehecha 2026-11-18).

Pedido del usuario: mas simple, con jerarquia, solo lo que se ocupa para
correr (los .py de las ventanas, Herramientas/ y Script/), casillas en las
carpetas (cerradas hasta apretar la flecha), marcado por defecto lo que
cambio, y la lista desplegable de ramas en la ventana.
"""

import importlib.util
import io
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest import mock

from desarrollo.tests.test_ventana_rueda import _hay_ventana, tk

RAIZ = Path(__file__).resolve().parents[2]


def _modulo():
    spec = importlib.util.spec_from_file_location(
        "actualizar_github", RAIZ / "Herramientas" / "Actualizar_desde_GitHub.py")
    modulo = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(modulo)
    return modulo


def _zip(archivos):
    datos = io.BytesIO()
    with zipfile.ZipFile(datos, "w") as z:
        for ruta, contenido in archivos.items():
            z.writestr(f"Balance_SSCC-main/{ruta}", contenido)
    return datos.getvalue()


REMOTOS = {
    "Balance_BESS.py": b"nuevo",
    "Balance_CRA.py": b"nuevo",
    "README.md": b"no va",
    "requirements.txt": b"pandas",
    "Script/nucleo/re545.py": b"cambiado",
    "Script/nucleo/igual.py": b"igual",
    "Script/nucleo/nuevo.py": b"nuevo",
    "Herramientas/Comparar_Meses.py": b"cambiado",
    "desarrollo/tests/x.py": b"no va",
    "docs/a.md": b"no va",
    "__config__/tks.json": b"nunca",
}


def _carpeta_local(base):
    for ruta, contenido in {
        "Balance_BESS.py": b"viejo", "Balance_CRA.py": b"viejo",
        "requirements.txt": b"pandas",
        "Script/nucleo/re545.py": b"viejo", "Script/nucleo/igual.py": b"igual\r\n".replace(b"\r\n", b""),
        "Herramientas/Comparar_Meses.py": b"viejo",
    }.items():
        (base / ruta).parent.mkdir(parents=True, exist_ok=True)
        (base / ruta).write_bytes(contenido)


class TestComparar(unittest.TestCase):

    def test_solo_lo_que_se_ocupa_para_correr(self):
        m = _modulo()
        self.assertTrue(m.en_alcance("Balance_BESS.py"))
        self.assertTrue(m.en_alcance("Balance_CRA.py"))
        self.assertTrue(m.en_alcance("requirements.txt"))
        self.assertTrue(m.en_alcance("Script/nucleo/re545.py"))
        self.assertTrue(m.en_alcance("Herramientas/Comparar_Meses.py"))
        for fuera in ("README.md", "desarrollo/tests/x.py", "docs/a.md",
                      "__config__/tks.json", "Script/__pycache__/a.pyc"):
            self.assertFalse(m.en_alcance(fuera), fuera)

    def test_compara_y_aplica(self):
        m = _modulo()
        with tempfile.TemporaryDirectory() as carpeta:
            base = Path(carpeta)
            _carpeta_local(base)
            cambios = m.comparar(REMOTOS, base)
            self.assertEqual(cambios, {
                "Balance_BESS.py": "Modificado",
                "Balance_CRA.py": "Modificado",
                "Herramientas/Comparar_Meses.py": "Modificado",
                "Script/nucleo/nuevo.py": "Nuevo",
                "Script/nucleo/re545.py": "Modificado",
            })
            m.aplicar(REMOTOS, ["Script/nucleo/re545.py", "Script/nucleo/nuevo.py"], base)
            self.assertEqual((base / "Script/nucleo/re545.py").read_bytes(), b"cambiado")
            self.assertEqual((base / "Script/nucleo/nuevo.py").read_bytes(), b"nuevo")
            # Sin respaldo (pedido del usuario 2026-11-18).
            self.assertFalse((base / "__config__" / "respaldos_actualizacion").exists())
            self.assertEqual((base / "Balance_BESS.py").read_bytes(), b"viejo")


@unittest.skipUnless(_hay_ventana(), "sin tkinter o sin display")
class TestVentana(unittest.TestCase):

    def setUp(self):
        self.m = _modulo()
        self.tmp = tempfile.TemporaryDirectory()
        self.base = Path(self.tmp.name)
        _carpeta_local(self.base)
        self.parches = [
            mock.patch.object(self.m, "AQUI", self.base),
            mock.patch.object(self.m, "carpeta_config", lambda: self.base / "__config__"),
            mock.patch.object(self.m, "bajar_zip", lambda rama: _zip(REMOTOS)),
            mock.patch.object(self.m, "listar_ramas", lambda: ["main", "claude/una-rama"]),
        ]
        for parche in self.parches:
            parche.start()
        self.root = tk.Tk()
        self.v = self.m.Ventana(self.root)
        # Al abrir solo carga las ramas; no revisa hasta elegir una.
        self._esperar(lambda: "claude/una-rama" in self.v.combo.cget("values"))
        self.assertEqual(self.v.cambios, {})
        self.v.combo.set("main")
        self.v.combo.event_generate("<<ComboboxSelected>>")
        self._esperar(lambda: self.v.cambios)

    def tearDown(self):
        self.root.destroy()
        for parche in self.parches:
            parche.stop()
        self.tmp.cleanup()

    def _esperar(self, condicion, segundos=5):
        import time
        fin = time.time() + segundos
        while time.time() < fin:
            self.root.update()
            if condicion():
                return
            time.sleep(0.02)
        self.fail("no se cumplio a tiempo")

    def test_todo_marcado_y_carpetas_cerradas(self):
        self.assertEqual(self.v.marcados, set(self.v.cambios))
        script = self.v.item_de["Script"]
        self.assertFalse(self.v.arbol.item(script, "open"))
        self.assertTrue(self.v.arbol.item(script, "text").startswith(self.m.MARCADO))
        self.assertIn("claude/una-rama", self.v.combo.cget("values"))
        self.assertIn("(5)", self.v.btn_aplicar.cget("text"))
        self.assertIn("Balance_CRA.py", self.v.item_de)

    def test_la_carpeta_marca_lo_de_adentro(self):
        self.v.alternar("Script/")
        self.assertNotIn("Script/nucleo/re545.py", self.v.marcados)
        self.assertNotIn("Script/nucleo/nuevo.py", self.v.marcados)
        self.assertIn("Balance_BESS.py", self.v.marcados)
        self.v.alternar("Script/nucleo/re545.py")
        self.assertTrue(self.v.arbol.item(self.v.item_de["Script"], "text").startswith(self.m.PARCIAL))
        self.v.alternar("Script/")
        self.assertIn("Script/nucleo/nuevo.py", self.v.marcados)


if __name__ == "__main__":
    unittest.main()
