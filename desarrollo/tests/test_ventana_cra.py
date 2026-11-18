"""La ventana del CRA con el formato del BESS, probada con la ventana real:

  - mismas columnas del diagrama (Estructura, Estado, Accion) y la fila
    "Carpeta del mes" arriba;
  - un boton por fila que se puede generar, del color de los del BESS;
  - "Actualizar" de una hoja corre el proceso, escribe el registro con
    encabezado y cierre, y lo guarda en <mes>/CRA/Log/.

Necesita tkinter Y un display (en Linux: `xvfb-run -a python -m
unittest discover`); sin ellos se saltea.
"""

import tempfile
import unittest
from pathlib import Path
from unittest import mock

from desarrollo.tests.test_ventana_rueda import _hay_ventana, tk


@unittest.skipUnless(_hay_ventana(), "sin tkinter o sin display")
class TestVentanaCra(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        import Balance_CRA as app
        from Script import cra

        cls.app = app
        cls.cra = cra
        cls.tmp = tempfile.TemporaryDirectory()
        cls.mes = Path(cls.tmp.name) / "Balance 2608"
        cra.crear_carpetas_caso(cls.mes)

        cfg = {
            "carpeta_base": str(cls.mes), "aamm": "2608",
            app.CLAVE_CARPETAS_PERIODO: {"2608": str(cls.mes)},
        }
        cls.parches = [
            mock.patch.object(app, "leer_config", lambda: dict(cfg)),
            mock.patch.object(app, "guardar_config", lambda data: None),
            mock.patch.object(app, "carpetas_del_bess", lambda: {}),
            mock.patch.object(app.messagebox, "showinfo", lambda *a, **k: None),
        ]
        for parche in cls.parches:
            parche.start()

        ventana = {}
        original = tk.Tk.mainloop
        tk.Tk.mainloop = lambda self: ventana.setdefault("root", self)
        try:
            app.main()
        finally:
            tk.Tk.mainloop = original
        cls.root = ventana["root"]
        cls.root.update()

    @classmethod
    def tearDownClass(cls):
        cls.root.destroy()
        for parche in cls.parches:
            parche.stop()
        cls.tmp.cleanup()

    @classmethod
    def _todos(cls, widget):
        for hijo in widget.winfo_children():
            yield hijo
            yield from cls._todos(hijo)

    def _textos(self, clase):
        return [
            str(w.cget("text")) for w in self._todos(self.root)
            if w.winfo_class() == clase
        ]

    def test_columnas_y_filas_como_el_bess(self):
        etiquetas = self._textos("Label")
        for titulo in ("Estructura", "Estado", "Accion", "Carpeta del mes",
                       "CRA/Energia/", "FD y FMA/"):
            self.assertIn(titulo, etiquetas)
        self.assertIn(self.app.SIMBOLO["falta"], etiquetas)

        botones = [
            w for w in self._todos(self.root)
            if w.winfo_class() == "Button" and w.cget("text") == "Actualizar"
            and w.cget("bg") == "#fdf0d5"
        ]
        # La planilla + una por hoja.
        self.assertEqual(len(botones), 1 + len(self.cra.SECCIONES))
        self.assertIn("Ir a este mes", self._textos("Button"))

    def test_actualizar_una_hoja_corre_y_guarda_el_registro(self):
        llamadas = []

        def falso(base, aamm, secciones=None, registrar=None, progreso=None):
            llamadas.append((base, aamm, secciones))
            registrar("Listo.")

        falso.__name__ = "generar_balance_cra"   # el titulo del registro

        def hilo_en_el_lugar(target, daemon=None):
            return mock.Mock(start=target)

        # El boton toma cra.generar_balance_cra al apretarlo.
        with mock.patch.object(self.cra, "generar_balance_cra", falso), \
                mock.patch.object(self.app.threading, "Thread", hilo_en_el_lugar):
            self._boton_hoja("FP").invoke()
            for _ in range(20):
                self.root.update()

        self.assertEqual(llamadas, [(str(self.mes), "2608", ["fp"])])
        registro = self.cra.resolver_rutas(self.mes, "2608")["registro"]
        self.assertEqual(registro.parent, self.mes / "CRA" / "Log")
        texto = registro.read_text(encoding="utf-8")
        self.assertIn("Planilla 5_REMUNERACION_CRA: fp", texto)
        self.assertIn("listo en", texto)

    def _boton_hoja(self, hoja):
        for fila in self._todos(self.root):
            if fila.winfo_class() != "Frame":
                continue
            hijos = list(self._todos(fila))
            textos = [str(w.cget("text")) for w in hijos if w.winfo_class() == "Label"]
            if f"hoja '{hoja}'" in textos and len(textos) < 6:
                return next(w for w in hijos if w.winfo_class() == "Button")
        raise AssertionError(f"no esta la fila de la hoja {hoja}")


if __name__ == "__main__":
    unittest.main()
