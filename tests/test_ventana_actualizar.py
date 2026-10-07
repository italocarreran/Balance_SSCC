"""Dos pedidos sobre la ventana del BESS, probados con la ventana real:

  - el boton "Actualizar" (junto a "Examinar") vuelve a mirar las
    carpetas: un archivo pegado a mano aparece sin cambiar de mes;
  - el boton de Medidas_SAE.xlsx arranca sin la advertencia previa
    ("Es el proceso mas lento... ¿Seguir?").

Necesita tkinter Y un display (en Linux: `xvfb-run -a python -m
unittest discover`); sin ellos se saltea, igual que la prueba de la
rueda.
"""

import tempfile
import unittest
from pathlib import Path
from unittest import mock

from tests.test_ventana_rueda import _hay_ventana, tk


@unittest.skipUnless(_hay_ventana(), "sin tkinter o sin display")
class TestVentanaActualizar(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        import Balance_BESS as app
        from Script import nucleo

        cls.app = app
        cls.nucleo = nucleo
        cls.tmp = tempfile.TemporaryDirectory()
        cls.caso = Path(cls.tmp.name) / "Caso 2609"
        for sub in nucleo.SUBCARPETAS_CASO:
            (cls.caso / sub).mkdir(parents=True, exist_ok=True)

        cfg = {
            "carpeta_base": str(cls.caso),
            "aamm": "2609",
            app.CLAVE_CARPETAS_PERIODO: {"2609": str(cls.caso)},
        }
        cls.parches = [
            mock.patch.object(app, "leer_config", lambda: dict(cfg)),
            mock.patch.object(app, "guardar_config", lambda data: None),
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

    def _fila(self, contiene):
        """
        El Frame de la fila del arbol cuyo nombre contiene ese texto
        (fila -> celda de estructura -> etiqueta del nombre). El nombre
        puede cambiar al del archivo real cuando aparece.
        """
        for w in self._todos(self.root):
            if (w.winfo_class() == "Label"
                    and contiene.lower() in str(w.cget("text")).lower()):
                return w.master.master
        self.fail(f"no hay fila con '{contiene}' en el arbol")

    def _estado_de_fila(self, contiene):
        """El simbolo de estado (ok / falta / ...) de esa fila."""
        simbolos = set(self.app.SIMBOLO.values())
        for w in self._todos(self._fila(contiene)):
            if w.winfo_class() == "Label" and w.cget("text") in simbolos:
                return w.cget("text")
        self.fail(f"la fila con '{contiene}' no tiene estado")

    def _boton(self, texto, dentro_de):
        for w in self._todos(dentro_de):
            if w.winfo_class() == "Button" and w.cget("text") == texto:
                return w
        self.fail(f"no hay boton '{texto}'")

    def _marco_carpeta(self):
        for w in self._todos(self.root):
            if (w.winfo_class() == "Labelframe"
                    and w.cget("text") == "Carpeta base del caso"):
                return w
        self.fail("no esta el marco 'Carpeta base del caso'")

    def test_actualizar_muestra_un_archivo_pegado_a_mano(self):
        etiqueta = "homologacion"
        antes = self._estado_de_fila(etiqueta)
        self.assertEqual(antes, self.app.SIMBOLO["falta"])

        aux = self.nucleo.resolver_rutas(self.caso)["auxiliares_dir"]
        archivo = aux / "Homologacion ClavesTF y PRMTE.xlsx"
        archivo.write_bytes(b"x")
        try:
            self.root.update()
            # Sin tocar nada, el diagrama todavia no lo ve.
            self.assertEqual(self._estado_de_fila(etiqueta), antes)

            self._boton("Actualizar", self._marco_carpeta()).invoke()
            self.root.update()

            self.assertEqual(
                self._estado_de_fila(etiqueta), self.app.SIMBOLO["ok"]
            )
        finally:
            archivo.unlink()
            self._boton("Actualizar", self._marco_carpeta()).invoke()
            self.root.update()

    def test_medidas_arranca_sin_advertencia(self):
        llamadas = []

        def generar_falso(**kwargs):
            llamadas.append(kwargs.get("aamm"))

        fila = self._fila(self.nucleo.ARCHIVO_MEDIDAS_SAE)
        boton = next(
            w for w in self._todos(fila) if w.winfo_class() == "Button"
        )

        class HiloEnElMismo:
            """
            Sin mainloop, un hilo aparte no puede tocar tkinter: el
            trabajo se corre en el hilo principal, que es lo mismo que
            hace el hilo real pero sin carrera.
            """

            def __init__(self, target, daemon=None, **_):
                self.target = target

            def start(self):
                self.target()

        preguntas = []

        with mock.patch.object(
            self.app.messagebox, "askyesno",
            side_effect=lambda *a, **k: preguntas.append(a) or False,
        ), mock.patch.object(
            self.nucleo, "generar_medidas_sae", generar_falso
        ), mock.patch.object(self.app.threading, "Thread", HiloEnElMismo), \
                mock.patch.object(self.app.messagebox, "showinfo"), \
                mock.patch.object(self.app.messagebox, "showerror"):
            boton.invoke()
            for _ in range(5):
                self.root.update()

        self.assertEqual(preguntas, [])

        self.assertEqual(llamadas, ["2609"])


if __name__ == "__main__":
    unittest.main()
