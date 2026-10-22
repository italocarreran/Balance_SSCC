"""La rueda del mouse en la ventana: cada zona desplaza lo suyo.

Necesita tkinter Y un display. Sin alguno de los dos la prueba se
saltea (en Windows, que es donde corre el programa, estan los dos; en
un servidor sin pantalla anda con `xvfb-run`).
"""

import os
import unittest

try:
    import tkinter as tk
except ImportError:  # pragma: no cover - depende del interprete
    tk = None


def _hay_ventana():
    if tk is None:
        return False
    if os.name != "nt" and not os.environ.get("DISPLAY"):
        return False
    try:
        raiz = tk.Tk()
    except Exception:  # pragma: no cover - servidor sin pantalla
        return False
    raiz.destroy()
    return True


@unittest.skipUnless(_hay_ventana(), "sin tkinter o sin display")
class TestRuedaDelMouse(unittest.TestCase):
    """Abre la ventana real y le manda eventos de rueda de verdad."""

    @classmethod
    def setUpClass(cls):
        import Balance_BESS as app

        cls.app = app
        # main() termina en mainloop(), que bloquea: se intercepta para
        # quedarse con la ventana ya armada y manejarla desde aca.
        ventana = {}
        original = tk.Tk.mainloop
        tk.Tk.mainloop = lambda self: ventana.setdefault("root", self)
        try:
            app.main()
        finally:
            tk.Tk.mainloop = original
        cls.root = ventana["root"]
        cls.root.geometry("1000x600")
        cls.root.update()
        cls.canvas = cls._buscar(cls.root, "Canvas")
        cls.txt = cls._buscar(cls.root, "Text")
        for i in range(200):
            cls.txt.insert("end", f"linea de registro numero {i}\n")
        cls.root.update()

    @classmethod
    def tearDownClass(cls):
        cls.root.destroy()

    @classmethod
    def _buscar(cls, widget, clase):
        for hijo in widget.winfo_children():
            if hijo.winfo_class() == clase:
                return hijo
            encontrado = cls._buscar(hijo, clase)
            if encontrado is not None:
                return encontrado
        return None

    def _rueda(self, widget, delta, veces=1):
        x = widget.winfo_rootx() + widget.winfo_width() // 2
        y = widget.winfo_rooty() + widget.winfo_height() // 2
        for _ in range(veces):
            widget.event_generate(
                "<MouseWheel>", delta=delta, rootx=x, rooty=y,
                x=widget.winfo_width() // 2, y=widget.winfo_height() // 2,
            )
        self.root.update()

    def _reiniciar(self):
        self.txt.yview_moveto(0.0)
        self.canvas.yview_moveto(0.0)
        self.root.update()

    def test_sobre_el_registro_solo_se_mueve_el_registro(self):
        self._reiniciar()
        self._rueda(self.txt, -120, 3)
        self.assertGreater(self.txt.yview()[0], 0.0)
        self.assertEqual(self.canvas.yview()[0], 0.0)

    def test_fuera_del_registro_solo_se_mueve_la_ventana(self):
        self._reiniciar()
        self._rueda(self.canvas, -120, 3)
        self.assertEqual(self.txt.yview()[0], 0.0)
        self.assertGreater(self.canvas.yview()[0], 0.0)

    def test_el_registro_al_final_no_arrastra_la_ventana(self):
        self._reiniciar()
        self.txt.yview_moveto(1.0)
        self.root.update()
        antes = self.canvas.yview()[0]
        self._rueda(self.txt, -120, 5)
        self.assertEqual(self.canvas.yview()[0], antes)

    def test_los_deltas_chicos_del_touchpad_suman_una_muesca(self):
        # El bug del "pegado": int(-40/120) era 0, asi que un touchpad
        # de precision no movia nada hasta un gesto grande.
        self._reiniciar()
        self._rueda(self.canvas, -120, 1)
        una_muesca = self.canvas.yview()[0]
        self._reiniciar()
        self._rueda(self.canvas, -40, 3)
        self.assertEqual(self.canvas.yview()[0], una_muesca)

    def test_el_registro_no_se_desplaza_dos_veces(self):
        # Si la atadura de clase de Text se sumara a la nuestra, una
        # muesca moveria el doble de lo pedido.
        self._reiniciar()
        self.txt.yview_scroll(self.app.LINEAS_POR_MUESCA, "units")
        self.root.update()
        esperado = self.txt.index("@0,0")
        self._reiniciar()
        self._rueda(self.txt, -120, 1)
        self.assertEqual(self.txt.index("@0,0"), esperado)


if __name__ == "__main__":
    unittest.main()
