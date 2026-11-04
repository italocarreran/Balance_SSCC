# -*- coding: utf-8 -*-
"""
Buscar una central con TODOS sus nombres del Diccionario.

Pedido del usuario: "para subastas, ofertas y todo, que se busque con
todos los nombres del diccionario [...] y que se indique tipo 'central
conejo no coincide con nombre de subastas del diccionario, pero si se
encuentra por el nombre de FD'".

Cada cruce sigue probando PRIMERO el nombre que le corresponde (la
columna FD para el FD, FMA_CPF para el FMA CPF, Ofertas para las
ofertas...). Solo si ese no esta en la tabla de destino se prueban los
demas nombres de la misma fila del Diccionario, y si alguno cruza se
usa y se avisa (DIC-002), una vez por central y cruce. Un Diccionario
en el formato viejo (bloques) no tiene filas: todo queda como antes.
"""

from .lectura import (
    ROL_BALANCE_BESS, ROL_FD, ROL_FMA_CPF, ROL_OFERTAS, ROL_SUBASTAS,
    filas_diccionario,
)
from .utiles import _texto_seguro, _tiene_valor, normalizar


ETIQUETA_ROL = {
    ROL_BALANCE_BESS: "Balance_BESS",
    ROL_SUBASTAS: "Subastas",
    ROL_FD: "FD",
    ROL_OFERTAS: "Ofertas",
    ROL_FMA_CPF: "FMA_CPF",
}
# El orden en que se prueban los otros nombres.
ORDEN_ROLES = tuple(ETIQUETA_ROL)

ID_AVISO = "DIC-002"


class Equivalencias:
    """Las filas del Diccionario, buscables por cualquiera de sus nombres."""

    def __init__(self, diccionario, normalizar_fn=normalizar):
        self._norm = normalizar_fn
        self.filas = (
            filas_diccionario(diccionario) if diccionario is not None else []
        )
        self._fila_de = {}
        for fila in self.filas:
            for rol in ORDEN_ROLES:
                clave = self._norm(fila.get(rol))
                if clave:
                    self._fila_de.setdefault(clave, fila)
        self.avisos = []

    def __bool__(self):
        return bool(self.filas)

    def fila(self, nombre):
        return self._fila_de.get(self._norm(nombre))

    def buscar(self, nombre, rol_destino, disponibles, donde):
        """
        El nombre con el que 'nombre' (de una central, en cualquier
        nomenclatura) esta en el destino, o None.

        disponibles: los nombres del destino, ya normalizados con la
        misma funcion que esta instancia. Primero el de rol_destino;
        despues los demas de su fila (con aviso); despues el nombre tal
        cual (como hacia cada cruce antes).
        """

        fila = self.fila(nombre)
        if fila is None:
            return nombre if self._norm(nombre) in disponibles else None

        esperado = fila.get(rol_destino)
        if _tiene_valor(esperado) and self._norm(esperado) in disponibles:
            return esperado

        for rol in ORDEN_ROLES:
            if rol == rol_destino:
                continue
            otro = fila.get(rol)
            if _tiene_valor(otro) and self._norm(otro) in disponibles:
                self._avisar(fila, rol_destino, rol, otro, donde)
                return otro

        return None

    def central_de(self, nombre, rol_origen, donde):
        """
        El nombre Balance_BESS de 'nombre', que viene de una tabla en
        nomenclatura rol_origen (una unidad de la hoja FD, un nombre de
        Ofertas). Si 'nombre' no esta en esa columna del Diccionario
        pero si en otra de la misma fila, se usa esa fila y se avisa.
        None si no esta en ninguna.
        """

        fila = self.fila(nombre)
        if fila is None:
            return None

        central = _texto_seguro(fila.get(ROL_BALANCE_BESS))
        if self._norm(fila.get(rol_origen)) == self._norm(nombre):
            return central or None

        rol = next(
            r for r in ORDEN_ROLES
            if self._norm(fila.get(r)) == self._norm(nombre)
        )
        self._avisar(fila, rol_origen, rol, nombre, donde)
        return central or None

    def _avisar(self, fila, rol_esperado, rol_usado, valor, donde):
        central = _texto_seguro(fila.get(ROL_BALANCE_BESS)) or "?"
        esperado = _texto_seguro(fila.get(rol_esperado))
        etiqueta = ETIQUETA_ROL.get(rol_esperado, rol_esperado)
        primera = (
            f"su nombre {etiqueta} del Diccionario ('{esperado}') no "
            f"coincide con {donde}"
            if esperado else
            f"no tiene nombre {etiqueta} en el Diccionario"
        )
        texto = (
            f"central '{central}': {primera}, pero si se encuentra por su "
            f"nombre {ETIQUETA_ROL.get(rol_usado, rol_usado)} ('{valor}')."
        )
        if texto not in self.avisos:
            self.avisos.append(texto)

    def informar(self, registrar):
        """Pasa los avisos al log (una linea por central y cruce)."""

        for texto in self.avisos:
            registrar(f"  [AVISO] {ID_AVISO}: {texto}")
        self.avisos = []
