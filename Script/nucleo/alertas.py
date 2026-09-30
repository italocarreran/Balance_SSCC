# -*- coding: utf-8 -*-
"""
El registro de alertas de una corrida y su estado final.

Viene del catalogo de controles (docs/Alertas_y_Controles_Traspaso_
Python_BESS.md). La idea de fondo: un error no puede volverse 0, vacio
o una fila descartada sin dejar rastro, y el rastro tiene que
sobrevivir al cierre del programa -- hasta ahora los [AVISO] salian
solo a la caja de texto de la ventana.

COMO SE USA (sin tocar ninguna firma)
-------------------------------------
Todo el nucleo ya recibe `registrar` para escribir al log. Un Registro
ES un `registrar`: se puede llamar como funcion igual que print. La
diferencia es que ademas sabe guardar alertas con estructura. El
codigo de calculo hace:

    anotar(registrar, Alerta(...))

y eso funciona con los dos: con un Registro guarda la alerta entera y
escribe la linea; con un print (o una lista, o lo que sea) escribe la
linea nomas. Asi las pruebas y los scripts sueltos siguen andando sin
armar un Registro.
"""

import datetime

import pandas as pd


# ============================================================
# SEVERIDAD
#
# El orden importa: decide el estado final de la corrida.
# ============================================================

CRITICA = "CRITICA"
ALTA = "ALTA"
MEDIA = "MEDIA"
INFO = "INFO"

SEVERIDADES = (CRITICA, ALTA, MEDIA, INFO)

APROBADA = "APROBADA"
APROBADA_CON_ADVERTENCIAS = "APROBADA CON ADVERTENCIAS"
NO_APROBADA = "NO APROBADA - REQUIERE REVISION"
FALLIDA = "FALLIDA - ERROR BLOQUEANTE"

# Columnas de la hoja "Alertas", en orden.
COLUMNAS_ALERTA = (
    "id_alerta", "severidad", "etapa", "central", "clave", "mensaje",
    "valor_encontrado", "valor_esperado", "accion", "archivo", "hoja",
    "origen_control",
)


class Alerta:
    """Una alerta concreta. Los campos que no aplican quedan vacios."""

    __slots__ = COLUMNAS_ALERTA

    def __init__(
        self, id_alerta, severidad, etapa, mensaje, central="", clave="",
        valor_encontrado="", valor_esperado="", accion="", archivo="",
        hoja="", origen_control="",
    ):
        if severidad not in SEVERIDADES:
            raise ValueError(f"Severidad desconocida: {severidad!r}")

        self.id_alerta = id_alerta
        self.severidad = severidad
        self.etapa = etapa
        self.central = central
        self.clave = clave
        self.mensaje = mensaje
        self.valor_encontrado = valor_encontrado
        self.valor_esperado = valor_esperado
        self.accion = accion
        self.archivo = archivo
        self.hoja = hoja
        self.origen_control = origen_control

    def como_fila(self):
        return {campo: getattr(self, campo) for campo in COLUMNAS_ALERTA}

    def linea(self):
        """La linea de una sola alerta para el log de pantalla."""

        partes = [f"  [{self.severidad}] {self.id_alerta}: {self.mensaje}"]
        if self.central:
            partes.append(f" (central: {self.central})")
        return "".join(partes)


class Registro:
    """
    Junta las alertas de una corrida y dice como termino.

    Es callable: `registro("texto")` escribe al log de pantalla, igual
    que print, y ademas deja el texto guardado para el resumen.
    """

    def __init__(self, salida=print, etapa=""):
        self.salida = salida
        self.etapa = etapa
        self.alertas = []
        self.lineas = []
        self.error_bloqueante = ""
        self.inicio = datetime.datetime.now()

    # -- como registrar ------------------------------------------

    def __call__(self, texto):
        self.lineas.append(str(texto))
        if self.salida is not None:
            self.salida(texto)

    # -- como registro -------------------------------------------

    def anotar(self, alerta):
        """Guarda la alerta. No escribe nada: de eso se encarga anotar()."""

        self.alertas.append(alerta)
        return alerta

    def cuantas(self, severidad):
        return sum(1 for a in self.alertas if a.severidad == severidad)

    def conteo(self):
        return {s: self.cuantas(s) for s in SEVERIDADES}

    def estado(self):
        """
        El estado final, segun las reglas de aprobacion del catalogo
        (seccion 21): APROBADA solo si no hay CRITICAS ni ALTAS.
        """

        if self.error_bloqueante:
            return FALLIDA
        if self.cuantas(CRITICA):
            return NO_APROBADA
        if self.cuantas(ALTA):
            return NO_APROBADA
        if self.cuantas(MEDIA):
            return APROBADA_CON_ADVERTENCIAS

        return APROBADA

    def aprobada(self):
        return self.estado() == APROBADA

    # -- salidas --------------------------------------------------

    def tabla(self):
        """Las alertas como DataFrame, en el orden en que aparecieron."""

        if not self.alertas:
            return pd.DataFrame(columns=list(COLUMNAS_ALERTA))

        return pd.DataFrame(
            [a.como_fila() for a in self.alertas],
            columns=list(COLUMNAS_ALERTA),
        )

    def resumen(self, periodo="", hojas_regeneradas=()):
        """
        El bloque de texto que va a la hoja "Ejecucion" y al final del
        log de pantalla.
        """

        conteo = self.conteo()
        lineas = [
            f"EJECUCION BALANCE BESS {periodo}".strip(),
            "=" * 40,
            "",
            f"Estado: {self.estado()}",
            f"Corrida: {self.inicio:%Y-%m-%d %H:%M:%S}",
        ]

        if hojas_regeneradas:
            lineas.append(
                f"Hojas recalculadas en esta corrida: "
                f"{', '.join(hojas_regeneradas)}"
            )
            lineas.append(
                "  (las demas hojas del libro vienen de una corrida "
                "anterior: este estado NO habla de ellas)"
            )

        if self.error_bloqueante:
            lineas += ["", f"Error bloqueante: {self.error_bloqueante}"]

        lineas += ["", "Alertas:"]
        lineas += [f"  {s:8s}: {conteo[s]:,}" for s in SEVERIDADES]

        if self.alertas:
            lineas += ["", "Detalle en la hoja 'Alertas'. Principales:"]
            for severidad in (CRITICA, ALTA, MEDIA):
                for alerta in self.alertas:
                    if alerta.severidad == severidad:
                        lineas.append(f"  - {alerta.id_alerta}: {alerta.mensaje}")
                if len(lineas) > 60:
                    lineas.append("  - (...)")
                    break

        return "\n".join(lineas)


def anotar(registrar, alerta):
    """
    Escribe la alerta en el log y, si `registrar` sabe guardarla, la
    guarda. Funciona con print, con una lista.append y con un Registro.
    """

    if hasattr(registrar, "anotar"):
        registrar.anotar(alerta)

    registrar(alerta.linea())

    return alerta


def anotar_muchas(registrar, alertas, linea_resumen):
    """
    Guarda muchas alertas del mismo tipo y escribe UNA sola linea.

    Es el caso de los cruces que fallan: el detalle completo (una fila
    por central o por clave) va al registro, y a la pantalla va el
    resumen. Sin esto habria que elegir entre inundar el log o perder
    casos, que es justo lo que el catalogo (FD-005) prohibe.
    """

    if hasattr(registrar, "anotar"):
        for alerta in alertas:
            registrar.anotar(alerta)

    if linea_resumen:
        registrar(linea_resumen)

    return list(alertas)
