# -*- coding: utf-8 -*-
"""
Todo lo que arma Medidas_SAE.xlsx, la primera entrada del caso.

Reemplaza a los cuatro scripts sueltos que se corrian a mano, uno
detras de otro, en la carpeta donde estuvieran los archivos (autor
original: Freddy.Arriagada):

    0_diccionario_prmte_a_claves_balance.py  -> Homologacion.py
    1_generacion_prmte.py                    -> Descarga_PRMTE.py
    2_generacion_claves_Balance.py           -> Claves_Balance.py
    3_Generacion_Real.py                     -> Generacion_Real.py

Ahora los cuatro pasos corren de un viaje desde el boton "Actualizar"
de la fila Medidas_SAE.xlsx (ver nucleo.generar_medidas_sae), con dos
cambios de fondo pedidos por el usuario:

  - el Excel de homologacion vive en Auxiliares/, al lado de
    Centrales.xlsx, en vez de al lado del .py;
  - la lista de centrales del paso 3 sale de la hoja "Gen real" de ese
    mismo Excel de homologacion, en vez de estar escrita en el codigo.

Ninguno de estos modulos importa `nucleo` (misma regla que Cmg/):
reciben rutas y datos, y levantan ErrorMedidas, que nucleo traduce a
ErrorEntrada. Los intermedios (lotes descargados, marca de reanudacion,
diagnosticos) van a una carpeta de trabajo aparte que la ventana no
muestra.
"""

from .comun import ErrorMedidas

__all__ = ["ErrorMedidas"]
