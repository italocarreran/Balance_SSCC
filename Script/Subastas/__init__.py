# -*- coding: utf-8 -*-
"""
Script.Subastas — la etapa Subastas leida desde su origen real.

Hasta ahora las subastas se sacaban de la hoja "DB" de la planilla 3
(3_REMUNERACIÓN_SUBASTAS_E_ID_*), pero esa planilla no es el origen:
ella misma se arma pegando la salida de entradas_sscc.py, que lee los
Access OfertasSSCCAdj*.accdb. Este paquete corta ese intermediario.
"""
