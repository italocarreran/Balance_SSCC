# -*- coding: utf-8 -*-
"""
Paquete de los scripts de calculo de Balance BESS.

Balance_BESS.py (la ventana) vive un nivel mas arriba y es lo unico
que se ejecuta: todo lo que hay aca adentro es importable, no se corre
suelto.

    Balance_BESS.py
    Script/
        nucleo/            <- el calculo del caso, un modulo por etapa
        Cmg/
            Extrae_CMG_barras.py   <- arma cmg.xlsx desde el CSV
                                      15-minutal
        Subastas/
            Ofertas_Adjudicadas.py <- trae y lee los Access
                                      OfertasSSCCAdj*.accdb

`nucleo` era un solo archivo grande; hoy es un paquete con un modulo
por etapa (ver Script/nucleo/__init__.py, que es solo la fachada:
`nucleo.lo_que_sea` sigue funcionando igual que antes).
"""
