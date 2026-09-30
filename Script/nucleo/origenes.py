# -*- coding: utf-8 -*-
"""
De donde sale cada cosa que el programa trae de afuera del caso.

Dos usos, los dos de la ventana:

  - el detalle de una fila con boton "Traer" dice "Origen: DCO", y ese
    "DCO" es un link a la carpeta exacta de la que se copia;
  - una fila que NO se trae pero se arma con insumos que si vienen de
    afuera (las tres de FMA, cmg.xlsx) dice "Origen inputs: ..." con el
    mismo tipo de link.

La ruta se resuelve RECIEN AL HACER CLICK, no al pintar el diagrama:
resolverla implica mirar el servidor (que version esta publicada, si
existe la carpeta del mes) y eso, con la unidad de red desconectada,
puede tardar segundos. El diagrama se repinta en cada revisada, asi que
ahi solo viaja la etiqueta.
"""

from .externos import (
    extrae_cmg, indicadores_dco, indices_fma, ofertas_adj,
)


# id de origen -> (titulo del detalle, etiqueta que queda como link,
#                  funcion que resuelve la ruta con el periodo)
ORIGENES = {
    "fd": (
        "Origen",
        "DCO",
        lambda aamm: indicadores_dco.ruta_origen_fd(aamm),
    ),
    "cmg_csv": (
        "Origen",
        "CMg Reales",
        lambda aamm: extrae_cmg.carpeta_origen_csv(aamm),
    ),
    "cmg_xlsx": (
        "Origen inputs",
        "CSV 15-minutal de CMg Reales",
        lambda aamm: extrae_cmg.carpeta_origen_csv(aamm),
    ),
    "subastas": (
        "Origen",
        "progdiar_adjudicaSEN",
        lambda aamm: ofertas_adj.ruta_origen(),
    ),
    "fma_cpf": (
        "Origen inputs",
        "reportes diarios de CPF del DCO",
        lambda aamm: indices_fma.ruta_origen_cpf(aamm),
    ),
    "fma_csf": (
        "Origen inputs",
        "reportes del AGC",
        lambda aamm: indices_fma.ruta_origen_csf(aamm),
    ),
    "fma_ctf": (
        "Origen inputs",
        "CTF_AAMM.csv del DCO",
        lambda aamm: indices_fma.ruta_origen_ctf(aamm),
    ),
}


def origen(id_origen):
    """
    Lo que va en la fila del diagrama: el id (para poder resolver la
    ruta despues), el titulo y la etiqueta que queda como link. None si
    el id no es de los conocidos.
    """

    if id_origen not in ORIGENES:
        return None

    titulo, etiqueta, _ = ORIGENES[id_origen]

    return {"id": id_origen, "titulo": titulo, "etiqueta": etiqueta}


def ruta_origen(id_origen, aamm=None):
    """
    La carpeta de la que sale ese origen para el periodo dado. Se llama
    al hacer click en el link, no al pintar el diagrama. None si el id
    no es de los conocidos.
    """

    if id_origen not in ORIGENES:
        return None

    return ORIGENES[id_origen][2](aamm)
