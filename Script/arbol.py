# -*- coding: utf-8 -*-
"""
Prefijos del arbol tipo consola (├── / └── / │) que dibujan las
ventanas (Balance_BESS.py, Balance_CRA.py).

Es presentacion pura, sin tkinter: recibe la lista plana de niveles que
devuelve revisar_estructura() (0 = raiz del caso) y devuelve el prefijo
de cada fila. Salio de Balance_BESS.py para que las dos ventanas usen
la misma.
"""


def _es_ultimo_en_su_nivel(niveles, i):
    """
    True si, mirando hacia adelante desde i, se sube de nivel antes
    de encontrar otra fila con el MISMO nivel (o se llega al final de
    la lista): o sea, si i es la ultima de su grupo.
    """
    nivel = niveles[i]
    for j in range(i + 1, len(niveles)):
        if niveles[j] < nivel:
            return True
        if niveles[j] == nivel:
            return False
    return True


def _prefijos_arbol(niveles):
    """
    Prefijos tipo consola (├── / └── / │) para una lista plana de
    niveles (0 = raiz), calculando el relleno de cada ancestro segun
    si ESE ancestro es o no el ultimo de su propio grupo.
    """
    n = len(niveles)
    ultimos = [_es_ultimo_en_su_nivel(niveles, i) for i in range(n)]
    prefijos = []

    for i, nivel in enumerate(niveles):

        if nivel == 0:
            prefijos.append("")
            continue

        relleno = ""
        for ancestro_nivel in range(1, nivel):
            indice_ancestro = next(
                (k for k in range(i - 1, -1, -1) if niveles[k] == ancestro_nivel),
                None,
            )
            relleno += (
                "    "
                if (indice_ancestro is None or ultimos[indice_ancestro])
                else "│   "
            )

        prefijos.append(relleno + ("└── " if ultimos[i] else "├── "))

    return prefijos
