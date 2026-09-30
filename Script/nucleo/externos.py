# -*- coding: utf-8 -*-
"""
Los paquetes hermanos de Script/, en un solo lugar.

Cada etapa que necesita uno los toma de aca (`from .externos import
extrae_cmg`) en vez de repetir el try/except: asi el fallback de import
esta escrito una sola vez.

El try/except tolera las dos formas de llegar: como parte del paquete
Script (lo normal, desde Balance_BESS.py) o con Script/ en el sys.path.
"""

try:
    from ..Cmg import Extrae_CMG_barras as extrae_cmg
    from ..Medidas import Homologacion, Descarga_PRMTE, Claves_Balance
    from ..Medidas import Generacion_Real
    from ..Medidas.comun import ErrorMedidas
    from ..Subastas import Ofertas_Adjudicadas as ofertas_adj
    from ..Subastas import Fma as fma_subastas
    from ..Fd import Indicadores_DCO as indicadores_dco
    from ..Fd import Indices_FMA as indices_fma
    from ..Fd import Desempeno_Horario as desempeno_fd
except ImportError:  # pragma: no cover - depende de como se importe
    from Cmg import Extrae_CMG_barras as extrae_cmg
    from Medidas import Homologacion, Descarga_PRMTE, Claves_Balance
    from Medidas import Generacion_Real
    from Medidas.comun import ErrorMedidas
    from Subastas import Ofertas_Adjudicadas as ofertas_adj
    from Subastas import Fma as fma_subastas
    from Fd import Indicadores_DCO as indicadores_dco
    from Fd import Indices_FMA as indices_fma
    from Fd import Desempeno_Horario as desempeno_fd

__all__ = [
    "extrae_cmg", "Homologacion", "Descarga_PRMTE", "Claves_Balance",
    "Generacion_Real", "ErrorMedidas", "ofertas_adj", "fma_subastas",
    "indicadores_dco", "indices_fma", "desempeno_fd",
]
