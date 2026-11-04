"""
Las pruebas nunca tocan la configuracion real: __config__ (config.json,
tks.json y lo de cada mes) va a una carpeta temporal por corrida (ver
Script/config.py, BESS_CARPETA_CONFIG).
"""

import atexit
import os
import shutil
import tempfile

if not os.environ.get("BESS_CARPETA_CONFIG"):
    _CARPETA = tempfile.mkdtemp(prefix="bess_config_pruebas_")
    os.environ["BESS_CARPETA_CONFIG"] = _CARPETA
    atexit.register(shutil.rmtree, _CARPETA, ignore_errors=True)
