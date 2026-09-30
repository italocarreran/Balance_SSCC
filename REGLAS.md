# REGLAS.md — checklist obligatorio de inicio/cierre de sesión

Se lee **entero, primero**, antes que nada. Ver `METODOLOGIA.md` para el
resto (regla de expansión de contexto, convenciones, trampas conocidas).

---

## Al empezar

1. `git log` reciente — ver de quién es el último commit y leerlo completo
   si no es propio.
2. Leer `BITACORA.md` entero, empezando por "Pendientes abiertos".
3. Sincronizar la rama de trabajo con la rama principal antes del primer
   cambio.

## Mientras se trabaja

4. No dejar un cambio a medias — si toca varios archivos, terminarlo y
   verificarlo en la misma sesión, o dejar explícito en `BITACORA.md` qué
   falta y por qué se cortó.
5. Antes de dar un cambio por terminado: correr
   `python -m py_compile Balance_BESS.py Script/nucleo/*.py
   Script/*/*.py` **y** `python -m unittest discover` desde la raíz
   (revisar que el total de pruebas no baje: si dice `Ran 0 tests`,
   algo se rompió en el descubrimiento, no es un verde). Las pruebas de
   la ventana (`tests/test_ventana_rueda.py`) se saltean si no hay
   tkinter o no hay pantalla: en Linux corren con
   `xvfb-run -a python -m unittest discover`. Si hay un caso de
   prueba disponible (una carpeta con la estructura de
   `docs/Plan_Traspaso_Python_Balance_BESS.md` §3.3), ejecutar
   `nucleo.ejecutar(carpeta)` sobre ella y revisar la hoja `Log` de
   `Consolidado_entradas.xlsx`. No existe todavía un script de verificación
   automatizado (`scripts/verificar.sh`); si se crea uno, esta regla debe
   apuntar a él.
6. Nunca reescribir el historial compartido (`push --force`,
   `commit --amend` sobre algo ya subido, `reset --hard` contra la rama
   remota).

## Al cerrar, si se cambió algo

7. Agregar una entrada en `BITACORA.md`.
8. Actualizar `MAPA.md` y/o `METODOLOGIA.md` si cambió algo estructural o se
   tomó una decisión de diseño (p. ej. se confirmó la lógica de una columna
   pendiente, o se definió la ubicación de OfertasSSCC).
9. Commitear y pushear — un cambio que queda solo en el working tree de la
   sesión no existe para nadie más.
