# Trazabilidad FD → columna `DB!Y` de Planilla 3

## 1. Objetivo

Este documento describe cómo generar directamente la columna:

```text
DB!Y = FD
```

de la Planilla 3:

```text
3_REMUNERACIÓN_SUBASTAS_E_ID_AAMM_Definitivo.xlsm
```

a partir del archivo de desempeño mensual:

```text
SSCC_Desempeño_<Mes>_AAAA_V2.xlsx
```

Para julio de 2026 se revisó:

```text
SSCC_Desempeño_Julio_2026_V2.xlsx
```

La idea es eliminar la dependencia de que alguien copie previamente estos datos a la Planilla 3.

---

# 2. Resumen ejecutivo

El archivo de desempeño contiene tres hojas horarias relevantes:

```text
CPF Horario
CSF Horario
CTF Horario
```

La Planilla 3 copia esas tablas respectivamente a:

```text
CPF Horario → CPF_FD
CSF Horario → CSF_FD
CTF Horario → CTF_FD
```

Luego `DB!Y` selecciona el FD correspondiente según `DB!C`:

```text
CPF → FD_CPF
CSF → FD_CSF
CTF → FD_CTF
```

La búsqueda se hace mediante la clave:

```text
Unidad Infotécnica + Hora_Mes
```

Es decir:

```text
DB!N + DB!J
```

Por lo tanto, el flujo objetivo puede reducirse a:

```text
SSCC_Desempeño_AAAA_V2.xlsx
        │
        ├── CPF Horario ──> FD_CPF
        ├── CSF Horario ──> FD_CSF
        └── CTF Horario ──> FD_CTF
                │
                ▼
      normalización de nombre de unidad
                +
           cálculo Hora_Mes
                │
                ▼
       Unidad + Hora_Mes
                │
                ▼
             DB!Y
```

---

# 3. Fórmula actual de `DB!Y`

La fórmula observada en la Planilla 3 es conceptualmente:

```excel
=SI(
   C="CPF",
   buscar CPF,
   SI(
      C="CSF",
      buscar CSF,
      SI(
         C="CTF",
         buscar CTF,
         "ERROR"
      )
   )
)
```

La fórmula real utiliza:

```text
DB!C = familia del servicio: CPF / CSF / CTF
DB!N = Unidad INFOTECNICA
DB!J = Hora_Mes
```

y busca:

```text
DB!N & DB!J
```

en las respectivas hojas FD.

---

# 4. Entrada original

## Archivo

```text
SSCC_Desempeño_<Mes>_AAAA_V2.xlsx
```

Ejemplo:

```text
SSCC_Desempeño_Julio_2026_V2.xlsx
```

## Hojas necesarias

Solamente son necesarias para `DB!Y`:

```text
CPF Horario
CSF Horario
CTF Horario
```

Las hojas mensuales:

```text
CPF Mensual
CSF Mensual
CTF Mensual
```

no intervienen en la fórmula actual de `DB!Y`.

---

# 5. Entrada CPF

## 5.1. Hoja de origen

```text
CPF Horario
```

La tabla comienza con los encabezados en la fila 11.

Las columnas relevantes son:

| Columna origen | Campo |
|---|---|
| B | Fecha |
| C | Hora |
| D | Unidad |
| E | Respuesta CPF+ (`Fact_CPF+`) |
| F | Respuesta CPF- (`Fact_CPF-`) |
| G | Disponibilidad (`Fdis_CPF`) |
| H | Desempeño (`DCPF`) |
| I | Factor de Desempeño (`Fd_CPF`) |
| J | Cuenta con equipo registrador validado |

Para generar `DB!Y`, el valor principal que finalmente interesa es:

```text
CPF Horario!I = Fd_CPF
```

---

# 6. Cómo se pega CPF en Planilla 3

La correspondencia observada es:

| Archivo `CPF Horario` | Planilla 3 `CPF_FD` |
|---|---|
| B Fecha | F Fecha |
| C Hora | G Hora |
| D Unidad | H Unidad |
| E Respuesta CPF+ | I Respuesta CPF+ |
| F Respuesta CPF- | J Respuesta CPF- |
| G Disponibilidad | K Disponibilidad |
| H Desempeño | L Desempeño |
| I `Fd_CPF` | M `Factor de Desempeño (Fd_CPF)` |
| J Equipo registrador | N Equipo registrador |

Por tanto:

```text
CPF Horario!I
        ↓
CPF_FD!M
        ↓
DB!Y para las filas CPF
```

No es necesario copiar físicamente la tabla a Planilla 3 si el nuevo código lee directamente `CPF Horario`.

---

# 7. Entrada CSF

## 7.1. Hoja de origen

```text
CSF Horario
```

Encabezados en fila 11.

| Columna origen | Campo |
|---|---|
| B | Fecha |
| C | Hora |
| D | Unidad |
| E | Respuesta CSF (`Fact_CSF`) |
| F | Disponibilidad (`Fdis_CSF`) |
| G | Desempeño (`DCSF`) |
| H | Factor de Desempeño (`Fd_CSF`) |

Para `DB!Y`, interesa:

```text
CSF Horario!H = Fd_CSF
```

---

# 8. Cómo se pega CSF en Planilla 3

La correspondencia es:

| Archivo `CSF Horario` | Planilla 3 `CSF_FD` |
|---|---|
| B Fecha | F Fecha |
| C Hora | G Hora |
| D Unidad | H Unidad Infotecnica |
| E Respuesta CSF | I Respuesta CSF |
| F Disponibilidad | J Disponibilidad horaria |
| G Desempeño | K Desempeño horario |
| H `Fd_CSF` | L `FD_CSF` |

Por tanto:

```text
CSF Horario!H
       ↓
CSF_FD!L
       ↓
DB!Y para las filas CSF
```

---

# 9. Entrada CTF

## 9.1. Hoja de origen

```text
CTF Horario
```

Encabezados en fila 11.

| Columna origen | Campo |
|---|---|
| B | Fecha |
| C | Hora |
| D | Unidad |
| E | Respuesta CTF+ (`Fact_CTF+`) |
| F | Respuesta CTF- (`Fact_CTF-`) |
| G | Disponibilidad (`Fdis_CTF`) |
| H | Desempeño (`DCTF`) |
| I | Factor de Desempeño (`Fd_CTF`) |

Para `DB!Y`, interesa:

```text
CTF Horario!I = Fd_CTF
```

---

# 10. Cómo se pega CTF en Planilla 3

La correspondencia es:

| Archivo `CTF Horario` | Planilla 3 `CTF_FD` |
|---|---|
| B Fecha | F Fecha |
| C Hora | G Hora |
| D Unidad | H Unidad Infotecnica |
| E Respuesta CTF+ | I Respuesta CTF+ |
| F Respuesta CTF- | J Respuesta CTF- |
| G Disponibilidad | K Disponibilidad horaria |
| H Desempeño | L Desempeño horario |
| I `Fd_CTF` | M `FD_CTF` |

Por tanto:

```text
CTF Horario!I
       ↓
CTF_FD!M
       ↓
DB!Y para las filas CTF
```

---

# 11. Cálculo de `Hora_Mes`

El archivo de desempeño utiliza:

```text
Fecha
Hora
```

donde `Hora` se encuentra en formato:

```text
0 ... 23
```

La Planilla 3 transforma esto a una hora mensual de tipo:

```text
1 ... 744
```

según corresponda al mes.

La fórmula observada en las tres hojas FD es:

```excel
(DIA(Fecha)-1)*24
+ Hora
+ 1
+ SI(DIA(Fecha)>SUBASTAS!F2;1;0)
```

Conceptualmente:

```python
Hora_Mes = (dia - 1) * 24 + hora + 1
```

y existe además un ajuste de cambio de hora:

```python
if dia > dia_cambio_hora:
    Hora_Mes += 1
```

donde el día de cambio de hora se obtiene actualmente desde:

```text
SUBASTAS!F2
```

## Importante

Para reproducir exactamente Planilla 3, este ajuste debe conservarse y el valor equivalente a:

```text
SUBASTAS!F2
```

debe estar disponible para el nuevo código.

---

# 12. Nomenclatura de unidades

Este es el otro punto fundamental.

`DB!Y` NO busca utilizando directamente la `Configuración` de subastas.

Utiliza:

```text
DB!N = Unidad INFOTECNICA
```

La entrada de desempeño también posee una columna:

```text
Unidad
```

pero Planilla 3 utiliza diccionarios para resolver diferencias de nomenclatura.

Por tanto, para desacoplarse de Planilla 3 debe existir un diccionario que permita relacionar:

```text
Unidad de subastas
        ↓
Unidad INFOTECNICA
        ↓
Unidad utilizada por SSCC_Desempeño
```

Este diccionario puede mantenerse en:

```text
Centrales.xlsx
```

junto con los otros diccionarios ya definidos para el nuevo proceso.

---

# 13. Clave principal de búsqueda

Para las tres familias, la clave final es:

```text
Unidad_INFOTECNICA + Hora_Mes
```

Ejemplo conceptual:

```text
HE RALCO U2 + 2
```

se transforma en:

```text
HE RALCO U22
```

La Planilla 3 realiza una concatenación directa, sin separador.

En Python **no es recomendable concatenar textos**.

Es mejor hacer el `merge` por dos columnas:

```python
[
    "Unidad_INFOTECNICA",
    "Hora_Mes"
]
```

Esto evita colisiones y hace la lógica más clara.

---

# 14. CPF: comportamiento especial de nomenclatura

En `CPF_FD` existen dos claves auxiliares:

```text
B = Unidad normalizada + Hora_Mes
A = Unidad alternativa TG/TV + Hora_Mes
```

La columna A intercambia los sufijos:

```text
TG ↔ TV
```

La fórmula observada es conceptualmente:

```python
if unidad termina en "TG":
    unidad_alternativa = unidad sin "TG" + "TV"

elif unidad termina en "TV":
    unidad_alternativa = unidad sin "TV" + "TG"

else:
    unidad_alternativa = unidad
```

`DB!Y` primero intenta la clave normal y, si falla, prueba la alternativa.

La fórmula equivalente es:

```text
buscar Unidad + Hora_Mes
si no existe:
    buscar Unidad_alternativa + Hora_Mes
```

y devuelve siempre:

```text
FD_CPF
```

---

# 15. Búsqueda CPF realizada por DB

La lógica actual es:

```excel
SI(
  C="CPF",
  SI.ERROR(
      BUSCARV(N&J; CPF_FD!B:R; 12; 0),
      SI.ERROR(
          BUSCARV(N&J; CPF_FD!A:R; 13; 0),
          "ERRORCPF"
      )
  )
)
```

Ambas búsquedas devuelven:

```text
CPF_FD!M = Fd_CPF
```

Por tanto, en Python:

```python
fd = buscar(unidad_infotecnia, hora_mes)

if no existe:
    fd = buscar(unidad_alternativa_tg_tv, hora_mes)
```

---

# 16. Búsqueda CSF realizada por DB

La fórmula actual utiliza:

```excel
BUSCARV(
    N & J;
    CSF_FD!B:O;
    11;
    0
)
```

La columna 11 contando desde B es:

```text
L = FD_CSF
```

Por tanto:

```python
FD = buscar(
    Unidad_INFOTECNICA,
    Hora_Mes
)["FD_CSF"]
```

Si no encuentra:

```text
ERRORCSF
```

---

# 17. Búsqueda CTF realizada por DB

La fórmula actual utiliza:

```excel
BUSCARV(
    N & J;
    CTF_FD!B:Q;
    12;
    0
)
```

La columna 12 contando desde B es:

```text
M = FD_CTF
```

Por tanto:

```python
FD = buscar(
    Unidad_INFOTECNICA,
    Hora_Mes
)["FD_CTF"]
```

Si no encuentra:

```text
ERRORCTF
```

---

# 18. Lógica completa de `DB!Y`

En pseudocódigo:

```python
if Control == "CPF":

    fd = buscar_cpf(
        Unidad_INFOTECNICA,
        Hora_Mes
    )

    if no_encontrado:
        fd = buscar_cpf(
            Unidad_alternativa_TG_TV,
            Hora_Mes
        )

elif Control == "CSF":

    fd = buscar_csf(
        Unidad_INFOTECNICA,
        Hora_Mes
    )

elif Control == "CTF":

    fd = buscar_ctf(
        Unidad_INFOTECNICA,
        Hora_Mes
    )

else:

    fd = "ERROR"
```

---

# 19. Diseño recomendado para Python

## 19.1. Leer solamente las tres hojas necesarias

```python
cpf = pd.read_excel(
    archivo_fd,
    sheet_name="CPF Horario",
    header=10,
    usecols="B:J"
)

csf = pd.read_excel(
    archivo_fd,
    sheet_name="CSF Horario",
    header=10,
    usecols="B:H"
)

ctf = pd.read_excel(
    archivo_fd,
    sheet_name="CTF Horario",
    header=10,
    usecols="B:I"
)
```

---

# 20. Preparar CPF

Mantener:

```text
Fecha
Hora
Unidad
Factor de Desempeño (Fd_CPF)
```

y calcular:

```python
cpf["Hora_Mes"] = (
    (cpf["Fecha"].dt.day - 1) * 24
    + cpf["Hora"]
    + 1
)
```

Aplicar luego el ajuste de cambio de hora cuando corresponda.

Renombrar:

```python
cpf = cpf.rename(
    columns={
        "Unidad": "Unidad_INFOTECNICA",
        "Factor de Desempeño\n(Fd_CPF)": "FD"
    }
)
```

Resultado deseado:

```text
Unidad_INFOTECNICA
Hora_Mes
FD
```

---

# 21. Preparar CSF

Mantener:

```text
Fecha
Hora
Unidad
Factor de Desempeño (Fd_CSF)
```

Calcular `Hora_Mes`.

Resultado:

```text
Unidad_INFOTECNICA
Hora_Mes
FD
```

donde:

```text
FD = Fd_CSF
```

---

# 22. Preparar CTF

Mantener:

```text
Fecha
Hora
Unidad
Factor de Desempeño (Fd_CTF)
```

Calcular `Hora_Mes`.

Resultado:

```text
Unidad_INFOTECNICA
Hora_Mes
FD
```

donde:

```text
FD = Fd_CTF
```

---

# 23. Tabla FD normalizada recomendada

Una alternativa más limpia es unir las tres entradas en una sola tabla:

| Control | Unidad_INFOTECNICA | Hora_Mes | FD |
|---|---|---:|---:|
| CPF | HE RALCO U2 | 2 | 1 |
| CSF | HE RALCO U2 | 2 | 0.95 |
| CTF | HE RALCO U2 | 2 | 1 |

Entonces la clave de unión con la DB pasa a ser:

```python
[
    "Control",
    "Unidad_INFOTECNICA",
    "Hora_Mes"
]
```

y `DB!Y` se obtiene mediante un único `merge`.

---

# 24. Unión con la futura DB generada desde subastas

De la reconstrucción anterior de DB ya tenemos:

```text
DB!C → Control
DB!J → Hora_Mes
DB!N → Unidad_INFOTECNICA
```

Por tanto:

```python
db = db.merge(
    fd_normalizado,
    on=[
        "Control",
        "Unidad_INFOTECNICA",
        "Hora_Mes"
    ],
    how="left"
)
```

Luego:

```text
FD → equivalente a DB!Y
```

---

# 25. Tratamiento de `No Participó`, `No Activado`, etc.

Para `DB!Y` NO se recalcula el FD a partir de las respuestas.

El valor final ya viene calculado en el archivo:

```text
SSCC_Desempeño_...xlsx
```

Por ejemplo, la entrada puede contener:

```text
Respuesta = "No Participó"
FD = 1
```

o:

```text
Respuesta = 0.336...
FD = 0
```

El nuevo proceso debe tomar directamente:

```text
Fd_CPF
Fd_CSF
Fd_CTF
```

tal como vienen en el archivo de desempeño.

No debe intentar reconstruir el FD salvo que en el futuro se quiera replicar también el proceso que genera `SSCC_Desempeño`.

---

# 26. Dependencia con la trazabilidad FMA: Vector Participación CSF

La revisión de `CSF_FD` también permite cerrar el punto pendiente de la trazabilidad FMA.

`DB!AC` contiene:

```text
Vector Participación CSF
```

La fórmula actual es:

```excel
SI(
   C="CSF",
   BUSCARV(
      N & J;
      CSF_FD!B:O;
      14;
      0
   ),
   1
)
```

La columna 14 contando desde `B` es:

```text
O = Indicador de participación
```

Por tanto:

```text
DB!AC = CSF_FD!O
```

para CSF.

Para cualquier otro servicio:

```text
DB!AC = 1
```

---

# 27. Cómo se calcula el Indicador de Participación CSF

En `CSF_FD`, la columna:

```text
O = Indicador de participación
```

se obtiene principalmente desde:

```text
I = Respuesta CSF
```

La lógica observada es:

```text
si la unidad figura como "No Participó"
y tampoco existe participación mediante la unidad alternativa TG/TV:
    Indicador = 0
si no:
    Indicador = 1
```

Conceptualmente:

```python
if respuesta_csf == "No Participó":
    verificar_unidad_alternativa_tg_tv()

    if alternativa_tambien_no_participo_o_no_existe:
        indicador_participacion = 0
    else:
        indicador_participacion = 1
else:
    indicador_participacion = 1
```

Por tanto, para cerrar completamente el FMA CSF:

```text
FMA final CSF
=
FMA base CSF
×
Indicador Participación CSF
```

Este indicador puede obtenerse desde la misma entrada:

```text
CSF Horario
```

sin depender de la Planilla 3.

---

# 28. Flujo final conjunto FMA + FD

Con esta trazabilidad el proceso futuro puede quedar:

```text
subastas_AAMM.xlsx
        │
        └── genera estructura base DB
                 │
                 ├── B:K
                 ├── Unidad Infotécnica
                 └── Hora_Mes
                          │
                          ├──────────────────────────┐
                          │                          │
                          ▼                          ▼
                  archivos FMA          SSCC_Desempeño_AAAA_V2.xlsx
                          │                          │
                 FMA CPF/CSF/CTF         CPF/CSF/CTF Horario
                          │                          │
                          │                 ┌────────┴────────┐
                          │                 │                 │
                          │                 ▼                 ▼
                          │                FD            Participación CSF
                          │                 │                 │
                          │                 ▼                 ▼
                          │               DB!Y              DB!AC
                          │                                   │
                          └───────────────────────────────────┘
                                           │
                                           ▼
                                      FMA final DB!V
```

---

# 29. Entradas necesarias hasta este punto

Para reconstruir las columnas ya analizadas de Planilla 3 se necesitan:

## Subastas / DB base

```text
subastas_AAMM.xlsx
```

## FMA

```text
Reportes diarios CPF
csf_AAAAMMDD.xlsx
CTF_AAAAMM.csv
```

o, si se mantienen las salidas actuales de `entradas_sscc.py`:

```text
fma_cpf_AAMM.xlsx
fma_csf_AAMM.xlsx
fma_cft_AAMM.xlsx
```

## FD

```text
SSCC_Desempeño_<Mes>_AAAA_V2.xlsx
```

utilizando:

```text
CPF Horario
CSF Horario
CTF Horario
```

## Diccionarios

Se debe mantener en un archivo auxiliar como `Centrales.xlsx` la traducción de nomenclaturas necesaria entre:

```text
Configuración subasta
Unidad INFOTECNICA
Nomenclatura FD
Nomenclatura FMA CPF
```

---

# 30. Validación recomendada

Al implementar el código se recomienda comparar:

```text
FD_Planilla3 = DB!Y
FD_Python    = nuevo cálculo
```

por:

```text
Control
Unidad_INFOTECNICA
Hora_Mes
```

y generar:

```text
FD_Planilla3
FD_Python
Coincide
```

Regla:

```python
Coincide = FD_Planilla3 == FD_Python
```

Para valores numéricos puede utilizarse una tolerancia pequeña.

Además, validar por separado:

```text
CPF
CSF
CTF
```

y revisar especialmente:

```text
TG ↔ TV
cambio de hora
unidades sin diccionario
duplicados Unidad + Hora_Mes
```

---

# 31. Conclusión

La columna:

```text
DB!Y = FD
```

puede generarse directamente desde:

```text
SSCC_Desempeño_<Mes>_AAAA_V2.xlsx
```

sin esperar que sus datos sean copiados a Planilla 3.

El mapeo esencial es:

```text
CPF Horario!I → FD_CPF → DB!Y
CSF Horario!H → FD_CSF → DB!Y
CTF Horario!I → FD_CTF → DB!Y
```

utilizando como clave:

```text
Control + Unidad_INFOTECNICA + Hora_Mes
```

Adicionalmente, desde `CSF Horario` puede reconstruirse:

```text
DB!AC = Vector Participación CSF
```

lo que permite cerrar también el cálculo pendiente de:

```text
DB!V = FMA
```

para las filas CSF.
