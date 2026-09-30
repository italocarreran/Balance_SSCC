# Trazabilidad: `subastas_AAMM.xlsx` → hoja `DB` de Planilla 3

## 1. Objetivo

Documentar cómo obtener directamente desde el archivo `subastas_AAMM.xlsx` la información que actualmente termina en las columnas **B:K** de la hoja `DB` de la planilla:

`3_REMUNERACIÓN_SUBASTAS_E_ID_AAMM_Definitivo.xlsm`

El objetivo final es que un proceso externo pueda leer `subastas_AAMM.xlsx` y construir estas columnas de `DB` **sin esperar a que se complete la Planilla 3**.

> Por ahora quedan fuera de este documento:
>
> - `DB!L` = Propietario
> - `DB!V` = FMA
> - `DB!Y` = FD
>
> Estas columnas se resolverán por separado.

---

## 2. Flujo actual

Actualmente la información sigue esta cadena:

```text
Archivos OfertasSSCCAdj*.accdb
        ↓
entradas_sscc.py
        ↓
subastas_AAMM.xlsx
        ↓
Hoja SUBASTAS de Planilla 3
        ↓
Hoja DB de Planilla 3
```

Para las columnas `DB!B:K`, la hoja `SUBASTAS` funciona principalmente como una capa intermedia.

La propuesta es reemplazar ese flujo por:

```text
subastas_AAMM.xlsx
        ↓
Transformaciones equivalentes a las fórmulas de SUBASTAS
        ↓
Datos equivalentes a DB!B:K
```

---

## 3. Estructura del archivo `subastas_AAMM.xlsx`

El archivo generado por `entradas_sscc.py` tiene esta estructura:

| Columna Excel | Campo |
|---|---|
| A | Índice exportado por pandas |
| B | CONFIGURACIÓN |
| C | SERVICIO |
| D | AÑO |
| E | MES |
| F | DIA |
| G | HORA |
| H | BANDA |
| I | CANTIDAD MW |
| J | PRECIO USD/MW |
| K | FECHA |
| L | CONTROL |
| M | SUB_BAJ |
| N | HORA_MES |
| O | CANTIDAD PONDERADA MW |
| P | Hora_PID |

En el archivo generado por Python, las columnas `K:N` son creadas inicialmente vacías. En la Planilla 3 esos cuatro campos se vuelven a calcular mediante fórmulas en la hoja `SUBASTAS`.

Por lo tanto, para reproducir `DB!B:K` no se deben depender de los valores existentes en `subastas_AAMM.xlsx!K:N`; se deben reconstruir con la misma lógica de la Planilla 3.

---

## 4. Relación de filas

### Archivo `subastas_AAMM.xlsx`

- Encabezados: fila **1**
- Primer registro: fila **2**

### Hoja `SUBASTAS` de Planilla 3

- Encabezados: fila **8**
- Primer registro: fila **9**

Por lo tanto:

```text
subastas_AAMM fila 2  → SUBASTAS fila 9
subastas_AAMM fila 3  → SUBASTAS fila 10
subastas_AAMM fila 4  → SUBASTAS fila 11
...
```

Es decir:

```text
fila_SUBASTAS = fila_subastas_AAMM + 7
```

### Hoja `DB`

- Encabezados: fila **2**
- Primer registro: fila **3**

La primera fila de datos de DB referencia la primera fila de datos de SUBASTAS:

```text
DB fila 3 → SUBASTAS fila 9
```

Por consiguiente, al ir directo desde `subastas_AAMM.xlsx`:

```text
subastas_AAMM fila 2 → DB fila 3
subastas_AAMM fila 3 → DB fila 4
subastas_AAMM fila 4 → DB fila 5
...
```

O sea:

```text
fila_DB = fila_subastas_AAMM + 1
```

Este desplazamiento solo importa si se reproduce físicamente la estructura de Excel. Si se trabaja con pandas/DataFrames, basta con mantener el mismo orden de registros.

---

# 5. Mapeo final `subastas_AAMM.xlsx` → `DB!B:K`

## Resumen

| DB | Nombre en DB | Origen final en `subastas_AAMM.xlsx` | Transformación |
|---|---|---|---|
| B | Concepto | C = SERVICIO | Directa |
| C | Control | C = SERVICIO | Primeros 3 caracteres |
| D | Sub_Baj | C = SERVICIO | Determinar SUBIDA / BAJADA |
| E | Fecha | D:F = AÑO, MES, DIA | Construir fecha |
| F | Año | D = AÑO | Directa |
| G | Mes | E = MES | Directa |
| H | Dia | F = DIA | Directa |
| I | Hora_dia | G = HORA | Directa |
| J | Hora_mes | F = DIA y G = HORA | Fórmula Hora Mes |
| K | Configuración | B = CONFIGURACIÓN | Directa |

---

## 5.1. `DB!B` — Concepto

### Flujo actual

```excel
DB!B3 = SUBASTAS!C9
```

`SUBASTAS!C` corresponde a `SERVICIO`.

### Origen directo

```text
subastas_AAMM.xlsx!C:C
```

### Regla

```python
DB_B_Concepto = SERVICIO
```

### Ejemplos

```text
CSF(-)
CSF(+)
CPF(-)
CPF(+)
CTF(-)
CTF(+)
```

---

## 5.2. `DB!C` — Control

### Flujo actual

```excel
DB!C3 = SUBASTAS!L9
```

En `SUBASTAS!L9` existe la fórmula:

```excel
=LEFT($C9,3)
```

Por lo tanto, `Control` se deriva de `SERVICIO`.

### Origen directo

```text
subastas_AAMM.xlsx!C:C = SERVICIO
```

### Regla

Tomar los primeros tres caracteres de `SERVICIO`:

```python
Control = SERVICIO[:3]
```

### Ejemplos

```text
CSF(-) → CSF
CSF(+) → CSF
CPF(-) → CPF
CPF(+) → CPF
CTF(-) → CTF
CTF(+) → CTF
```

---

## 5.3. `DB!D` — Sub_Baj

### Flujo actual

```excel
DB!D3 = SUBASTAS!M9
```

En `SUBASTAS!M9` existe la fórmula:

```excel
=IF(LEFT(RIGHT($C9,2),1)="+","SUBIDA","BAJADA")
```

Por lo tanto se deriva de `SERVICIO`.

### Origen directo

```text
subastas_AAMM.xlsx!C:C = SERVICIO
```

### Regla equivalente

```python
Sub_Baj = "SUBIDA" if "+" in SERVICIO else "BAJADA"
```

Para replicar de forma estricta la lógica Excel actual, se puede revisar específicamente el signo ubicado antes del paréntesis final.

### Ejemplos

```text
CSF(+) → SUBIDA
CSF(-) → BAJADA
CPF(+) → SUBIDA
CPF(-) → BAJADA
CTF(+) → SUBIDA
CTF(-) → BAJADA
```

---

## 5.4. `DB!E` — Fecha

### Flujo actual

```excel
DB!E3 = SUBASTAS!K9
```

En `SUBASTAS!K9` existe la fórmula:

```excel
=DATE($D9,$E9,$F9)
```

Donde:

```text
SUBASTAS!D = AÑO
SUBASTAS!E = MES
SUBASTAS!F = DIA
```

### Origen directo

```text
subastas_AAMM.xlsx!D = AÑO
subastas_AAMM.xlsx!E = MES
subastas_AAMM.xlsx!F = DIA
```

### Regla

```python
Fecha = datetime(AÑO, MES, DIA)
```

En pandas, por ejemplo:

```python
df["Fecha"] = pd.to_datetime(
    dict(year=df["AÑO"], month=df["MES"], day=df["DIA"])
)
```

---

## 5.5. `DB!F` — Año

### Flujo actual

```excel
DB!F3 = SUBASTAS!D9
```

### Origen directo

```text
subastas_AAMM.xlsx!D:D = AÑO
```

### Regla

```python
Año = AÑO
```

Sin transformación.

---

## 5.6. `DB!G` — Mes

### Flujo actual

```excel
DB!G3 = SUBASTAS!E9
```

### Origen directo

```text
subastas_AAMM.xlsx!E:E = MES
```

### Regla

```python
Mes = MES
```

Sin transformación.

---

## 5.7. `DB!H` — Dia

### Flujo actual

```excel
DB!H3 = SUBASTAS!F9
```

### Origen directo

```text
subastas_AAMM.xlsx!F:F = DIA
```

### Regla

```python
Dia = DIA
```

Sin transformación.

---

## 5.8. `DB!I` — Hora_dia

### Flujo actual

```excel
DB!I3 = SUBASTAS!G9
```

### Origen directo

```text
subastas_AAMM.xlsx!G:G = HORA
```

### Regla

```python
Hora_dia = HORA
```

Sin transformación.

> Importante: la lógica observada utiliza directamente el valor de `HORA`; no se debe sumar o restar 1 salvo que en otro proceso se defina expresamente una corrección horaria.

---

## 5.9. `DB!J` — Hora_mes

### Flujo actual

```excel
DB!J3 = SUBASTAS!N9
```

En `SUBASTAS!N9` se observa la fórmula:

```excel
=($F9-1)*24+$G9+IF(F9>$F$2,1,0)
```

Por lo tanto, la base del cálculo es:

```text
Hora_mes = (DIA - 1) * 24 + HORA
```

más un posible ajuste asociado a `SUBASTAS!F2`.

### Origen directo

```text
subastas_AAMM.xlsx!F = DIA
subastas_AAMM.xlsx!G = HORA
```

### Regla exacta observada

```python
Hora_mes = (DIA - 1) * 24 + HORA + ajuste_cambio_hora
```

con:

```python
ajuste_cambio_hora = 1 if DIA > DIA_CAMBIO_HORA else 0
```

La celda usada por la fórmula actual como límite es:

```text
SUBASTAS!F2
```

### Observación importante

Este es el único campo de `DB!B:K` cuya reproducción directa necesita conservar la lógica mensual de **cambio de hora** de la Planilla 3.

Por lo tanto, al reemplazar la Planilla 3 se debe definir explícitamente de dónde vendrá `DIA_CAMBIO_HORA`.

Si para un mes no corresponde ajuste horario, el resultado práctico queda simplemente como:

```python
Hora_mes = (DIA - 1) * 24 + HORA
```

No conviene eliminar el ajuste del diseño definitivo hasta verificar cómo se configura `SUBASTAS!F2` en todos los meses.

---

## 5.10. `DB!K` — Configuración

### Flujo actual

```excel
DB!K3 = SUBASTAS!B9
```

### Origen directo

```text
subastas_AAMM.xlsx!B:B = CONFIGURACIÓN
```

### Regla

```python
Configuración = CONFIGURACIÓN
```

Sin transformación.

---

# 6. Equivalencia completa por registro

Tomando el primer registro del archivo `subastas_AAMM.xlsx`:

```text
subastas_AAMM.xlsx fila 2
```

se debe generar:

```text
DB fila 3
```

con el siguiente esquema:

```python
DB["Concepto"]      = src["SERVICIO"]
DB["Control"]       = src["SERVICIO"][:3]
DB["Sub_Baj"]       = "SUBIDA" si el servicio es (+), en otro caso "BAJADA"
DB["Fecha"]         = fecha(src["AÑO"], src["MES"], src["DIA"])
DB["Año"]           = src["AÑO"]
DB["Mes"]           = src["MES"]
DB["Dia"]           = src["DIA"]
DB["Hora_dia"]      = src["HORA"]
DB["Hora_mes"]      = (src["DIA"] - 1) * 24 + src["HORA"] + ajuste_cambio_hora
DB["Configuración"] = src["CONFIGURACIÓN"]
```

---

# 7. Transformación propuesta en pandas

Una implementación equivalente podría tener esta forma:

```python
import pandas as pd


df_subastas = pd.read_excel("subastas_AAMM.xlsx")


df_db = pd.DataFrame()

# DB B
df_db["Concepto"] = df_subastas["SERVICIO"]

# DB C
df_db["Control"] = df_subastas["SERVICIO"].astype(str).str[:3]

# DB D
df_db["Sub_Baj"] = df_subastas["SERVICIO"].astype(str).apply(
    lambda x: "SUBIDA" if "+" in x else "BAJADA"
)

# DB E
df_db["Fecha"] = pd.to_datetime(
    dict(
        year=df_subastas["AÑO"],
        month=df_subastas["MES"],
        day=df_subastas["DIA"],
    )
)

# DB F:I
df_db["Año"] = df_subastas["AÑO"]
df_db["Mes"] = df_subastas["MES"]
df_db["Dia"] = df_subastas["DIA"]
df_db["Hora_dia"] = df_subastas["HORA"]

# DB J
# DIA_CAMBIO_HORA debe definirse según la lógica mensual de Planilla 3.
df_db["Hora_mes"] = (
    (df_subastas["DIA"] - 1) * 24
    + df_subastas["HORA"]
    + (df_subastas["DIA"] > DIA_CAMBIO_HORA).astype(int)
)

# DB K
df_db["Configuración"] = df_subastas["CONFIGURACIÓN"]
```

Si se confirma que para el mes procesado no corresponde ajuste por cambio de hora:

```python
df_db["Hora_mes"] = (
    (df_subastas["DIA"] - 1) * 24
    + df_subastas["HORA"]
)
```

---

# 8. Columnas de `subastas_AAMM.xlsx` realmente necesarias para `DB!B:K`

Para generar `DB!B:K` no es necesario leer todas las columnas del archivo.

Son suficientes:

```text
B = CONFIGURACIÓN
C = SERVICIO
D = AÑO
E = MES
F = DIA
G = HORA
```

Es decir, conceptualmente:

```python
usecols = [
    "CONFIGURACIÓN",
    "SERVICIO",
    "AÑO",
    "MES",
    "DIA",
    "HORA",
]
```

Las columnas siguientes del archivo no son necesarias para construir `DB!B:K`:

```text
BANDA
CANTIDAD MW
PRECIO USD/MW
FECHA
CONTROL
SUB_BAJ
HORA_MES
CANTIDAD PONDERADA MW
Hora_PID
```

Aunque algunas sí son utilizadas en otras columnas posteriores de `DB`.

---

# 9. Relación con el script `entradas_sscc.py`

El script que genera `subastas_AAMM.xlsx` obtiene desde los archivos Access los campos:

```text
CONFIGURACIÓN
SERVICIO
AÑO
MES
DIA
HORA
BANDA
CANTIDAD MW
PRECIO USD/MW
CANTIDAD PONDERADA MW
```

Posteriormente agrega:

```text
FECHA
CONTROL
SUB_BAJ
HORA_MES
Hora_PID
```

Antes de exportar, deja las columnas en este orden:

```text
CONFIGURACIÓN
SERVICIO
AÑO
MES
DIA
HORA
BANDA
CANTIDAD MW
PRECIO USD/MW
FECHA
CONTROL
SUB_BAJ
HORA_MES
CANTIDAD PONDERADA MW
Hora_PID
```

Sin embargo, `FECHA`, `CONTROL`, `SUB_BAJ` y `HORA_MES` se crean vacíos en Python y luego la Planilla 3 los recalcula en `SUBASTAS`.

Por esto, la nueva implementación debe hacer esas cuatro transformaciones directamente en Python y no depender de las columnas vacías exportadas.

---

# 10. Resultado esperado del reemplazo

Una vez implementada esta transformación, el proceso que actualmente toma las columnas de Planilla 3 podrá usar directamente:

```text
subastas_AAMM.xlsx
```

para reproducir:

```text
DB!B:K
```

sin necesidad de que exista o esté terminada la Planilla 3.

Quedarán pendientes únicamente los campos que se acordó tratar aparte:

```text
DB!L = Propietario
DB!V = FMA
DB!Y = FD
```

---

# 11. Punto pendiente antes de automatizar definitivamente

El único punto que debe revisarse antes de considerar `DB!B:K` completamente independiente es:

```text
SUBASTAS!F2 → parámetro utilizado en el ajuste de Hora_mes
```

Debe determinarse cómo se fija ese valor cada mes y trasladar esa misma regla al nuevo proceso.

Todo el resto de `DB!B:K` puede generarse directamente desde `subastas_AAMM.xlsx` con las equivalencias detalladas en este documento.
