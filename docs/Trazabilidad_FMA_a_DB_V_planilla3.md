# Trazabilidad FMA → columna `DB!V` de Planilla 3

## 1. Objetivo

Este documento describe cómo reconstruir la columna **`FMA` (`DB!V`)** de la Planilla 3 directamente a partir de las salidas generadas por `entradas_sscc.py`:

- `fma_cpf_AAMM.xlsx`
- `fma_csf_AAMM.xlsx`
- `fma_cft_AAMM.xlsx` / `fma_cft_AAMM.csv`

El objetivo final es **no depender de que la Planilla 3 esté terminada** para obtener el FMA de cada fila que ya puede construirse desde `subastas_AAMM.xlsx`.

La trazabilidad se obtuvo comparando:

1. la lógica de generación de los archivos FMA en `entradas_sscc.py`;
2. las salidas FMA entregadas para marzo de 2026;
3. las hojas `FMA_CPF`, `FMA_CSF`, `FMA_CTF` de la Planilla 3;
4. la fórmula utilizada actualmente en `DB!V`.

---

# 2. Resumen ejecutivo

`DB!V` depende del tipo de servicio indicado en `DB!B` (`Concepto`).

| `DB!B` | Fuente principal | Clave de búsqueda | Resultado base |
|---|---|---|---|
| `CPF(+)` | `fma_cpf_AAMM.xlsx` | Central FMA CPF + Año + Mes + Día + Hora | `FMA CPF(+)` |
| `CPF(-)` | `fma_cpf_AAMM.xlsx` | Central FMA CPF + Año + Mes + Día + Hora | `FMA CPF(-)` |
| `CSF(+)` | `fma_csf_AAMM.xlsx` | Año + Mes + Día + Hora | `FMA CSF-m(+) / 100` |
| `CSF(-)` | `fma_csf_AAMM.xlsx` | Año + Mes + Día + Hora | `FMA CSF-m(-) / 100` |
| `CTF(+)` | `fma_cft_AAMM.xlsx` | Configuración + Año + Mes + Día + Hora | suma de duración de activaciones positivas |
| `CTF(-)` | `fma_cft_AAMM.xlsx` | Configuración/Unidad + Año + Mes + Día + Hora | suma de duración de activaciones negativas |
| otro | — | — | `0` / error controlado |

Hay dos dependencias auxiliares importantes:

1. **CPF:** requiere traducir la unidad de Infotécnica a la nomenclatura usada por FMA CPF.
2. **CSF:** el valor FMA encontrado se multiplica posteriormente por el **Vector de Participación CSF** (`DB!AC`), proveniente de `CSF_FD`.

Por lo tanto, para reemplazar completamente `DB!V`:

```text
subastas_AAMM
      │
      ├── Concepto, Año, Mes, Día, Hora, Configuración
      │
      ├── CPF ──> fma_cpf_AAMM + diccionario nomenclatura CPF
      │
      ├── CSF ──> fma_csf_AAMM ──> FMA base ──> × Vector Participación CSF
      │
      └── CTF ──> fma_cft_AAMM ──> agregación por configuración/hora
                                      │
                                      ▼
                                   DB!V
```

---

# 3. Campos de DB utilizados para calcular FMA

La fórmula de `DB!V` utiliza principalmente:

| Columna DB | Nombre | Uso |
|---|---|---|
| `B` | Concepto | Determina CPF(+), CPF(-), CSF(+), CSF(-), CTF(+) o CTF(-) |
| `C` | Control | Se usa para aplicar el vector CSF |
| `F` | Año | Parte de la clave temporal |
| `G` | Mes | Parte de la clave temporal |
| `H` | Día | Parte de la clave temporal |
| `I` | Hora_día | Parte de la clave temporal |
| `K` | Configuración | Identificador utilizado especialmente en CTF |
| `N` | Unidad INFOTECNICA | Necesaria para traducir nombres en CPF |
| `AC` | Vector Participación CSF | Multiplicador final solamente para CSF |
| `BL` | Central FMA CPF- | Nombre de central compatible con el archivo FMA CPF |

## 3.1. Clave temporal común

La Planilla 3 arma las claves usando:

```text
Año_Mes_Día_Hora
```

Ejemplo:

```text
2026_3_1_5
```

No se utiliza `Hora_mes` para buscar FMA.

---

# 4. FMA CPF

## 4.1. Archivo generado por Python

`entradas_sscc.py` genera:

```text
fma_cpf_AAMM.xlsx
```

La salida contiene, entre otros:

| Campo salida Python |
|---|
| Año |
| Mes |
| Día |
| Hora |
| Central |
| Hace CPF [hrs] |
| Paráms. Fuera de rango |
| Hace CPF con lím superior >Pmax |
| Hace CPF con lím inferior <MT |
| No hace CPF [hrs] |
| Indef. [hrs] |
| Otro [hrs] |
| Tiempo f<49.975 [%] |
| Tiempo f>50.025 [%] |

En la salida revisada el Excel además contiene una columna índice de pandas a la izquierda. Esa columna **no forma parte del dato funcional**.

---

## 4.2. Qué hace actualmente la hoja `FMA_CPF`

Los datos de la salida se llevan a `FMA_CPF`.

La hoja agrega el campo:

```text
Suma Hace CPF [hrs]
```

con la lógica:

```excel
=SI(SUMA(Hace CPF : Otro)>=0,98;1;SUMA(Hace CPF : Otro))
```

Equivalentemente:

```python
suma_hace_cpf = suma(
    Hace_CPF,
    Parametros_fuera_rango,
    CPF_lim_superior_Pmax,
    CPF_lim_inferior_MT,
    No_hace_CPF,
    Indef,
    Otro
)

if suma_hace_cpf >= 0.98:
    suma_hace_cpf = 1
```

Luego calcula:

```text
FMA CPF(+) = Suma Hace CPF × Tiempo f<49.975
FMA CPF(-) = Suma Hace CPF × Tiempo f>50.025
```

En la hoja actual:

```excel
T = O * M   → FMA CPF(+)
U = O * N   → FMA CPF(-)
```

### Importante

Los campos `Tiempo f<49.975 [%]` y `Tiempo f>50.025 [%]` de la salida revisada ya vienen expresados como fracción decimal.

Ejemplo:

```text
0,33
```

representa `33 %`.

Por lo tanto **no se vuelve a dividir por 100**.

---

# 5. Nomenclatura necesaria para CPF

Este es un punto crítico para eliminar la dependencia de Planilla 3.

`DB!V` no busca CPF usando directamente `DB!K`.

Primero obtiene:

```text
DB!N = Unidad INFOTECNICA
```

y luego transforma esa unidad a la nomenclatura usada por el archivo FMA CPF.

En la planilla actual esa equivalencia está en `FMA_CPF!AD:AE`.

Ejemplo conceptual:

| Unidad INFOTECNICA | Central FMA CPF |
|---|---|
| `HE EL TORO U1` | `El Toro - U1` |
| `HE EL TORO U2` | `El Toro - U2` |
| ... | ... |

La lógica actual de `DB!BL` es:

```excel
SI(Control="CPF";
   BUSCARV(Unidad_INFOTECNICA; tabla_nomenclatura_CPF; 2; 0);
   ""
)
```

## Recomendación para la automatización

Mover esta equivalencia al archivo auxiliar de centrales/diccionarios que utilizará el nuevo proceso.

Campo sugerido:

```text
Unidad_INFOTECNICA | Central_FMA_CPF
```

Sin esta equivalencia no se puede reproducir CPF de forma robusta solamente con `subastas_AAMM.xlsx`.

---

# 6. Clave CPF

Una vez obtenida `Central_FMA_CPF`, la Planilla 3 crea:

```text
Central_FMA_CPF_Año_Mes_Día_Hora
```

Ejemplo:

```text
El Toro - U1_2026_3_1_5
```

La hoja `FMA_CPF` genera la misma clave:

```excel
=Central & "_" & Año & "_" & Mes & "_" & Día & "_" & Hora
```

Por lo tanto el cálculo directo puede hacerse con un `merge` por:

```python
[
    "Central_FMA_CPF",
    "Año",
    "Mes",
    "Día",
    "Hora"
]
```

No es necesario construir físicamente el texto concatenado si se hace en Python.

---

# 7. Resultado CPF hacia DB!V

Para cada fila de DB:

### `CPF(+)`

```python
FMA = FMA_CPF_mas
```

donde:

```python
FMA_CPF_mas = suma_hace_cpf * tiempo_f_menor_49_975
```

### `CPF(-)`

```python
FMA = FMA_CPF_menos
```

donde:

```python
FMA_CPF_menos = suma_hace_cpf * tiempo_f_mayor_50_025
```

Si la búsqueda no encuentra coincidencia, la fórmula actual termina entregando:

```text
0
```

debido al `IFERROR(...,0)` exterior de `DB!V`.

---

# 8. FMA CSF

## 8.1. Archivo generado por Python

`entradas_sscc.py` genera:

```text
fma_csf_AAMM.xlsx
```

La salida revisada contiene:

| Campo |
|---|
| Año |
| Mes |
| Día |
| Hora Día |
| Hora Mes |
| UTC |
| CSF-m [s] |
| FMA CSF-m [%] |
| CSF+m [s] |
| FMA CSF+m [%] |
| CSF- [s] |
| FMA CSF- [%] |
| CSF+ [s] |
| FMA CSF+ [%] |
| MOD [s] |
| MOD [%] |

La salida también posee una columna índice de pandas que debe ignorarse.

---

# 9. Transformación CSF realizada por Planilla 3

La hoja `FMA_CSF` agrega:

```text
Q = concatenado
R = FMA CSF(-)
S = FMA CSF(+)
```

La clave se construye como:

```excel
=Año & "_" & Mes & "_" & Día & "_" & (Hora Día + 1)
```

Es decir:

```text
Hora_DB = Hora Día del archivo FMA + 1
```

Esto es importante porque la salida FMA CSF utiliza horas:

```text
0 ... 23
```

mientras la búsqueda de `DB` utiliza:

```text
1 ... 24
```

---

## 9.1. Valor utilizado para CSF(-)

La Planilla 3 utiliza:

```excel
R = FMA CSF-m [%] / 100
```

Por lo tanto:

```python
FMA_CSF_menos_base = FMA_CSF_m_porcentaje / 100
```

**No utiliza directamente `FMA CSF- [%]`.**

---

## 9.2. Valor utilizado para CSF(+)

La Planilla 3 utiliza:

```excel
S = FMA CSF+m [%] / 100
```

Por lo tanto:

```python
FMA_CSF_mas_base = FMA_CSF_mas_m_porcentaje / 100
```

**No utiliza directamente `FMA CSF+ [%]`.**

Esta distinción debe mantenerse para replicar exactamente la Planilla 3.

---

# 10. Clave CSF

La clave utilizada por `DB!V` es:

```text
Año_Mes_Día_Hora_DB
```

Ejemplo:

Salida FMA:

```text
Año      = 2026
Mes      = 3
Día      = 1
Hora Día = 4
```

Clave para DB:

```text
2026_3_1_5
```

Por lo tanto, en Python:

```python
fma_csf["Hora_DB"] = fma_csf["Hora Día"] + 1
```

y el `merge` puede hacerse directamente por:

```python
["Año", "Mes", "Día", "Hora_DB"]
```

---

# 11. Vector de participación CSF

Aquí existe una dependencia adicional.

La fórmula final de `DB!V` termina con:

```excel
* SI(C="CSF"; AC; 1)
```

donde:

```text
AC = Vector Participación CSF
```

Por ello:

### CSF(-)

```python
FMA_final = FMA_CSF_menos_base * Vector_Participacion_CSF
```

### CSF(+)

```python
FMA_final = FMA_CSF_mas_base * Vector_Participacion_CSF
```

El vector se obtiene actualmente desde la hoja `CSF_FD`.

## Consecuencia para la nueva automatización

Con solamente:

```text
subastas_AAMM.xlsx
+
fma_csf_AAMM.xlsx
```

podemos calcular:

```text
FMA_CSF_base
```

pero **no reproducir el `DB!V` final de CSF** hasta incorporar la trazabilidad de `CSF_FD`.

Como FD se analizará aparte, se recomienda que el nuevo proceso conserve temporalmente dos campos:

```text
FMA_base
Vector_Participacion_CSF
```

y después:

```python
FMA = FMA_base * Vector_Participacion_CSF
```

Para CPF y CTF:

```python
Vector_Participacion_CSF = 1
```

---

# 12. FMA CTF

## 12.1. Archivo generado por Python

`entradas_sscc.py` lee el CSV CTF, elimina la zona horaria de:

```text
t0
tfin
```

y exporta:

```text
fma_cft_AAMM.xlsx
fma_cft_AAMM.csv
```

Los campos relevantes para reproducir `DB!V` son principalmente:

| Campo salida | Uso |
|---|---|
| `t0` | fecha/hora inicial de la activación |
| `tfin` | fecha/hora final |
| `unidad` | primera nomenclatura utilizada para construir clave |
| `Configuracion` | segunda nomenclatura disponible |
| `variacion` | determina si la activación es positiva o negativa |

El resto de columnas del archivo CTF no son necesarias para calcular `DB!V` según la lógica actual.

---

# 13. Transformaciones CTF en Planilla 3

La hoja `FMA_CTF` obtiene desde `t0`:

```excel
Año  = AÑO(t0)
Mes  = MES(t0)
Día  = DIA(t0)
Hora = HORA(t0) + 1
```

Por lo tanto:

```python
anio = t0.year
mes = t0.month
dia = t0.day
hora_db = t0.hour + 1
```

---

# 14. Claves CTF

La Planilla 3 crea dos claves.

## 14.1. Clave por unidad

```text
unidad_Año_Mes_Día_Hora
```

Ejemplo:

```text
PEHUENCHE-2_2026_7_1_3
```

## 14.2. Clave por configuración

```text
Configuracion_Año_Mes_Día_Hora
```

Ejemplo:

```text
PEHUENCHE_2026_7_1_3
```

La existencia de ambas claves permite manejar diferencias de nomenclatura entre el archivo de subastas y el archivo CTF.

---

# 15. Duración FMA de cada activación CTF

La Planilla 3 calcula:

```excel
FMA CTF = (tfin - t0) * 24
```

Es decir, duración en **horas**.

En Python:

```python
duracion_horas = (tfin - t0).total_seconds() / 3600
```

Luego separa la duración según el signo de `variacion`.

## CTF(-)

```excel
FMA CTF- = duración_horas * SI(variacion < 0; 1; 0)
```

En Python:

```python
fma_ctf_menos = duracion_horas if variacion < 0 else 0
```

## CTF(+)

```excel
FMA CTF+ = duración_horas * SI(variacion > 0; 1; 0)
```

En Python:

```python
fma_ctf_mas = duracion_horas if variacion > 0 else 0
```

---

# 16. Agregación CTF

La hoja `FMA_CTF` no utiliza cada activación individual directamente en `DB`.

Antes crea un resumen equivalente a una tabla dinámica.

Para cada clave se calcula:

```text
Cuenta de activaciones
Suma de FMA CTF-
Suma de FMA CTF+
```

Por lo tanto, si existen dos o más activaciones para la misma unidad/configuración y hora, deben **sumarse**.

Lógica Python equivalente:

```python
ctf_resumen_unidad = (
    fma_ctf
    .groupby(["unidad", "Año", "Mes", "Día", "Hora_DB"], as_index=False)
    .agg(
        Cantidad_Activaciones=("unidad", "size"),
        FMA_CTF_menos=("FMA_CTF_menos", "sum"),
        FMA_CTF_mas=("FMA_CTF_mas", "sum"),
    )
)
```

Debe existir también un resumen alternativo por:

```python
["Configuracion", "Año", "Mes", "Día", "Hora_DB"]
```

para reproducir el mecanismo de respaldo de nomenclatura.

---

# 17. Cómo busca CTF actualmente DB!V

## 17.1. CTF(-)

La lógica observable es:

1. buscar primero usando una clave basada en `DB!K`;
2. si no encuentra coincidencia en una de las nomenclaturas, intentar la alternativa;
3. devolver la suma de `FMA CTF-`.

Conceptualmente:

```python
valor = buscar_ctf_menos_por_unidad(clave)

if no_encontrado:
    valor = buscar_ctf_menos_por_configuracion(clave)
```

---

## 17.2. CTF(+): peculiaridad de la fórmula actual

La fórmula actual de `DB!V` contiene dos bloques consecutivos con condición:

```excel
B="CTF(+)"
```

Por la estructura de `SI` anidados, el segundo bloque `CTF(+)` queda **inalcanzable cuando el primero ya fue evaluado como verdadero**.

En consecuencia, el comportamiento real observado para CTF(+) es:

```text
buscar solamente en la primera tabla/rango utilizado por la fórmula
```

y si falla, el `IFERROR` exterior puede convertir el resultado en `0`.

### Posible intención original

La presencia del segundo bloque sugiere que probablemente se pretendía una lógica semejante a CTF(-):

```text
buscar por una nomenclatura
→ si no existe
buscar por la segunda nomenclatura
```

Pero esto **no debe modificarse automáticamente** sin validarlo contra resultados históricos.

Para una primera réplica debe conservarse el comportamiento actual; posteriormente puede compararse contra una versión corregida.

---

# 18. Fórmula conceptual completa de `DB!V`

La lógica puede resumirse así:

```python
if Concepto == "CPF(+)":
    FMA = buscar_CPF_mas(
        Central_FMA_CPF,
        Año,
        Mes,
        Día,
        Hora
    )

elif Concepto == "CPF(-)":
    FMA = buscar_CPF_menos(
        Central_FMA_CPF,
        Año,
        Mes,
        Día,
        Hora
    )

elif Concepto == "CSF(+)":
    FMA_base = buscar_CSF_mas(
        Año,
        Mes,
        Día,
        Hora
    )
    FMA = FMA_base * Vector_Participacion_CSF

elif Concepto == "CSF(-)":
    FMA_base = buscar_CSF_menos(
        Año,
        Mes,
        Día,
        Hora
    )
    FMA = FMA_base * Vector_Participacion_CSF

elif Concepto == "CTF(+)":
    FMA = buscar_CTF_mas(
        Configuración,
        Año,
        Mes,
        Día,
        Hora
    )

elif Concepto == "CTF(-)":
    FMA = buscar_CTF_menos(
        Configuración,
        Año,
        Mes,
        Día,
        Hora
    )

else:
    FMA = 0
```

Ante una búsqueda fallida, el comportamiento general de la fórmula actual es:

```python
FMA = 0
```

---

# 19. Diseño recomendado para el nuevo código

Se recomienda **no imitar BUSCARV ni concatenaciones de Excel**. Es más robusto normalizar las salidas y hacer `merge` por columnas.

## 19.1. Tabla normalizada CPF

Generar:

```text
Central_FMA_CPF
Año
Mes
Día
Hora
FMA_CPF_mas
FMA_CPF_menos
```

Clave lógica:

```text
Central_FMA_CPF + Año + Mes + Día + Hora
```

---

## 19.2. Tabla normalizada CSF

Generar:

```text
Año
Mes
Día
Hora_DB
FMA_CSF_mas_base
FMA_CSF_menos_base
```

donde:

```python
Hora_DB = Hora_Día + 1
FMA_CSF_menos_base = FMA_CSF_m_porcentaje / 100
FMA_CSF_mas_base   = FMA_CSF_mas_m_porcentaje / 100
```

---

## 19.3. Tabla normalizada CTF

Generar dos tablas o una tabla con ambas nomenclaturas:

```text
unidad
Configuracion
Año
Mes
Día
Hora_DB
FMA_CTF_mas
FMA_CTF_menos
```

Las activaciones repetidas deben agregarse por hora.

---

# 20. Pseudocódigo de preparación

```python
# ==========================================================
# CPF
# ==========================================================

cpf["Suma_Hace_CPF"] = (
    cpf["Hace CPF [hrs]"]
    + cpf["Paráms. Fuera de rango"]
    + cpf["Hace CPF con lím superior >Pmax"]
    + cpf["Hace CPF con lím inferior <MT"]
    + cpf["No hace CPF [hrs]"]
    + cpf["Indef. [hrs]"]
    + cpf["Otro [hrs]"]
)

cpf["Suma_Hace_CPF"] = cpf["Suma_Hace_CPF"].where(
    cpf["Suma_Hace_CPF"] < 0.98,
    1
)

cpf["FMA_CPF_mas"] = (
    cpf["Suma_Hace_CPF"]
    * cpf["Tiempo f<49.975 [%]"]
)

cpf["FMA_CPF_menos"] = (
    cpf["Suma_Hace_CPF"]
    * cpf["Tiempo f>50.025 [%]"]
)


# ==========================================================
# CSF
# ==========================================================

csf["Hora_DB"] = csf["Hora Día"] + 1

csf["FMA_CSF_menos_base"] = (
    csf["FMA CSF-m [%]"] / 100
)

csf["FMA_CSF_mas_base"] = (
    csf["FMA CSF+m [%]"] / 100
)


# ==========================================================
# CTF
# ==========================================================

ctf["Año"] = ctf["t0"].dt.year
ctf["Mes"] = ctf["t0"].dt.month
ctf["Día"] = ctf["t0"].dt.day
ctf["Hora_DB"] = ctf["t0"].dt.hour + 1

ctf["Duracion_horas"] = (
    ctf["tfin"] - ctf["t0"]
).dt.total_seconds() / 3600

ctf["FMA_CTF_menos"] = np.where(
    ctf["variacion"] < 0,
    ctf["Duracion_horas"],
    0
)

ctf["FMA_CTF_mas"] = np.where(
    ctf["variacion"] > 0,
    ctf["Duracion_horas"],
    0
)
```

---

# 21. Datos auxiliares que todavía deben externalizarse

Para que el proceso quede completamente independiente de Planilla 3 hay que conservar fuera de ella:

## 21.1. Diccionario CPF

```text
Unidad_INFOTECNICA → Central_FMA_CPF
```

Actualmente vive dentro de la hoja `FMA_CPF`.

Se recomienda llevarlo al archivo auxiliar `Centrales.xlsx`.

## 21.2. Vector de Participación CSF

```text
Unidad/Hora → Vector_Participacion_CSF
```

Actualmente se obtiene a partir de `CSF_FD`.

Este punto se completará cuando se documente la trazabilidad FD.

---

# 22. Qué ya puede automatizarse

Con la información disponible se puede implementar ahora:

- lectura directa de `fma_cpf_AAMM.xlsx`;
- cálculo de `FMA CPF(+)`;
- cálculo de `FMA CPF(-)`;
- lectura directa de `fma_csf_AAMM.xlsx`;
- cálculo de `FMA CSF(+) base`;
- cálculo de `FMA CSF(-) base`;
- lectura directa de `fma_cft_AAMM.xlsx` o CSV;
- cálculo de duración de cada activación CTF;
- separación CTF(+) / CTF(-);
- agregación horaria de múltiples activaciones;
- cálculo directo de CTF(+) y CTF(-);
- unión de estos resultados a las filas construidas desde `subastas_AAMM.xlsx`.

Para obtener **exactamente** `DB!V` en CSF falta incorporar:

```text
Vector_Participacion_CSF
```

---

# 23. Flujo objetivo final

```text
entradas_sscc.py
│
├── subastas_AAMM.xlsx
│       │
│       └── genera filas base DB
│
├── fma_cpf_AAMM.xlsx
│       │
│       ├── calcular Suma Hace CPF
│       ├── calcular FMA CPF(+)
│       └── calcular FMA CPF(-)
│
├── fma_csf_AAMM.xlsx
│       │
│       ├── ajustar Hora Día + 1
│       ├── calcular FMA CSF(-) base
│       └── calcular FMA CSF(+) base
│
└── fma_cft_AAMM.xlsx
        │
        ├── extraer fecha/hora desde t0
        ├── calcular duración tfin - t0
        ├── separar signo de variación
        └── agrupar por configuración/unidad y hora

                       │
                       ▼

              dataframe equivalente a DB

                       │
                       ├── CPF → FMA directo
                       ├── CTF → FMA directo
                       └── CSF → FMA base × vector CSF

                       │
                       ▼

                  columna FMA
                  equivalente a DB!V
```

---

# 24. Validación recomendada al implementar

Antes de reemplazar la Planilla 3, ejecutar ambos métodos para un mes histórico:

```text
Método A = DB!V de Planilla 3
Método B = FMA calculado directamente por Python
```

Comparar fila a fila por una clave como:

```text
Concepto + Configuración + Año + Mes + Día + Hora
```

Generar:

```text
FMA_Planilla3
FMA_Python
Diferencia
```

y exigir:

```python
abs(FMA_Planilla3 - FMA_Python) < tolerancia
```

Debe revisarse por separado:

- CPF(+)
- CPF(-)
- CSF(+)
- CSF(-)
- CTF(+)
- CTF(-)

Especial atención a:

1. cambios de nomenclatura CPF;
2. cambio de hora `0–23` → `1–24`;
3. activaciones CTF repetidas dentro de una hora;
4. diferencias entre `unidad` y `Configuracion` en CTF;
5. vector de participación CSF;
6. comportamiento actual de la fórmula duplicada para `CTF(+)`.

---

# 25. Conclusión

La columna `DB!V` puede desacoplarse prácticamente por completo de la Planilla 3.

Los tres archivos FMA generados por `entradas_sscc.py` contienen la información primaria necesaria. La Planilla 3 realiza principalmente:

```text
normalización
+ construcción de claves
+ cálculos simples
+ agregación CTF
+ búsquedas
```

La única dependencia adicional para reproducir el resultado final es:

```text
CPF → diccionario Unidad INFOTECNICA ↔ Central FMA CPF
CSF → Vector Participación proveniente de CSF_FD
```

Por ello, la arquitectura recomendada es que el nuevo código lea directamente:

```text
subastas_AAMM.xlsx
fma_cpf_AAMM.xlsx
fma_csf_AAMM.xlsx
fma_cft_AAMM.xlsx
Centrales.xlsx / diccionarios
```

y genere la columna `FMA` sin esperar que la Planilla 3 haya sido completada.
