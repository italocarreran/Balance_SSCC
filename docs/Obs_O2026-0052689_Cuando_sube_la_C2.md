# Cuándo sube la Componente 2 al contar solo lo vendido

Relacionado con la observación **O2026-0052689 (AES Andes)** al Balance Preliminar SSCC de octubre 2026 (`11_PAGOS_BESS_2610`, hoja `Calculo RE545`).

## Contexto

La RE545 paga a un BESS lo que dejó de ganar por guardar energía para dar reserva:

> **Pago = máx(Componente 1 − Componente 2, 0)**

- **Componente 1:** lo que habría ganado vendiendo la energía guardada (la "Energía Disponible") en los bloques de 15 minutos más caros, en el espacio que le queda libre a cada bloque.
- **Componente 2:** lo que igual vendió. Se cuenta desde el bloque más barato hacia el más caro, hasta juntar la Energía Disponible, y se descuenta.

**Las dos reglas que se comparan para la Componente 2:**

- **Regla de hoy (la macro):** en un bloque donde la batería vendió algo, aunque sea poco, se cuenta lo vendido **más** la energía que la Componente 1 puso en ese bloque.
- **Regla propuesta:** en cada bloque se cuenta **solo lo que la batería realmente vendió**.

Al comparar octubre con las dos reglas, la Componente 1 queda igual en todas las centrales y solo cambia la Componente 2. En la mayoría la Componente 2 baja y el pago sube. Pero en algunas centrales, como ESTEPA y DEL-DESIERTO, **la Componente 2 sube y el pago baja**. Este documento explica cuándo pasa eso.

## El ejemplo

- Una batería de **1 MW**. En un bloque de 15 minutos puede vender como máximo **250 kWh** (1.000 kW × ¼ h): es "la caja".
- Por la reserva tuvo que guardar **500 kWh** (Energía Disponible).
- Para simplificar, la ventana tiene 2 horas (A, la cara, y B, la barata) de 4 bloques cada una:

| Bloque | A1 | A2 | A3 | A4 | B1 | B2 | B3 | B4 |
|---|---|---|---|---|---|---|---|---|
| CMg | 100 | 95 | 90 | 85 | 50 | 45 | 40 | 35 |

- La batería vende en tres bloques (501 kWh en total):

| Bloque | CMg | Vendió |
|---|---|---|
| A1 | 100 | **250** (caja llena) |
| A3 | 90 | **1** |
| B1 | 50 | **250** (caja llena) |

## Componente 1 (igual con las dos reglas)

A1 está lleno con lo que vendió, así que los 500 kWh se ponen en el espacio libre de los bloques caros que siguen:

| Bloque | CMg | Caja libre | Se pone | Valor |
|---|---|---|---|---|
| A2 | 95 | 250 | 250 | 23.750 |
| A3 | 90 | 249 | 249 | 22.410 |
| A4 | 85 | 250 | 1 | 85 |
| **Total** | | | **500** | **46.245** |

## Componente 2 con la regla de hoy

Se cuenta desde el bloque más barato hasta juntar 500 kWh. En un bloque con venta se cuenta lo vendido + lo que puso la Componente 1:

| Bloque | CMg | Se cuenta | Acumulado | Valor |
|---|---|---|---|---|
| B1 | 50 | 250 + 0 = 250 | 250 | 12.500 |
| A3 | 90 | 1 + 249 = **250** | **500, tope lleno** | 22.500 |
| A1 | 100 | **0** (ya no cabe) | 500 | 0 |
| **Total** | | | | **35.000** |

> **Pago = 46.245 − 35.000 = 11.245**

## Componente 2 contando solo lo vendido

| Bloque | CMg | Se cuenta | Acumulado | Valor |
|---|---|---|---|---|
| B1 | 50 | 250 | 250 | 12.500 |
| A3 | 90 | **1** | 251 | 90 |
| A1 | 100 | **249** (lo que falta para 500) | 500 | 24.900 |
| **Total** | | | | **37.490** |

> **Pago = 46.245 − 37.490 = 8.755**

## Por qué sube la Componente 2

- **Con la regla de hoy**, en A3 la batería vendió 1 kWh, pero se contaron 250: los 249 que la Componente 1 imaginó ahí. Esos 249 llenaron el tope de 500 a precio 90, y la venta **real** de A1 (250 kWh a 100) quedó **fuera** del conteo.
- **Contando solo lo vendido**, A3 cuenta 1 kWh. Queda espacio en el tope y entra la venta real de A1, que es más cara.

La regla de hoy cambiaba ventas reales caras (A1 a 100) por energía imaginaria más barata (A3 a 90). Al corregirla, la Componente 2 puede subir.

## Cuándo pasa

La Componente 2 sube con la regla propuesta cuando se dan las tres cosas:

1. La batería vende poco en un bloque donde la Componente 1 puso harta energía (como A3). Con la regla de hoy, ese bloque contaba casi la caja entera.
2. Por eso el tope de Energía Disponible se llenaba antes de llegar a las ventas reales más caras.
3. Y la batería sí tenía ventas reales en bloques más caros (como A1).

Es lo que pasa en ESTEPA y DEL-DESIERTO, que venden harto en horas caras. En las centrales que venden poco ocurre lo contrario: la Componente 2 baja y el pago sube (por ejemplo, con solo 1 kWh vendido en A1, el pago pasa de 23.740 a 48.640).

Las dos cosas son coherentes con la regla propuesta: a unas centrales la regla de hoy les descontaba de más y a otras les contaba de menos sus ventas reales.

Los números salen de correr el código del repositorio (`Script/nucleo/re545_componentes.py`) con este ejemplo.
