# Observación O2026-0052689 (AES Andes): Componente 2 de la RE545

**Balance:** Preliminar SSCC octubre 2026, archivo `11_PAGOS_BESS_2610`, hoja `Calculo RE545`.

**Lo que observa AES Andes:**

> "Con el cálculo actual de la Asignación de Energía Disponible de la Componente 2 se puede provocar que una inyección pequeña habilite el descuento de una cantidad de energía muy superior a la efectivamente inyectada, reduciendo la compensación asociada a una venta que no se materializó."

**Lo que pide:** revisar y corregir el cálculo, para que la compensación cumpla con la definición de costo de oportunidad de la resolución.

---

## La idea de la RE545

La batería tuvo que **guardar energía** para dar reserva (SSCC) y por eso no la pudo vender. Se le paga lo que perdió:

> **Pago = Componente 1 − Componente 2** (nunca negativo)

- **Componente 1:** lo que habría ganado vendiendo la energía guardada en los bloques más caros.
- **Componente 2:** lo que igual vendió. Se descuenta porque, si igual vendió, no perdió tanto.

---

## El ejemplo, paso a paso

### Paso 1: La batería y el día

- Potencia máxima **1 MW**. En un bloque de 15 minutos puede vender como máximo 1.000 kW × ¼ h = **250 kWh** ("la caja", Pmax × 1000 / 4).
- Para simplificar, el día tiene **2 horas** (en la realidad la ventana es de 24 h). Cada hora tiene 4 bloques de 15 minutos, así que son **8 bloques**.

| Hora | Bloque | CMg |
|---|---|---|
| Hora A | A1 | 100 |
| Hora A | A2 | 95 |
| Hora A | A3 | 90 |
| Hora A | A4 | 85 |
| Hora B | B1 | 50 |
| Hora B | B2 | 45 |
| Hora B | B3 | 40 |
| Hora B | B4 | 35 |

La hora A es la cara y la hora B la barata.

### Paso 2: La reserva y la Energía Disponible

La reserva se adjudica **por hora**: 0,25 MW de subida en las dos horas, con FD = 1.

| Hora | Reserva | Energía comprometida |
|---|---|---|
| Hora A | 0,25 MW × 1 h | 250 kWh |
| Hora B | 0,25 MW × 1 h | 250 kWh |
| **Total** | | **500 kWh** |

La **Energía Disponible** es de **500 kWh**: lo que la batería tuvo que tener guardado sin vender, por si el sistema lo necesitaba. La batería tenía esa energía, así que no se recorta.

### Paso 3: Componente 1, sin ventas

Si no hubiera tenido que guardar la energía, la habría vendido en lo más caro. Los 500 kWh se ponen en los bloques más caros primero, en cajas de 250 kWh:

| Bloque | CMg | Caja libre | Se pone | Valor |
|---|---|---|---|---|
| A1 | 100 | 250 | 250 | 25.000 |
| A2 | 95 | 250 | 250 | 23.750 |
| A3 a B4 | 90 a 35 | 250 | 0 | 0 |
| **Total** | | | **500** | **48.750** |

**Componente 1 = 48.750**

### Paso 4: Componente 2, sin ventas

Se cuenta lo que vendió, desde el bloque más barato hacia el más caro, hasta juntar 500 kWh. No vendió nada:

**Componente 2 = 0**

> **Pago = 48.750 − 0 = 48.750**

Tiene sentido: no vendió porque guardó la energía para la reserva, así que se le paga todo lo que habría ganado.

### Paso 5: Ahora vende 1 kWh en A1 (Componente 1)

Ese 1 kWh ocupa parte de la caja de A1, que queda con **249 kWh** libres.

| Bloque | CMg | Vendió | Caja libre | Se pone | Valor |
|---|---|---|---|---|---|
| A1 | 100 | 1 | 249 | 249 | 24.900 |
| A2 | 95 | 0 | 250 | 250 | 23.750 |
| A3 | 90 | 0 | 250 | 1 | 90 |
| A4 a B4 | 85 a 35 | 0 | 250 | 0 | 0 |
| **Total** | | | | **500** | **48.740** |

**Componente 1 = 48.740**, casi igual que antes.

### Paso 6: Componente 2 como la calcula hoy la planilla

La regla actual (heredada de la macro):

- En un bloque donde **no vendió nada**, cuenta 0.
- En un bloque donde **vendió algo, aunque sea poco**, cuenta **lo vendido más la energía que la Componente 1 puso en ese bloque**, sin pasar los 500 kWh.

| Bloque | CMg | Vendió | Puso la C1 | ¿Vendió algo? | Se cuenta | Valor |
|---|---|---|---|---|---|---|
| B4 a A2 | 35 a 95 | 0 | 0 a 250 | No | 0 | 0 |
| **A1** | **100** | **1** | **249** | **Sí** | **1 + 249 = 250** | **25.000** |
| **Total** | | | | | **250** | **25.000** |

**Componente 2 = 25.000**

> **Pago = 48.740 − 25.000 = 23.740**

**El problema:** vendió **1 kWh**, que vale **100**, y se le descuentan **25.000**, como si hubiera vendido la caja completa de 250 kWh. Por 1 kWh vendido pierde casi la mitad del pago. Eso es lo que reclama AES.

### Paso 7: Componente 2 contando solo lo que vendió

En cada bloque se cuenta **solo lo que la batería realmente vendió**, con el mismo orden y el mismo tope:

| Bloque | CMg | Vendió | Se cuenta | Valor |
|---|---|---|---|---|
| B4 a A2 | 35 a 95 | 0 | 0 | 0 |
| **A1** | **100** | **1** | **1** | **100** |
| **Total** | | | **1** | **100** |

**Componente 2 = 100**

> **Pago = 48.740 − 100 = 48.640**

### Comparación

| Caso | Pago hoy (planilla) | Pago contando solo lo vendido |
|---|---|---|
| No vendió nada | 48.750 | 48.750 |
| Vendió 1 kWh en A1 | **23.740** | **48.640** |
| Vendió 1 kWh en cada bloque de la hora A (4 kWh) | **30** | **48.365** |
| Vendió la caja llena (250 kWh) en A1 | 21.250 | 21.250 |

- Si vendió **parte** de un bloque, la corrección descuenta solo lo vendido.
- Si vendió la **caja llena**, da lo mismo que hoy.
- Si **no vendió nada**, tampoco cambia.

Los números salen de correr el código del repositorio (`Script/nucleo/re545_componentes.py`) con este ejemplo.

---

## Qué dice la RE545 (§2.2.2.2)

**Definición del costo de oportunidad:**

> "El Costo de Oportunidad generado por la prestación de un Servicio Complementario de un Sistema de Almacenamiento de Energía […] corresponderá a los costos incurridos o **pérdida de oportunidad de la venta de energía** en el mercado de corto plazo, ocasionados por la prestación de un Servicio Complementario."

> Remuneración Costo de Oportunidad_T = máx{(Componente 1_T − Componente 2_T), 0}

**Componente 2:**

> "**Componente 2_T**: corresponde a la valorización de la contabilización de las **inyecciones físicas y de la Energía Disponible** de un SAE o de la CACRCA, según corresponda, para la respectiva Ventana de Valorización T. Para lo anterior, se deberá realizar el siguiente procedimiento:
>
> 1. Se deberán contabilizar las **inyecciones físicas y valorizadas de la instalación en la operación real**, así como también, la **Energía Disponible asignada en el numeral 4 anterior** y su valorización, considerando los periodos de la curva descendente señalada en el numeral 2 anterior.
> 2. La contabilización de la Energía Disponible t y de las inyecciones físicas, deberá iniciar **desde la posición t igual a n hasta la posición t igual a 1**, **hasta alcanzar la Energía Disponible T** determinada previamente.
> 3. En el caso de que la instalación preste un servicio complementario en la última hora de la Ventana de Valorización T, […]"

### En simple, la RE dice tres cosas

1. **Qué se cuenta:** lo que la batería vendió de verdad (las inyecciones físicas) y también la Energía Disponible que se puso en la Componente 1.
2. **En qué orden:** del bloque más barato (posición n) al más caro (posición 1).
3. **Hasta cuándo:** hasta juntar la Energía Disponible (los 500 kWh del ejemplo).

### Lo que la RE NO dice

- No dice "si vendiste algo en un bloque, cuenta la caja completa".
- La regla de que con 1 kWh vendido se cuenten 250 kWh la puso la macro. No está en la RE.

### El problema de leer el texto al pie de la letra

Si se cuenta la Energía Disponible en **todos** los bloques, también donde no vendió, en el caso sin ventas:

- La Componente 2 contaría los 500 kWh que la Componente 1 puso en A1 y A2.
- Daría 48.750, igual que la Componente 1.
- **El pago sería 0.**

Eso contradice la propia definición de la RE: la batería guardó la energía por la reserva, no vendió nada, y no se le pagaría nada.

### Conclusión

- La letra de la RE es confusa: nombra las dos cosas, pero contarlas siempre deja el pago en 0.
- La macro lo resolvió contando las dos **solo donde hubo venta**, y eso crea el problema que reclama AES.
- Contar **solo lo que realmente vendió** es lo que mejor calza con la definición de la RE ("pérdida de oportunidad de la venta de energía"): se descuenta lo que igual ganó, ni más ni menos. Con este cambio, la Energía Disponible que la RE nombra en la Componente 2 queda solo como el **tope** del conteo.

---

## Propuesta

- **Respuesta a AES:** acoger la observación.
- **Cambio en el cálculo:** en la Componente 2 ("Asignacion Edisponible"), cada bloque cuenta solo la inyección real del bloque, con el mismo orden (de menor a mayor CMg) y el mismo tope de hoy. Hay que corregir también la macro.
- **Pendiente:** cuantificar el efecto en octubre 2026 por central con el `11_PAGOS_BESS_2610`.
