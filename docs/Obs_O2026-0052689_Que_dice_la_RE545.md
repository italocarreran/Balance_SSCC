# Qué dice la RE545 sobre la Componente 2

Resolución Exenta CNE N°545/2025, §2.2.2.2, "Remuneración por Costo de Oportunidad para un Sistema de Almacenamiento de Energía". Relacionado con la observación **O2026-0052689 (AES Andes)** al Balance Preliminar SSCC de octubre 2026 (`11_PAGOS_BESS_2610`, hoja `Calculo RE545`).

## Contexto

La RE545 paga a un BESS lo que dejó de ganar por guardar energía para dar reserva:

> **Pago = máx(Componente 1 − Componente 2, 0)**

- **Componente 1:** lo que habría ganado vendiendo la energía guardada (la "Energía Disponible") en los bloques de 15 minutos más caros de la ventana.
- **Componente 2:** lo que igual vendió. Se descuenta.

**Ejemplo para seguir el texto:**
- Una batería de **1 MW**. En un bloque de 15 minutos puede vender como máximo **250 kWh** (1.000 kW × ¼ h).
- Por la reserva tuvo que guardar **500 kWh** (Energía Disponible).
- Los dos bloques más caros de la ventana son **A1** (CMg 100) y **A2** (CMg 95).
- **Componente 1:** los 500 kWh se ponen en esos dos bloques: 250 × 100 + 250 × 95 = **48.750**.

## El texto de la Componente 2

> "**Componente 2_T**: corresponde a la valorización de la contabilización de las **inyecciones físicas** y de la **Energía Disponible** de un SAE o de la CACRCA, según corresponda, para la respectiva Ventana de Valorización T. Para lo anterior, se deberá realizar el siguiente procedimiento:
>
> 1. Se deberán contabilizar las **inyecciones físicas y valorizadas de la instalación en la operación real**, así como también, la **Energía Disponible asignada** en el numeral 4 anterior y su valorización, considerando los periodos de la curva descendente señalada en el numeral 2 anterior.
> 2. La contabilización de la Energía Disponible t y de las inyecciones físicas, deberá iniciar **desde la posición t igual a n hasta la posición t igual a 1**, **hasta alcanzar la Energía Disponible T** determinada previamente."

Y la definición de la misma sección:

> "El Costo de Oportunidad […] corresponderá a los costos incurridos o **pérdida de oportunidad de la venta de energía** en el mercado de corto plazo, ocasionados por la prestación de un Servicio Complementario."

## En simple, la RE dice tres cosas

1. **Qué se cuenta:** lo que la batería vendió de verdad (las "inyecciones físicas") y también la Energía Disponible que se puso en la Componente 1.
2. **En qué orden:** del bloque más barato (posición n) al más caro (posición 1).
3. **Hasta cuándo:** hasta juntar la Energía Disponible (los 500 kWh del ejemplo).

## Lo que la RE NO dice

No dice que, si la batería vendió algo en un bloque, se cuente la caja completa de ese bloque. Esa regla la puso la macro de la planilla.

**Ejemplo:** la batería vende **1 kWh en A1**.
- La planilla cuenta en A1 el 1 kWh vendido más los 249 kWh de Energía Disponible que la Componente 1 puso ahí: 250 kWh × 100 = **25.000**.
- El pago cae de 48.750 a **23.740**, cuando lo que realmente vendió vale **100**.

Eso es lo que reclama AES: una inyección pequeña hace descontar mucha más energía que la inyectada.

## El problema del texto

La letra al pie tampoco sirve. Si se cuenta la Energía Disponible en todos los bloques, también donde la batería no vendió nada:

- Con la batería sin vender nada, la Componente 2 contaría los 500 kWh que la Componente 1 puso en A1 y A2.
- Daría 48.750, igual que la Componente 1.
- **El pago sería 0.**

Eso contradice la definición de la propia RE: la batería no vendió porque guardó la energía para la reserva, y no se le pagaría nada.

## Entonces

- La letra de la RE es confusa: nombra las dos cosas, pero contar las dos siempre deja el pago en 0.
- La macro lo resolvió contando las dos solo donde hubo venta, y eso crea el problema que reclama AES.
- Contar **solo lo que la batería realmente vendió** es lo que mejor calza con la definición de la RE. En el ejemplo, por vender 1 kWh en A1 se descuentan 100, y el pago queda en 48.640.

## En fórmulas

La RE no trae fórmula para la Componente 2; esto es su texto escrito matemáticamente. Se deja fuera el ajuste de la última hora de la ventana (numeral 3), que aquí no influye. Se usa el mismo ejemplo de arriba.

### Notación

Los bloques de 15 minutos se ordenan de mayor a menor precio (la "curva descendente"): posición **t = 1** el más caro, **t = n** el más barato.

| Símbolo | Qué es | Ejemplo |
|---|---|---|
| $c_t$ | CMg del bloque en la posición $t$ | $c_1 = 100$ (A1), $c_2 = 95$ (A2), $c_3 = 90$ (A3) |
| $I_t$ | inyección real en ese bloque | $I_1 = 1$ kWh, el resto 0 |
| $P$ | caja: lo máximo por bloque, $P = P_{max} \cdot 1000 / 4$ | 250 kWh |
| $E$ | Energía Disponible de la ventana | 500 kWh |

### Componente 1 (numeral 4: "asignar secuencialmente" desde t = 1)

La Energía Disponible se reparte desde el bloque más caro, en el espacio libre de cada caja:

$$ED_t = \max\Big(0,\ \min\big(P - I_t,\ E - \textstyle\sum_{k<t} ED_k\big)\Big)$$

$$C_1 = \sum_{t=1}^{n} ED_t \cdot c_t$$

Ejemplo, con 1 kWh vendido en A1: $ED_1 = 249$, $ED_2 = 250$, $ED_3 = 1$, entonces $C_1 = 24.900 + 23.750 + 90 = 48.740$.

### Componente 2 según la letra de la RE

"Contabilizar las inyecciones físicas y la Energía Disponible t, desde t = n hasta t = 1, hasta alcanzar la Energía Disponible". En cada bloque se cuenta $I_t + ED_t$, partiendo por el más barato, con tope $E$:

$$X_t = \max\Big(0,\ \min\big(I_t + ED_t,\ E - \textstyle\sum_{k>t} X_k\big)\Big)$$

$$C_2 = \sum_{t=1}^{n} X_t \cdot c_t \qquad\qquad \text{Pago} = \max(C_1 - C_2,\ 0)$$

**El problema:** si la batería no vende nada ($I_t = 0$ en todos los bloques), queda $X_t = ED_t$ en cada bloque, porque la suma de los $ED_t$ es justo $E$. Entonces $C_2 = C_1$ y el pago es 0. Con 1 kWh vendido en A1 pasa lo mismo: $X_3 = 1$, $X_2 = 250$ y $X_1 = 249$, así que $C_2 = 48.740 = C_1$. En general, con esta lectura $C_2$ es el valor de los $E$ kWh más baratos entre inyecciones y Energía Disponible, y el pago solo sale positivo si la batería inyectó en bloques baratos.

### Componente 2 como la calcula hoy la planilla (la macro)

Lo mismo, pero solo en los bloques con inyección:

$$X_t = \begin{cases} 0 & \text{si } I_t = 0 \\ \max\Big(0,\ \min\big(I_t + ED_t,\ E - \sum_{k>t} X_k\big)\Big) & \text{si } I_t > 0 \end{cases}$$

Ejemplo: solo A1 tiene inyección, así que $X_1 = \min(1 + 249,\ 500) = 250$. Entonces $C_2 = 250 \cdot 100 = 25.000$ y el pago es $48.740 - 25.000 = 23.740$.

El salto está en la condición $I_t > 0$: basta que $I_t$ sea apenas mayor que 0 para que el bloque cuente $I_t + ED_t$ (casi $P$) en vez de 0. Esa condición no está en la RE.

### Componente 2 contando solo lo vendido (la propuesta)

$$X_t = \max\Big(0,\ \min\big(I_t,\ E - \textstyle\sum_{k>t} X_k\big)\Big)$$

Ejemplo: $X_1 = \min(1,\ 500) = 1$. Entonces $C_2 = 100$ y el pago es $48.740 - 100 = 48.640$.

Esta fórmula es continua: si $I_t$ cambia un poco, $C_2$ cambia un poco, sin saltos. La Energía Disponible queda solo como el tope $E$ del conteo.

### Resumen

| Lectura | Lo que cuenta cada bloque | Pago sin ventas | Pago si vende 1 kWh en A1 |
|---|---|---|---|
| Letra de la RE | $I_t + ED_t$, siempre | 0 | 0 |
| Planilla hoy (macro) | $I_t + ED_t$, solo si $I_t > 0$ | 48.750 | 23.740 |
| Propuesta | $I_t$ | 48.750 | 48.640 |
