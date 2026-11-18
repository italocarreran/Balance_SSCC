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
