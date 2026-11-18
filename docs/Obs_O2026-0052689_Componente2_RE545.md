# Qué dice la RE545 sobre la Componente 2

Resolución Exenta CNE N°545/2025, §2.2.2.2 "Remuneración por Costo de Oportunidad para un Sistema de Almacenamiento de Energía". Relacionado con la observación O2026-0052689 (AES Andes).

## El texto

> "**Componente 2_T**: corresponde a la valorización de la contabilización de las **inyecciones físicas** y de la **Energía Disponible** […]
>
> 1. Se deberán contabilizar las **inyecciones físicas y valorizadas de la instalación en la operación real**, así como también, la **Energía Disponible asignada** en el numeral 4 anterior […]
> 2. La contabilización […] deberá iniciar **desde la posición t igual a n hasta la posición t igual a 1**, **hasta alcanzar la Energía Disponible**."

## En simple, la RE dice tres cosas

1. **Qué se cuenta:** lo que la batería vendió de verdad (las "inyecciones físicas") y también la Energía Disponible que se puso en la Componente 1.
2. **En qué orden:** del bloque más barato (posición n) al más caro (posición 1).
3. **Hasta cuándo:** hasta juntar la Energía Disponible (los 500 kWh del ejemplo).

## Lo que la RE NO dice

- No dice "si vendiste algo en un bloque, cuenta la caja completa".
- La regla de que con 1 kWh vendido se cuenten los 250 kWh de A1 la puso la macro. No está en la RE.

## El problema del texto

Si se aplica la letra al pie, contando también la Energía Disponible en todos los bloques, en el caso sin ventas pasa esto:

- La Componente 2 contaría los 500 kWh que la Componente 1 puso en A1 y A2.
- Daría 48.750, igual que la Componente 1.
- **El pago sería 0.**

Eso no tiene sentido: la batería guardó la energía por la reserva, no vendió nada, y no se le pagaría nada. Contradice la definición de la propia RE: el costo de oportunidad es "la **pérdida de oportunidad de la venta de energía**".

## Entonces

- La letra de la RE es confusa: nombra las dos cosas, pero contar las dos siempre deja el pago en 0.
- La macro lo resolvió a su manera, contando las dos solo donde hubo venta, y eso crea el problema de AES.
- Contar **solo lo que realmente vendió** es lo que mejor calza con la idea de la RE: le descuentas lo que igual ganó, ni más ni menos.
