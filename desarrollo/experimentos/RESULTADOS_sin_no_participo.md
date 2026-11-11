# Experimento: ¿cuánto cambia si no se pagan las horas "No participó"?

Rama `claude/experimento-no-participo` (paréntesis: **no va a main**).
Script: `desarrollo/experimentos/sin_no_participo.py`.

## Método

Dos copias de 11_PAGOS_BESS_2610 recalculadas con el código actual (Calculo E
Costos, Calculo RE545, COMPENSACION_CENTRAL): la base tal cual y el escenario
con FD = 0 en las horas "No participó" de la hoja FD. Se comparan con
Comparar_Meses (responsables en pesos). La FMA no se toca.

Los dos cálculos usan el FD distinto:
- **E Costos**: por dirección (columnas CPF(+), CPF(−), CSF(+), CSF(−)).
- **RE545**: el Factor de Desempeño del servicio (Fd_CPF / Fd_CSF), sin dirección.

Por eso hay dos escenarios:
- **Por dirección** (principal): la dirección que no participó va a FD = 0; el
  factor del servicio, solo si ninguna de sus direcciones participó.
- **Estricto** (cota de arriba): el factor del servicio va a 0 si cualquiera
  de sus direcciones dice "No participó", aunque la dirección adjudicada sí
  haya respondido.

## Resultados, octubre 2026

| | Total del mes | RE545 | E Costos |
|---|---|---|---|
| Base | $1.841 MM | $1.349 MM | $492 MM |
| Por dirección | $1.829 MM (−$11 MM, −0,6 %) | $1.343 MM | $486 MM |
| Estricto | $1.572 MM (−$268 MM, −15 %) | $1.086 MM | $486 MM |

- En la dirección adjudicada, "No participó" es poco frecuente: 1 % de las
  filas CPF y 7-10 % de las CSF, y la mayoría ya tienen FD = 0. Por eso el
  escenario por dirección casi no cambia nada.
- La diferencia con el estricto viene de RE545: horas en que una dirección
  CPF dice "No participó" pero la adjudicada sí respondió, y el factor del
  servicio igual es 1.
- Por dirección, el mayor cambio: NUEVO-QUILLAGUA −$5 MM, ARICA2 −$3 MM,
  DEL-DESIERTO −$2 MM, MANZANO −$2 MM.
- Caso raro: ALFALFAL sube (+$1 MM), porque en E Costos el FD también
  multiplica la carga, que resta en C1. Con FD = 0 en horas de carga, esa
  carga deja de restar.
