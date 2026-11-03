# Ordenar la hoja "Calculo E Costos" — notas (sin programar todavía)

Pedido del usuario (2026-11-03): ir anotando los cambios y programarlos todos
juntos al final.

Referencia: hoja actual (margen en A; grupos en la fila 1, nombres en la 2).

## Anotado

1. Agregar una fila vacía al principio: los grupos pasan a la fila 2 y los
   nombres de columna a la 3.
2. Quitar el grupo **B1:D1** "Dia" (sobre Mes / Dia / Hora): no es relevante.
3. Quitar los grupos de **H1:L1**: "Nombre" (H1:I1, sobre Configuracion /
   Barra) y "BESS" (J1:L1, sobre Descarga / Carga / SoC).

4. Quitar las 4 columnas CTF (siempre 0): AF CTF(-) y AI CTF(+) de las
   prorratas, AL CTF(-) y AO CTF(+) del FD. Todo lo que está a la derecha de
   AF se corre.
5. Comentario en el encabezado de cada columna traída de otra hoja con su
   origen, p. ej. en CMg: "Hoja origen: CMg, Cuadro 1 (F_CMg); se busca por
   Barra + Bloque horario".

6. Comentario también en las columnas que usan un dato de otra hoja que NO
   aparece como columna, diciendo cómo se obtiene y con qué condición:

   | Columna (letra actual) | Dato oculto | Comentario |
   |---|---|---|
   | P SoC sobre el minimo | % Energía sobre mínima (Centrales, Cuadro 1) | 1 si SoC > umbral; el umbral es el de la PRIMERA fila de Resumen BESS, uno para todas las centrales |
   | AB / AC Energía descargada / cargada | Pmax (MW) (Centrales, Cuadro 1) | energía máxima del ciclo repartida en bloques de Pmax/4·1000 kWh según el Bloque ordenado; Pmax por Configuracion; sin Pmax queda vacía |
   | AT / AU Asignación Descarga / Carga | Pmax (MW) (Centrales, Cuadro 1) | ídem, con la energía ajustada FD-prorrata |
   | AD:AI Prorratas | Subastas, Cuadro 1 | no se copian: por central y Hora mes, filas de Subastas de cada servicio / total (la fila suma 1) |

## En evaluación

- Separar la hoja en ~3 cuadros (el primero con la información traída de
  otras hojas) **o** un código de colores por tipo de columna. El usuario
  va a pintar a mano para decidir.

## Columnas que vienen de otras hojas (letras actuales)

Tal cual (copia o búsqueda sin transformar):

| Col | Columna | Origen |
|---|---|---|
| B–H | Mes, Dia, Hora, Hora mes, Minuto, Bloque horario, Configuracion | Medidores, Cuadro 1 (Mes, Dia, Hora, Hora Mes, Minutos, Cuarto de Hora, clave) |
| L | SoC | Medidores (SoC) |
| M | Ciclo de Carga del mes | Medidores (Copia_Ventana) |
| I | Barra | Centrales, Cuadro 1 Resumen BESS (Barra inyección), por Configuracion |
| N | CMg | CMg, Cuadro 1 (F_CMg), por Barra + Bloque horario |

Traídas con condición:

| Col | Columna | Origen y condición |
|---|---|---|
| J/K | Descarga / Carga | Medidores (Gen_Unidad) separada por signo, solo filas con Ventana_No_Completa = 1 |
| O | Adj SSCC | Subastas, Cuadro 1: 1 si la central tiene subasta en ese mes/día/hora |
| AJ:AO | FD (CPF/CSF/CTF ±) | FD: CPF del Cuadro 2, CSF del Cuadro 1; por Hora mes + unidad, solo si la prorrata del servicio ≠ 0 |

Calculadas: el resto (incluidas las prorratas AD:AI, que salen de contar
Subastas por central y hora).

Usadas pero sin columna propia: Pmax (MW) de Centrales (bloques de
Energía descargada/cargada) y % Energía sobre mínima de Centrales (umbral de
"SoC sobre el minimo").

## Dudas en curso

- El umbral de "SoC sobre el minimo" es el % Energía sobre mínima de la
  primera fila de Resumen BESS para todas las centrales (como la celda fija
  H8 del original). ¿Debería ser el de cada central?
