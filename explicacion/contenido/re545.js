/* ==================================================================
   explicacion/contenido/re545.js — CONTENIDO (lo escribe Claude).
   Solo textos y estructura de pasos. Los números van como {nombre}
   (salen de datos/ejemplo.js). Nada de letras de columnas de Excel.
   ================================================================== */
window.CONTENIDO = {
  hoja: "re545",
  eyebrow: "Balance BESS · Calculo RE545",
  titulo: "Paso a paso RE545",
  bajada: "Cómo se calcula el costo de oportunidad de un BESS cuando la ventana sí tuvo oferta completa (Resolución Exenta CNE N°545, §2.2.2.2), en pasos cortos y con un mismo ejemplo. Los números salen de correr el código del repositorio.",
  caso: [
    {valor: "SAE-DEMO · ventana 1", detalle: "una central, una ventana"},
    {valor: "{pmax_mw} MW · {capacidad_mwh} MWh", detalle: "Pmax y capacidad"},
    {valor: "{eficiencia}", detalle: "eficiencia"},
    {valor: "{mw_cpf} MW CPF(+) · {mw_csf} MW CSF(−)", detalle: "adjudicados las 3 horas · FD {fd} · FMA {fma}"},
  ],
  pasos: [
    {
      id: "grupo", titulo: "El grupo: central y ventana",
      texto: "Todo se calcula por <b>central + ventana de valorización</b>. La ventana dura 24 horas: 96 bloques de 15 minutos, desde la hora 10 de un día hasta la hora 9 del siguiente. Solo entran las ventanas con <b>oferta completa</b>. Cada grupo es independiente.",
      resultado: {etiqueta: "", valor: "SAE-DEMO · ventana 1", detalle: "en el ejemplo, reducida a {n_bloques:0} bloques"},
      bloques: [{tipo: "nota", titulo: "Dos tablas en la misma hoja.", texto: "Una tiene una fila por bloque de 15 minutos. La otra, el resumen, tiene una fila por central y ventana, y ahí quedan la energía inicial, la almacenada, la reserva total y la Energía Disponible. Las filas de las dos tablas no se corresponden."}],
    },
    {
      id: "inicial", titulo: "Energía inicial",
      texto: "Lo que tenía la batería al empezar la ventana. Se toma el SoC del <b>primer bloque</b>.",
      calculo: "SoC del primer bloque × Capacidad × 1000 = {re_soc0} × {capacidad_mwh} × 1000 = {re_eini} kWh",
      resultado: {etiqueta: "Energía inicial", valor: "{re_eini} kWh"},
      bloques: [{tipo: "nota", texto: "La capacidad sale de Resumen BESS (<code>Centrales.xlsx</code>)."}],
    },
    {
      id: "almacenada", titulo: "Energía almacenada",
      texto: "Todo lo que cargó en la ventana, descontada la eficiencia. Se suman <b>solo las cargas</b>: las descargas no se restan.",
      calculo: "Σ carga de la ventana × Eficiencia = {re_carga} × {eficiencia} = {re_ealm} kWh",
      resultado: {etiqueta: "Energía almacenada", valor: "{re_ealm} kWh"},
      bloques: [{tipo: "nota", texto: "La eficiencia sale de Resumen BESS. La RE545 la define como la energía almacenada durante la ventana; por eso no descuenta lo descargado."}],
    },
    {
      id: "reserva", titulo: "Reserva de cada bloque",
      texto: "En cada bloque se pasa la reserva adjudicada a energía. Cada servicio se multiplica por su FD. En la subida cuenta toda la reserva asignada; en la bajada, solo la parte activada, que se estima con el FMA.",
      calculo: "CPF(+): {mw_cpf} MW × FD {fd} × 1         = {mw_cpf_fd} MW\nCSF(−): {mw_csf} MW × FD {fd} × FMA {fma} = {mw_csf_fd_fma} MW\n{mw_reserva} MW ÷ 4 × 1000 = {re_res_bloque} kWh por bloque",
      resultado: {etiqueta: "Reserva por bloque", valor: "{re_res_bloque} kWh"},
      bloques: [
        {tipo: "nota", titulo: "De dónde salen los MW", texto: "Del Access de subastas (<code>OfertasSSCCAdj*.accdb</code>), cantidad ponderada en MW, sumada por central, hora y servicio. Da lo mismo pasar los MW a kWh por bloque antes o después de multiplicar por FD y FMA."},
        {tipo: "nota", texto: "RE545: reserva de subida <i>asignada</i> × FD + reserva de bajada <i>activada</i> × FD."},
      ],
    },
    {
      id: "total", titulo: "Reserva total de la ventana",
      texto: "Se suman las reservas de todos los bloques de la ventana. Es el tope de energía que la reserva pudo retener.",
      calculo: "{re_res_bloque} × {n_bloques:0} bloques = {re_res_total} kWh",
      resultado: {etiqueta: "Reserva total", valor: "{re_res_total} kWh"},
    },
    {
      id: "disponible", titulo: "Energía Disponible",
      texto: "La energía que la reserva le impidió usar para arbitrar. Es la menor entre lo que la batería tenía y almacenó, y la reserva total.",
      partes: [
        "<b>Inicial + almacenada:</b> lo que tuvo para entregar en la ventana.",
        "<b>mín(…, Capacidad):</b> la energía inicial se suma solo hasta llenar la batería.",
        "<b>máx(…, Almacenada):</b> lo cargado en la ventana nunca se recorta, aunque pase la capacidad, porque puede venir de varios ciclos.",
        "<b>mín(…, Reserva total):</b> no puede pasar lo que la reserva comprometió.",
        "Vale 0 si la ventana no tuvo oferta completa, y también en la ventana 31.",
      ],
      figura: "ed",
      calculo: "mín( máx( mín(Inicial + Almacenada, Capacidad), Almacenada ), Reserva total )\nmín( máx( mín({re_eini} + {re_ealm}, {re_capacidad_kwh}), {re_ealm} ), {re_res_total} ) = {re_edisp}",
      resultado: {etiqueta: "Energía Disponible", valor: "{re_edisp} kWh"},
      bloques: [
        {tipo: "nota", casos: [
          {titulo: "Varios ciclos: el paso 3 protege lo cargado", texto: "Capacidad 100, inicial 30. Carga 80, descarga y vuelve a cargar 70: almacenada 150.", cuenta: "mín(180, 100) = 100 → máx(100, 150) = 150 ✓"},
          {titulo: "Donde el tope recorta energía real", texto: "Capacidad 100, inicial 90. Descarga 50 y después carga 20. Pudo entregar 110 sin pasar la capacidad.", cuenta: "mín(110, 100) = 100 → máx(100, 20) = 100 ✗ se pierden 10", malo: true},
        ]},
        {tipo: "revision", titulo: "En revisión (observación 2)", texto: "La fórmula usa solo totales y no mira el orden de cargas y descargas. La RE545, leída al pie de la letra, también limita la energía inicial a la capacidad («siempre que la suma de ambas no supere la capacidad»). Falta además el respaldo normativo de que la ventana 31 valga 0."},
      ],
    },
    {
      id: "horas", titulo: "Ordenar las horas por precio",
      texto: "Se calcula el CMg promedio de cada hora (el promedio de sus 4 bloques) y se ordenan las horas de la más cara a la más barata. Así lo pide la RE545: por CMg <b>horario</b>.",
      figura: "orden_horas",
    },
    {
      id: "curva", titulo: "Abrir cada hora en sus bloques",
      texto: "Cada hora se abre en sus 4 bloques, de mayor a menor CMg. Así queda la <b>curva descendente</b>, con {n_bloques:0} puestos. Desde aquí cada fila es un <b>puesto de la curva</b>, no un bloque en el tiempo.",
      figura: "curva",
      bloques: [{tipo: "revision", titulo: "En revisión (observación 4)", texto: "El código busca el precio de cada puesto entre todas las filas que tengan el mismo CMg promedio de la hora, sin limitarse a la central. Si dos BESS están en la misma barra, sus bloques se mezclan y los precios de la curva salen corridos. Ya está corregido en la macro; en Python está listo pero se aplicará después."}],
    },
    {
      id: "inyeccion", titulo: "Inyección real de cada puesto",
      texto: "A cada puesto se le asigna lo que la batería realmente inyectó en el bloque que lo ocupa.",
      figura: "inyeccion",
    },
    {
      id: "c1", titulo: "Componente 1: llenar las cajas desde la más cara",
      texto: "Cada puesto es una <b>caja de {pmax_bloque:0} kWh</b>, lo máximo que la batería puede inyectar en 15 minutos (Pmax × 1000 ÷ 4). Parte de la caja ya la ocupa la inyección real; el resto es espacio libre. La Energía Disponible se echa en esos espacios, <b>partiendo por la caja más cara</b>, hasta que se acaba.",
      figura: "cajas",
      calculo: "Energía asignada = mín( {pmax_bloque:0} − inyección real , lo que queda de Energía Disponible )\nValor del puesto = energía asignada × CMg del puesto",
      resultado: {etiqueta: "Componente 1", valor: "{re_c1:0}", detalle: "lo que habría ganado vendiendo la energía retenida en los mejores bloques"},
    },
    {
      id: "siguiente", titulo: "Siguen: Componente 2 y monto a compensar", siguiente: true,
      texto: "Pendiente de revisar juntos. El Componente 2 usa la misma Energía Disponible y las mismas cajas, pero se llena desde la caja más barata hacia arriba y solo donde hubo inyección real. El monto es Componente 1 − Componente 2 ({re_c1:0} − {re_c2:0} = {re_monto:0}), y se reparte según la reserva de cada bloque.",
    },
  ],
  pie: [
    'Ver también: <a href="ecostos.html">Paso a paso E Costos</a>.',
    "Código: <code>Script/nucleo/re545.py</code>, <code>re545_reservas.py</code>, <code>re545_resumen.py</code>, <code>re545_componentes.py</code>. Norma: RE545 §2.2.2.2.",
  ],
};
