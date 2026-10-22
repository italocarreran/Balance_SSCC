/* ==================================================================
   explicacion/contenido/ecostos.js — CONTENIDO (lo escribe Claude).
   Solo textos y estructura de pasos. Los números van como {nombre}
   (salen de datos/ejemplo.js). Nada de letras de columnas de Excel.
   ================================================================== */
window.CONTENIDO = {
  hoja: "ecostos",
  eyebrow: "Balance BESS · Calculo E Costos",
  titulo: "Paso a paso E Costos",
  bajada: "Cómo se calcula el costo de oportunidad de un BESS cuando la ventana no tuvo oferta completa (Estudio de Costos SSCC 2024-2027, §10.2.5), en pasos cortos y con un mismo ejemplo. Los números salen de correr el código del repositorio.",
  caso: [
    {valor: "SAE-DEMO · ciclo 1", detalle: "una central, un ciclo"},
    {valor: "{pmax_mw} MW", detalle: "Pmax"},
    {valor: "{n_bloques:0} bloques", detalle: "ventana reducida a 3 horas"},
    {valor: "{mw_cpf} MW CPF(+) · {mw_csf} MW CSF(−)", detalle: "adjudicados las 3 horas · FD {fd}"},
  ],
  nota_inicial: "Cada ciclo de cada central se calcula por separado. La única excepción es la regla de los 10 kWh, que suma todas las centrales (paso 6).",
  pasos: [
    {
      id: "filtrar", titulo: "Filtrar los bloques con subasta y sumar la energía",
      texto: "En orden cronológico, se toman los bloques del ciclo con <b>Adj SSCC = 1</b> (la central tuvo subasta de subida o bajada en esa hora) y se suma por separado lo descargado y lo cargado.",
      figura: "filtro",
      resultado: {etiqueta: "Descarga", valor: "{ec_tot_desc} kWh", detalle: "Carga {ec_tot_carga} kWh"},
      bloques: [{tipo: "nota", titulo: "Las filas con Adj = 0 no se borran.", texto: "No suman energía, pero siguen contando en dos lugares: la regla de los 10 kWh (paso 6) y el reparto final del monto (paso 11)."}],
    },
    {
      id: "pmax", titulo: "Energía máxima de un bloque",
      texto: "Cada central tiene su Pmax en Resumen BESS (<code>Centrales.xlsx</code>). La energía máxima que puede mover en un bloque de 15 minutos es:",
      calculo: "Pmax × 1000 ÷ 4 = {pmax_mw} × 1000 ÷ 4 = {pmax_bloque} kWh",
      resultado: {etiqueta: "Un bloque lleno =", valor: "{pmax_bloque} kWh"},
      bloques: [{tipo: "nota", texto: "Si a la central le falta el Pmax o vale 0, este cálculo queda vacío y el programa avisa."}],
    },
    {
      id: "bloques", titulo: "Cuántos bloques llena esa energía",
      texto: "Se divide la energía total del paso 1 por la energía de un bloque lleno. Todavía no hay precios.",
      calculo: "Descarga: {ec_tot_desc} ÷ {pmax_bloque} = {ec_bloques_desc} bloques\nCarga:    {ec_tot_carga} ÷ {pmax_bloque} = {ec_bloques_carga} bloques",
      figura: "bloques",
    },
    {
      id: "ordenar", titulo: "Ordenar los CMg",
      texto: "En paralelo se ordenan los CMg de los bloques del paso 1 (Adj = 1 y energía ≠ 0). La descarga va de <b>mayor a menor</b> y la carga de <b>menor a mayor</b>. Desde aquí cada fila de la curva es un <b>puesto</b>, no un bloque.",
      figura: "ordenar",
      bloques: [{tipo: "nota", texto: "Ni la prorrata ni el FD se miran para ordenar."}],
    },
    {
      id: "puestos", titulo: "Poner la energía a Pmax en los primeros puestos",
      texto: "Los bloques del paso 3 se ponen en los primeros puestos de la curva: los más caros en la descarga y los más baratos en la carga.",
      figura: "puestos",
      bloques: [{tipo: "nota", texto: "Con la carga es igual, en la curva de menor a mayor."}],
    },
    {
      id: "precio", titulo: "Precio: promedio de los puestos usados",
      texto: "Toda la energía a Pmax se valoriza a un solo precio: el promedio del CMg de los puestos usados.",
      calculo: "Descarga: promedio de {ec_n_usados_desc:0} puestos = {ec_precio_desc}  →  {ec_tot_desc} × {ec_precio_desc} = {ec_ingreso:0}\nCarga:    promedio de {ec_n_usados_carga:0} puestos = {ec_precio_carga}  →  {ec_tot_carga} × {ec_precio_carga} = {ec_costo:0}",
      resultado: {etiqueta: "Ingreso a Pmax", valor: "{ec_ingreso:0}", detalle: "costo a Pmax {ec_costo:0} · todavía sin FD"},
      bloques: [{tipo: "aviso", titulo: "Regla de los 10 kWh (aplica aquí)", texto: "El ingreso solo cuenta si la descarga total del ciclo, sumando <b>todas las centrales y todos los bloques</b> (con o sin Adj), supera 10 kWh. Para el costo de carga, la carga total tiene que ser menor que −10 kWh. Si no se cumple, queda en 0 para todo el ciclo. Se puede revisar antes de ordenar, pero los pasos 3 a 5 igual se necesitan para el descuento por FD (paso 8), que no revisa esta regla. Viene de la macro VBA original; no tiene respaldo en el Estudio de Costos."}],
    },
    {
      id: "fd", titulo: "FD: cuánta energía quedó respaldada",
      texto: "Para cada puesto se calcula la energía respaldada por el desempeño. Lo que falta para llegar a la energía a Pmax es la energía <b>no respaldada</b>.",
      calculo: "Energía con FD = (prorrata CPF × FD CPF + prorrata CSF × FD CSF) × energía a Pmax\nPuesto 1: (0,5 × {fd} + 0,5 × {fd}) × {pmax_bloque} = 2.250  →  no respaldada 250",
      resultado: {etiqueta: "No respaldada: descarga", valor: "{ec_no_resp_desc} kWh", detalle: "carga {ec_no_resp_carga} kWh"},
      figura: "dual",
      bloques: [
        {tipo: "aviso", titulo: "La prorrata y el FD vienen de órdenes distintos", texto: "El FD se busca con el <b>bloque del puesto</b> (orden por precio). La prorrata se toma de la <b>fila</b> (orden cronológico), que es otro bloque. Si las prorratas de esa fila suman 0, el código usa factor 1 y no hay descuento, aunque el bloque del puesto sí tuvo subasta. Casos: FD = 0 con prorrata &gt; 0 → toda la energía queda no respaldada. Las dos prorratas = 0 → factor 1, sin descuento. Una prorrata = 0 → se usa solo la otra."},
        {tipo: "pausa", titulo: "En pausa: cómo debería aplicarse el FD", texto: "La norma no lo define. §10.2.5 (BESS) no menciona FD ni prorrata. La nota 6 (Dictamen 10/2021) dice que el FD se aplica a todas las componentes, sin decir cómo. §10.2.1 lo aplica al CO de cada hora.", opciones: [
          {nombre: "Hoy", texto: "Prorrata de la fila × FD del bloque del puesto. Hereda la mezcla de la macro original."},
          {nombre: "A", texto: "Prorrata y FD del bloque del puesto. Es el arreglo mínimo."},
          {nombre: "B · la preferida", texto: "Castigar la energía real de cada bloque con su propio FD antes de ordenar. Falta definir si también se castiga el Componente 2.", preferida: true},
        ]},
      ],
    },
    {
      id: "descuento", titulo: "Valorizar lo no respaldado",
      texto: "La energía no respaldada se descuenta, pero con <b>otro precio</b>: el promedio de los puestos que caben en <b>horas con subasta × 4</b> y que tienen precio.",
      figura: "precios_descuento",
      calculo: "Límite: {ec_horas_subasta:0} horas con subasta × 4 = {ec_limite_puestos:0} puestos\nDescarga: {ec_no_resp_desc} × {ec_precio_descuento_desc}\nCarga:    {ec_no_resp_carga} × {ec_precio_descuento_carga}",
      resultado: {etiqueta: "Descuento FD", valor: "{ec_descuento:0}"},
      bloques: [
        {tipo: "nota", titulo: "El límite cambia con las horas con subasta.", texto: "Con 1 hora se promedian los 4 primeros puestos (los más caros); con 24 horas, toda la curva. Por eso el kWh no respaldado puede descontarse más barato de lo que se pagó."},
        {tipo: "aviso", titulo: "Umbrales cruzados", texto: "En el código, el límite de la descarga usa las horas de BAJADA y el de la carga, las de SUBIDA. Así está en la macro original. En el ejemplo da lo mismo porque hay 3 horas de cada una."},
      ],
    },
    {
      id: "c1", titulo: "Componente 1",
      texto: "Lo que habría ganado arbitrando a plena potencia, menos lo que no respaldó con buen desempeño.",
      figura: "wf_c1",
      resultado: {etiqueta: "Componente 1 =", valor: "{ec_c1:0}"},
    },
    {
      id: "c2", titulo: "Componente 2",
      texto: "Lo que ganó en realidad, bloque por bloque en orden cronológico y a su propio CMg. Solo cuentan los bloques con Adj = 1.",
      calculo: "Descarga real: Σ descarga × CMg = {ec_c2_desc:0}\nCarga real:    Σ carga × CMg    = {ec_c2_carga:0}",
      figura: "wf_c2",
      resultado: {etiqueta: "Componente 2 =", valor: "{ec_c2:0}"},
    },
    {
      id: "monto", titulo: "Monto a compensar",
      texto: "La diferencia, solo si es positiva, repartida en partes iguales entre <b>todas</b> las filas del ciclo, también las que tienen Adj = 0.",
      calculo: "máx(0, {ec_c1:0} − {ec_c2:0}) = {ec_monto:0}\n{ec_monto:0} ÷ {n_bloques:0} filas = {ec_monto_bloque} por bloque",
      resultado: {etiqueta: "Monto a compensar", valor: "{ec_monto:0} $", detalle: "{ec_monto_bloque} en cada bloque del ciclo"},
      bloques: [{tipo: "nota", texto: "Como el reparto es parejo, un bloque sin subasta o con FD = 0 recibe la misma parte que el resto. Después, PRORRATA_RETIROS cobra ese monto a quienes retiraron en cada bloque."}],
    },
  ],
  pie: [
    'Ver también: <a href="re545.html">Paso a paso RE545</a>.',
    "Código: <code>Script/nucleo/columnas_compartidas.py</code>, <code>ecostos_columnas.py</code>, <code>ecostos_prorratas.py</code>, <code>ecostos_ciclo.py</code>. Norma: Estudio de Costos SSCC 2024-2027 §10.2.5 y §10.2.1.",
  ],
};
