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
  nota_inicial: "Cada ciclo de cada central se calcula por separado. La única excepción es la regla de los 10 kWh, que suma todas las centrales (paso 7). Método del FD vigente desde septiembre 2026: el FD se aplica a la energía de cada hora antes de valorizar.",
  pasos: [
    {
      id: "filtrar", titulo: "Filtrar los bloques con subasta",
      texto: "En orden cronológico, se toman los bloques del ciclo con <b>Adj SSCC = 1</b> (la central tuvo subasta de subida o bajada en esa hora). Se mira por separado lo descargado y lo cargado.",
      figura: "filtro",
      resultado: {etiqueta: "Descarga medida", valor: "{ec_tot_desc} kWh", detalle: "carga medida {ec_tot_carga} kWh"},
      bloques: [{tipo: "nota", titulo: "Las filas con Adj = 0 no se borran.", texto: "No aportan energía, pero siguen contando en dos lugares: la regla de los 10 kWh (paso 7) y el reparto final del monto (paso 9)."}],
    },
    {
      id: "fd", titulo: "Ajustar la energía de cada hora por FD",
      texto: "Antes de todo lo demás, la energía de cada bloque se multiplica por el <b>factor de su propia hora</b>: cuánto de esa hora estuvo respaldado con buen desempeño.",
      calculo: "factor de la hora = Σ (prorrata del servicio × FD del servicio)\n           = {ec_prorr_cpf_mas} × {ec_fd_cpf_mas} (CPF+) + {ec_prorr_csf_menos} × {ec_fd_csf_menos} (CSF−) = {ec_factor}\nenergía ajustada = energía medida × factor",
      partes: [
        "<b>Prorrata:</b> la participación de cada servicio adjudicado en la hora, con su dirección: CPF(−), CSF(−), CPF(+), CSF(+). Suman 100 %.",
        "<b>FD:</b> el de cada servicio en esa misma hora. Solo cuenta si su prorrata no es 0.",
        "El mismo factor se aplica a la descarga y a la carga del bloque.",
        "Las horas sin subasta (Adj = 0) quedan con energía ajustada 0.",
      ],
      figura: "ajuste",
      resultado: {etiqueta: "Descarga ajustada", valor: "{ec_aj_desc} kWh", detalle: "carga ajustada {ec_aj_carga} kWh"},
    },
    {
      id: "pmax", titulo: "Energía máxima de un bloque",
      texto: "Cada central tiene su Pmax en Resumen BESS (<code>Centrales.xlsx</code>). La energía máxima que puede mover en un bloque de 15 minutos es:",
      calculo: "Pmax × 1000 ÷ 4 = {pmax_mw} × 1000 ÷ 4 = {pmax_bloque} kWh",
      resultado: {etiqueta: "Un bloque lleno =", valor: "{pmax_bloque} kWh"},
      bloques: [{tipo: "nota", texto: "Si a la central le falta el Pmax o vale 0, la asignación queda en 0 y el programa avisa."}],
    },
    {
      id: "bloques", titulo: "Cuántos bloques llena la energía ajustada",
      texto: "Se divide la energía ajustada del paso 2 por la energía de un bloque lleno. Todavía no hay precios.",
      calculo: "Descarga: {ec_aj_desc} ÷ {pmax_bloque} = {ec_bloques_desc} bloques\nCarga:    {ec_aj_carga} ÷ {pmax_bloque} = {ec_bloques_carga} bloques",
      figura: "bloques",
    },
    {
      id: "ordenar", titulo: "Ordenar los CMg",
      texto: "Se ordenan los CMg de los bloques con Adj = 1 y energía medida distinta de 0. La descarga va de <b>mayor a menor</b> y la carga de <b>menor a mayor</b>. Desde aquí cada fila de la curva es un <b>puesto</b>, no un bloque.",
      figura: "ordenar",
      bloques: [{tipo: "nota", texto: "La curva es la de siempre: no se vuelve a ordenar con la energía ajustada."}],
    },
    {
      id: "puestos", titulo: "Poner la energía ajustada a Pmax en los primeros puestos",
      texto: "Los bloques del paso 4 se ponen en los primeros puestos de la curva: los más caros en la descarga y los más baratos en la carga.",
      figura: "puestos",
      bloques: [{tipo: "nota", texto: "Con la carga es igual, en la curva de menor a mayor."}],
    },
    {
      id: "precio", titulo: "Componente 1: valorizar al promedio de los puestos usados",
      texto: "Toda la energía ajustada se valoriza a un solo precio: el promedio del CMg de los puestos usados. La suma de descarga y carga es el Componente 1. Ya <b>no hay descuento FD</b> aparte: el FD quedó dentro de la energía en el paso 2.",
      calculo: "Descarga: promedio de {ec_n_usados_desc:0} puestos = {ec_precio_desc}  →  {ec_aj_desc} × {ec_precio_desc} = {ec_val_desc:0}\nCarga:    promedio de {ec_n_usados_carga:0} puestos = {ec_precio_carga}  →  {ec_aj_carga} × {ec_precio_carga} = {ec_val_carga:0}",
      figura: "wf_c1",
      resultado: {etiqueta: "Componente 1 =", valor: "{ec_c1:0}"},
      bloques: [{tipo: "aviso", titulo: "Regla de los 10 kWh (aplica aquí)", texto: "La descarga solo se valoriza si la descarga medida total del ciclo, sumando <b>todas las centrales y todos los bloques</b> (con o sin Adj), supera 10 kWh. La carga, si la carga total es menor que −10 kWh. Si no se cumple, queda en 0 para todo el ciclo. Viene de la macro VBA original; no tiene respaldo en el Estudio de Costos."}],
    },
    {
      id: "c2", titulo: "Componente 2",
      texto: "Lo que ganó en realidad, bloque por bloque en orden cronológico y a su propio CMg, con la energía <b>medida</b> (sin FD). Solo cuentan los bloques con Adj = 1.",
      calculo: "Descarga real: Σ descarga × CMg = {ec_c2_desc:0}\nCarga real:    Σ carga × CMg    = {ec_c2_carga:0}",
      figura: "wf_c2",
      resultado: {etiqueta: "Componente 2 =", valor: "{ec_c2:0}"},
    },
    {
      id: "monto", titulo: "Monto a compensar",
      texto: "La diferencia, solo si es positiva, repartida en partes iguales entre <b>todas</b> las filas del ciclo, también las que tienen Adj = 0.",
      calculo: "máx(0, {ec_c1:0} − {ec_c2:0}) = {ec_monto:0}\n{ec_monto:0} ÷ {n_bloques:0} filas = {ec_monto_bloque} por bloque",
      resultado: {etiqueta: "Monto a compensar", valor: "{ec_monto:0} $", detalle: "{ec_monto_bloque} en cada bloque del ciclo"},
      bloques: [
        {tipo: "nota", texto: "Como el reparto es parejo, un bloque sin subasta recibe la misma parte que el resto. Después, PRORRATA_RETIROS cobra ese monto a quienes retiraron en cada bloque."},
        {tipo: "nota", titulo: "Qué cambió respecto del método anterior", texto: "Antes la energía a Pmax se valorizaba sin FD y después se restaba un \"descuento FD\" a otro precio (promedio de los puestos que cabían en horas con subasta × 4, con umbrales de subida y bajada cruzados). La prorrata no tenía dirección y el FD se buscaba en el bloque del puesto de la curva, no en la hora de la energía. Ahora prorrata y FD son los de la propia hora, se aplican a la energía medida y no hay descuento aparte. El método anterior quedó en la rama Definitivo-Agosto."},
      ],
    },
  ],
  pie: [
    'Ver también: <a href="re545.html">Paso a paso RE545</a>.',
    "Código: <code>Script/nucleo/columnas_compartidas.py</code>, <code>ecostos_columnas.py</code>, <code>ecostos_prorratas.py</code>, <code>ecostos_valorizacion.py</code>. Norma: Estudio de Costos SSCC 2024-2027 §10.2.5; FD según el Dictamen 10/2021 del Panel de Expertos.",
  ],
};
