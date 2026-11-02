/* ==================================================================
   explicacion/pagina.js — DISEÑO: arma la página a partir de
   window.CONTENIDO (contenido/*.js) y window.EJEMPLO (datos/ejemplo.js).
   Quien diseña puede cambiar cómo se dibuja; no cambia el contenido.

   Contrato que respeta (ver README.md):
   - CONTENIDO.pasos[i] = {id, titulo, texto, calculo, resultado, figura,
     bloques, siguiente}
   - En cualquier texto, {nombre} se reemplaza por EJEMPLO.valores.nombre
     con formato chileno; {nombre:0} lo deja sin decimales.
   - figura = nombre de una función de window.FIGURAS (figuras.js).
   ================================================================== */
(function () {
  const C = window.CONTENIDO, V = window.EJEMPLO.valores;
  const fmt = (v, d = 2) => (v === null || v === undefined || Number.isNaN(v)) ? "—"
    : (typeof v === "number" ? v.toLocaleString("es-CL", {maximumFractionDigits: d}) : String(v));
  const rellenar = s => (s || "").replace(/\{(\w+)(?::(\d))?\}/g, (m, k, d) =>
    (k in V) ? fmt(V[k], d === undefined ? 2 : +d) : m);
  window.FMT = fmt;

  document.title = C.titulo;
  document.body.classList.add("hoja-" + C.hoja);
  const app = document.getElementById("app");

  const caso = (C.caso || []).map(f => `<div class="fact"><b>${rellenar(f.valor)}</b><span>${rellenar(f.detalle)}</span></div>`).join("");
  const toc = C.pasos.map((p, k) => `<a href="#${p.id}">${p.siguiente ? "→" : (k + 1) + "."} ${p.titulo}</a>`).join("");

  const bloque = b => {
    let extra = "";
    if (b.opciones) extra = `<div class="opciones">${b.opciones.map(o => `<div class="${o.preferida ? "preferida" : ""}"><b>${o.nombre}</b><span>${rellenar(o.texto)}</span></div>`).join("")}</div>`;
    if (b.casos) extra = `<div class="casos">${b.casos.map(c => `<div class="caso ${c.malo ? "malo" : ""}"><b>${c.titulo}</b><span>${rellenar(c.texto)}</span><span class="cuenta">${rellenar(c.cuenta)}</span></div>`).join("")}</div>`;
    return `<div class="bloque ${b.tipo}">${b.titulo ? `<b>${b.titulo}</b>` : ""}${b.texto ? `<span>${rellenar(b.texto)}</span>` : ""}${extra}</div>`;
  };

  const paso = (p, k) => `
    <section class="step" id="${p.id}">
      <span class="num ${p.siguiente ? "siguiente" : ""}">${p.siguiente ? "→" : k + 1}</span>
      <div class="body">
        <h2>${p.titulo}</h2>
        ${p.texto ? `<p class="q">${rellenar(p.texto)}</p>` : ""}
        ${p.partes ? `<ul class="partes">${p.partes.map((t, i) => `<li><span>${i + 1}</span><span>${rellenar(t)}</span></li>`).join("")}</ul>` : ""}
        ${p.figura ? `<div class="figura" data-figura="${p.figura}"></div>` : ""}
        ${p.calculo ? `<div class="calc">${rellenar(p.calculo)}</div>` : ""}
        ${p.resultado ? `<div class="result">${rellenar(p.resultado.etiqueta)} <b>${rellenar(p.resultado.valor)}</b>${p.resultado.detalle ? ` <span>${rellenar(p.resultado.detalle)}</span>` : ""}</div>` : ""}
        ${(p.bloques || []).map(bloque).join("")}
      </div>
    </section>`;

  app.innerHTML = `
    <header>
      <span class="eyebrow">${C.eyebrow}</span>
      <h1>${C.titulo}</h1>
      <p class="lede">${rellenar(C.bajada)}</p>
      <div class="facts">${caso}</div>
      <nav class="toc" aria-label="Pasos">${toc}</nav>
      ${C.nota_inicial ? `<p class="hint">${rellenar(C.nota_inicial)}</p>` : ""}
    </header>
    <div class="steps">${C.pasos.map(paso).join("")}</div>
    <footer>${(C.pie || []).map(t => `<p>${t}</p>`).join("")}</footer>`;

  document.querySelectorAll("[data-figura]").forEach(el => {
    const f = window.FIGURAS && window.FIGURAS[el.dataset.figura];
    if (f) f(el, window.EJEMPLO); else el.innerHTML = `<p class="hint">Falta la figura «${el.dataset.figura}» en figuras.js</p>`;
  });
})();
