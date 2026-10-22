/* ==================================================================
   explicacion/figuras.js — DISEÑO: gráficos y tablas de las páginas.
   Cada figura es window.FIGURAS[nombre](contenedor, EJEMPLO) y dibuja
   solo con los datos de EJEMPLO (datos/ejemplo.js). Los textos que
   explican la figura viven en contenido/*.js, no aquí.
   ================================================================== */
(function () {
  const fmt = (v, d) => window.FMT(v, d);
  const sum = a => a.reduce((s, v) => s + (typeof v === "number" ? v : 0), 0);
  const avg = a => a.length ? sum(a) / a.length : 0;
  const NS = "http://www.w3.org/2000/svg";
  // fill/stroke con var(...) van como style para que respeten el tema
  const el = (tag, attrs = {}, parent) => {
    const e = document.createElementNS(NS, tag);
    for (const k in attrs) {
      const v = String(attrs[k]);
      if ((k === "fill" || k === "stroke") && v.startsWith("var(")) e.style[k] = v; else e.setAttribute(k, v);
    }
    if (parent) parent.appendChild(e);
    return e;
  };
  const txt = (p, x, y, s, a = {}) => { const t = el("text", {x, y, ...a}, p); t.textContent = s; return t; };
  const MONO = "font-family:var(--f-mono);font-size:10.5px";
  const svg = (cont, w, h, label, clase) => {
    const s = el("svg", {viewBox: `0 0 ${w} ${h}`, role: "img", "aria-label": label});
    if (clase) s.setAttribute("class", clase);
    cont.appendChild(s);
    return s;
  };
  const tabla = (cont, cabeza, filas, clase = "tbl") => {
    cont.insertAdjacentHTML("beforeend", `<div class="${clase}"><table><thead><tr>${cabeza.map(c => `<th>${c}</th>`).join("")}</tr></thead><tbody>${filas.join("")}</tbody></table></div>`);
  };
  const leyenda = (cont, items) => cont.insertAdjacentHTML("beforeend",
    `<div class="legend">${items.map(([c, t, borde]) => `<span><i class="sw" style="${borde ? "border:1px dashed " + c : "background:" + c}"></i>${t}</span>`).join("")}</div>`);

  // Dos filas de barras: arriba el orden de la curva, abajo el orden en el tiempo
  const ordenBarras = (s, cmg, orden, marcados, titulo, subtitulo, colorOn) => {
    const L = 8, W = 384, n = cmg.length, bw = W / n, h = 34, max = Math.max(...cmg) * 1.04;
    const fila = (y0, idxs, t, abajo) => {
      txt(s, L, y0 - 4, t, {style: "font-size:11px;font-weight:600"});
      idxs.forEach((b, k) => {
        const x = L + k * bw, hh = cmg[b] / max * h, on = marcados.has(b);
        const r = el("rect", {x: x + 2, y: y0 + h - hh, width: bw - 4, height: hh, fill: on ? colorOn : "var(--real)"}, s);
        if (abajo) r.setAttribute("opacity", on ? "0.6" : "0.3");
        txt(s, x + bw / 2, y0 + h + 11, String(b + 1), {"text-anchor": "middle", style: "font-family:var(--f-mono);font-size:9.5px"});
      });
    };
    fila(16, orden, titulo, false);
    txt(s, L + W / 2, 86, "▲ " + subtitulo, {"text-anchor": "middle", style: "font-family:var(--f-mono);font-size:10px;fill:var(--muted)"});
    fila(112, [...Array(n).keys()], "en el tiempo: bloques 1 a " + n, true);
  };

  // Cascada horizontal: pasos que suman o restan y una barra de total
  const cascada = (s, pasos) => {
    const W = 520, rowH = 28, L = 140, R = W - 90;
    let run = 0;
    const filas = pasos.map(p => { if (p.total) { run = p.v; return {p, a: 0, b: p.v}; } const r = {p, a: run, b: run + p.v}; run += p.v; return r; });
    const max = Math.max(...filas.map(r => Math.max(r.a, r.b))), x = v => L + v / max * (R - L);
    s.setAttribute("viewBox", `0 0 ${W} ${filas.length * rowH + 6}`);
    filas.forEach((r, i) => {
      const y = 4 + i * rowH;
      txt(s, L - 8, y + 17, r.p.l, {"text-anchor": "end", style: "font-size:12.5px" + (r.p.total ? ";font-weight:700" : "")});
      el("rect", {x: x(Math.min(r.a, r.b)), y: y + 4, width: Math.max(Math.abs(x(r.b) - x(r.a)), 1), height: rowH - 10,
        fill: r.p.total ? r.p.c : (r.p.v >= 0 ? "var(--ok)" : "var(--warn)")}, s);
      txt(s, W - 4, y + 17, (r.p.v > 0 && !r.p.total ? "+" : "") + fmt(Math.round(r.p.v)), {"text-anchor": "end", style: "font-family:var(--f-mono);font-size:11.5px" + (r.p.total ? ";font-weight:700" : "")});
    });
  };

  /* ---------------- E Costos ---------------- */
  const ec = E => {
    const C = E.ec, blk = q => C["Cuarto de Hora"].indexOf(q);
    const posD = C.AB.map((p, i) => ({p, b: blk(C.Y[i]), e: C.AE[i]})).filter(o => o.p !== null);
    const posC = C.AD.map((p, i) => ({p, b: blk(C.AC[i]), e: -C.AF[i]})).filter(o => o.p !== null);
    return {C, posD, posC};
  };

  const F = {};

  F.filtro = (cont, E) => {
    const C = E.ec, V = E.valores;
    tabla(cont, ["Bloque", "Hora", "Adj SSCC", "Descarga kWh", "Carga kWh"],
      C.CMg.map((_, i) => `<tr class="${C.L[i] ? "" : "off"}"><td>${i + 1}</td><td>${C.Hora[i]}</td><td>${C.L[i]}</td><td>${fmt(C.Energia_Positiva[i])}</td><td>${fmt(-C.Energia_Negativa[i])}</td></tr>`)
        .concat(`<tr class="tot"><td>Total con Adj = 1</td><td></td><td></td><td>${fmt(V.ec_tot_desc)}</td><td>${fmt(V.ec_tot_carga)}</td></tr>`));
  };

  F.bloques = (cont, E) => {
    const V = E.valores, P = V.pmax_bloque, W = 120, partes = [];
    for (let r = V.ec_tot_desc; r > 0; r -= P) partes.push(Math.min(r, P));
    const s = svg(cont, 520, 70, partes.map(fmt).join(" + ") + " kWh", "fig-chica");
    partes.forEach((v, k) => {
      const x = 10 + k * (W + 10);
      el("rect", {x, y: 14, width: W, height: 30, fill: "none", stroke: "var(--line)", "stroke-dasharray": "4 3"}, s);
      el("rect", {x, y: 14, width: W * v / P, height: 30, fill: "var(--ec)"}, s);
      txt(s, x + W / 2, 62, `bloque ${k + 1}: ${fmt(v)}`, {"text-anchor": "middle", style: MONO});
    });
    txt(s, 10 + partes.length * (W + 10), 34, "= " + fmt(V.ec_tot_desc) + " kWh", {style: "font-family:var(--f-mono);font-size:12px;font-weight:600"});
  };

  F.ordenar = (cont, E) => {
    const {C, posD, posC} = ec(E);
    cont.insertAdjacentHTML("beforeend", `<div class="dos"><div><b>Descarga</b></div><div><b>Carga</b></div></div>`);
    const [d, c] = cont.querySelectorAll(".dos > div");
    ordenBarras(svg(d, 400, 168, "Bloques con descarga ordenados de mayor a menor CMg"), C.CMg, posD.map(o => o.b), new Set(posD.map(o => o.b)), "ordenado: de mayor a menor", "se ordenan por CMg", "var(--ec)");
    ordenBarras(svg(c, 400, 168, "Bloques con carga ordenados de menor a mayor CMg"), C.CMg, posC.map(o => o.b), new Set(posC.map(o => o.b)), "ordenado: de menor a mayor", "se ordenan por CMg", "var(--ec)");
  };

  F.puestos = (cont, E) => {
    const {C, posD} = ec(E);
    tabla(cont, ["Puesto", "Bloque", "CMg ordenado", "Descarga real", "Descarga a Pmax"],
      posD.map((o, k) => `<tr><td>${k + 1}</td><td>${o.b + 1}</td><td>${fmt(o.p)}</td><td>${fmt(C.Energia_Positiva[o.b])}</td><td class="${o.e > 0 ? "marca" : ""}">${fmt(o.e)}</td></tr>`));
  };

  F.precios_descuento = (cont, E) => {
    const V = E.valores, s = svg(cont, 520, 96, "Precio del kWh respaldado contra precio del descuento", "fig-chica");
    const L = 170, R = 400, max = Math.max(V.ec_precio_desc, V.ec_precio_descuento_desc) * 1.1, x = v => L + v / max * (R - L);
    [["kWh respaldado", V.ec_precio_desc, "var(--ec)"], ["kWh no respaldado (descuento)", V.ec_precio_descuento_desc, "var(--warn)"]].forEach(([t, v, c], k) => {
      const y = 12 + k * 40;
      txt(s, L - 8, y + 16, t, {"text-anchor": "end", style: "font-size:12px"});
      el("rect", {x: L, y, width: x(v) - L, height: 22, fill: c}, s);
      txt(s, x(v) + 6, y + 16, fmt(v) + " $/kWh", {style: "font-family:var(--f-mono);font-size:11.5px;font-weight:600"});
    });
  };

  // Tabla doble interactiva: prorrata (orden cronológico) vs FD (orden por precio)
  F.dual = (cont, E) => {
    const escenarios = {base: E.ec, sin1: E.ec_sin_subasta_hora1};
    let esc = "base", fila = 0;
    cont.insertAdjacentHTML("beforeend", `<div class="tbl dual"></div>
      <div style="display:flex;gap:10px;align-items:center;flex-wrap:wrap"><span class="hint">Escenario:</span>
        <div class="seg" role="group" aria-label="Escenario"><button type="button" data-e="base" aria-pressed="true">Ejemplo base</button><button type="button" data-e="sin1" aria-pressed="false">Si la hora 1 no hubiera tenido subasta</button></div></div>
      <div class="why" aria-live="polite"></div>`);
    const tbl = cont.querySelector(".tbl.dual"), why = cont.querySelector(".why");
    const draw = () => {
      const D = escenarios[esc], b = q => D["Cuarto de Hora"].indexOf(q), pb = D.Y.map(b), home = pb[fila];
      const head = `<thead><tr class="grp"><th class="c" colspan="4">Cronológico · fila = bloque</th><th class="sep"></th><th class="p" colspan="5">Por precio · fila = puesto</th><th class="sep"></th><th>Mezcla</th></tr>
        <tr><th class="c">Bloque</th><th class="c">Hora</th><th class="c">Prorr. CPF</th><th class="c">Prorr. CSF</th><th class="sep"></th><th class="p">Puesto</th><th class="p">Bloque del puesto</th><th class="p">CMg ordenado</th><th class="p">FD</th><th class="p">Energía a Pmax</th><th class="sep"></th><th>Energía con FD</th></tr></thead>`;
      const rows = D.CMg.map((_, i) => {
        const on = i === fila, u = on ? " use" : "", z = v => v === 0 ? " zero" : "", conPrecio = D.AB[i] !== null;
        return `<tr data-i="${i}" class="${on ? "on" : ""}"><td class="c${i === home ? " home" : ""}">${i + 1}</td><td class="c">${D.Hora[i]}</td><td class="c${u}${z(D.AG[i])}">${fmt(D.AG[i])}</td><td class="c${u}${z(D.AH[i])}">${fmt(D.AH[i])}</td><td class="sep"></td>
          <td class="p">${i + 1}</td><td class="p">${conPrecio ? pb[i] + 1 : "—"}</td><td class="p">${conPrecio ? fmt(D.AB[i]) : "—"}</td><td class="p${u}">${fmt(D.AM[i])}</td><td class="p${u}">${fmt(D.AE[i])}</td><td class="sep"></td><td>${fmt(D.AS[i])}</td></tr>`;
      }).join("");
      tbl.innerHTML = `<table>${head}<tbody>${rows}</tbody></table>`;
      tbl.querySelectorAll("tbody tr").forEach(tr => tr.addEventListener("click", () => { fila = +tr.dataset.i; draw(); }));
      const i = fila, a = D.AG[i], c = D.AH[i], e = D.AE[i], factor1 = a + c === 0, conPrecio = D.AB[i] !== null;
      why.innerHTML = `<b>Fila ${i + 1}</b><span>Prorrata de la fila: bloque ${i + 1} (hora ${D.Hora[i]}). FD del puesto ${i + 1}: ${conPrecio ? `bloque ${pb[i] + 1} (borde punteado)` : "sin bloque"}.</span>
        <span class="eq">Energía con FD = ${factor1 ? `factor 1 × ${fmt(e)}` : `(${fmt(a)} × ${fmt(D.AM[i])} + ${fmt(c)} × ${fmt(D.AN[i])}) × ${fmt(e)}`} = ${fmt(D.AS[i])}</span>
        ${factor1 && e > 0 ? `<span style="color:var(--warn)">Sin descuento: la hora ${D.Hora[i]} no tuvo subasta, aunque el bloque ${pb[i] + 1} sí.</span>` : ""}`;
    };
    cont.querySelectorAll(".seg button").forEach(bt => bt.addEventListener("click", () => {
      esc = bt.dataset.e;
      cont.querySelectorAll(".seg button").forEach(o => o.setAttribute("aria-pressed", o === bt));
      draw();
    }));
    draw();
  };

  F.wf_c1 = (cont, E) => {
    const V = E.valores;
    cascada(svg(cont, 520, 120, "Componente 1 = descarga a Pmax + carga a Pmax − descuento FD", "fig-media"), [
      {l: "Descarga a Pmax", v: V.ec_ingreso}, {l: "Carga a Pmax", v: V.ec_costo},
      {l: "Descuento FD", v: -V.ec_descuento}, {l: "Componente 1", v: V.ec_c1, total: true, c: "var(--ec)"}]);
  };

  F.wf_c2 = (cont, E) => {
    const V = E.valores;
    cascada(svg(cont, 520, 90, "Componente 2 = descarga real + carga real", "fig-media"), [
      {l: "Descarga real", v: V.ec_c2_desc}, {l: "Carga real", v: V.ec_c2_carga},
      {l: "Componente 2", v: V.ec_c2, total: true, c: "var(--real)"}]);
  };

  /* ---------------- RE545 ---------------- */
  // Bloque (en el tiempo) que ocupa cada puesto de la curva
  const ordenCurva = R => R.BI.map((bi, i) => R.S.map((s, j) => s === bi ? j : -1).filter(j => j >= 0)
    .sort((a, b) => R.CMg[b] - R.CMg[a])[R.BJ[i] / 15]);
  const etiqueta = (R, i) => `H${R.Hora[i]} :${String(R.Minutos[i]).padStart(2, "0")}`;

  F.ed = (cont, E) => {
    const V = E.valores, s = svg(cont, 520, 132, "Inicial más almacenada, reserva total y Energía Disponible", "fig-media");
    const L = 10, Rr = 440, max = Math.max(V.re_eini_mas_ealm, V.re_res_total), x = v => L + v / max * (Rr - L);
    [["inicial + almacenada", [[V.re_eini, "var(--real)"], [V.re_ealm, "var(--in)"]], V.re_eini_mas_ealm, false],
     ["reserva total", [[V.re_res_total, "var(--muted)"]], V.re_res_total, false],
     ["Energía Disponible: la menor", [[V.re_edisp, "var(--re)"]], V.re_edisp, true]].forEach(([t, partes, tot, fuerte], k) => {
      const y = 6 + k * 42;
      txt(s, L, y + 11, t, {style: "font-size:12px" + (fuerte ? ";font-weight:700" : "")});
      let acc = 0;
      partes.forEach(([v, c]) => { el("rect", {x: x(acc), y: y + 16, width: x(acc + v) - x(acc), height: 16, fill: c}, s); acc += v; });
      txt(s, x(tot) + 6, y + 29, fmt(tot) + " kWh", {style: MONO + (fuerte ? ";font-weight:700" : "")});
    });
    leyenda(cont, [["var(--real)", "inicial"], ["var(--in)", "almacenada"], ["var(--muted)", "reserva total"], ["var(--re)", "Energía Disponible"]]);
  };

  F.orden_horas = (cont, E) => {
    const R = E.re, horas = [...new Set(R.Hora)].sort((a, b) => R.S[R.Hora.indexOf(a)] - R.S[R.Hora.indexOf(b)]);
    tabla(cont, ["Hora", "CMg de sus bloques", "CMg promedio", "Puesto de la hora"], horas.map(h => {
      const idx = R.Hora.map((x, i) => x === h ? i : -1).filter(i => i >= 0);
      return `<tr><td>${h}</td><td>${idx.map(i => fmt(R.CMg[i])).join(", ")}</td><td class="marca">${fmt(R.R[idx[0]])}</td><td>${R.S[idx[0]]}</td></tr>`;
    }));
  };

  F.curva = (cont, E) => {
    const R = E.re, orden = ordenCurva(R);
    ordenBarras(svg(cont, 400, 168, "Los bloques en el tiempo y, arriba, ordenados por hora y luego por bloque", "fig-chica"),
      R.CMg, orden, new Set(orden), "curva: horas por CMg promedio, bloques por CMg", "se ordena", "var(--re)");
    tabla(cont, ["Puesto", "Hora", "Bloque", "CMg del puesto"],
      orden.map((b, k) => `<tr><td>${k + 1}</td><td>${R.Hora[b]}</td><td>${b + 1} (${etiqueta(R, b)})</td><td class="marca">${fmt(R.BM[k])}</td></tr>`));
  };

  F.inyeccion = (cont, E) => {
    const R = E.re, orden = ordenCurva(R);
    tabla(cont, ["Puesto", "Hora", "Bloque", "Inyección real kWh"],
      orden.map((b, k) => `<tr><td>${k + 1}</td><td>${R.Hora[b]}</td><td>${b + 1}</td><td>${fmt(R.BS[k])}</td></tr>`));
  };

  // Cajas de Pmax por puesto: real en gris y energía asignada en color
  const cajas = (cont, E, asignada, valor, aria) => {
    const R = E.re, P = E.valores.pmax_bloque, n = R.BM.length;
    const s = svg(cont, 520, 210, aria, "fig-media");
    const L = 46, Rr = 512, bw = (Rr - L) / n, top = 14, bot = 160, y = v => bot - v / (P * 1.04) * (bot - top);
    [0, P].forEach(v => txt(s, L - 6, y(v) + 4, fmt(v), {"text-anchor": "end", style: MONO}));
    R.BM.forEach((p, k) => {
      const x = L + k * bw + 4, w = bw - 8, real = R.BS[k], a = asignada[k];
      el("rect", {x, y: y(P), width: w, height: bot - y(P), fill: "none", stroke: "var(--muted)", "stroke-dasharray": "3 3"}, s);
      if (real > 0) el("rect", {x, y: y(real), width: w, height: bot - y(real), fill: "var(--real)"}, s);
      if (a > 0) el("rect", {x, y: y(real + a), width: w, height: y(real) - y(real + a), fill: "var(--re)"}, s);
      txt(s, x + w / 2, bot + 15, fmt(p), {"text-anchor": "middle", style: MONO});
      txt(s, x + w / 2, bot + 29, "p" + (k + 1), {"text-anchor": "middle", style: "font-family:var(--f-mono);font-size:9.5px;fill:var(--muted)"});
    });
    txt(s, L, bot + 45, "CMg del puesto, de la caja más cara a la más barata →", {style: MONO + ";fill:var(--muted)"});
    leyenda(cont, [["var(--real)", "inyección real"], ["var(--re)", "Energía Disponible asignada"], ["var(--muted)", "caja de " + fmt(P) + " kWh", true]]);
    tabla(cont, ["Puesto", "CMg", "Real", "Libre", "Asignada", "Valor"],
      R.BM.map((p, k) => `<tr><td>${k + 1}</td><td>${fmt(p)}</td><td>${fmt(R.BS[k])}</td><td>${fmt(P - R.BS[k])}</td><td class="${asignada[k] ? "marca" : ""}">${fmt(asignada[k])}</td><td>${fmt(valor[k])}</td></tr>`)
        .concat(`<tr class="tot"><td>Total</td><td></td><td></td><td></td><td>${fmt(sum(asignada))}</td><td>${fmt(sum(valor))}</td></tr>`));
  };

  F.cajas = (cont, E) => cajas(cont, E, E.re.BN, E.re.BO, "Cajas de Pmax llenadas con la Energía Disponible desde el puesto más caro");

  window.FIGURAS = F;
  window.FIG_UTIL = {el, txt, svg, tabla, leyenda, sum, avg};
})();
