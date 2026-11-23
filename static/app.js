/* CH Ventas — tablero de ventas netas (ventas − devoluciones). */
(() => {
  "use strict";

  const MESES = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio",
    "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"];
  const MES_CORTO = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"];
  // Color por posición del año (del más reciente al más antiguo).
  const COLORES = ["#2b7bd6", "#f0683a", "#34a37a", "#9b7bd1"];
  const COLORES_CLAROS = ["#8fb8ea", "#f7b39b", "#9ad3bd", "#cbbbe9"];
  const SONDEO_MS = 30000;

  const $ = (id) => document.getElementById(id);
  const nf = new Intl.NumberFormat("es-CO", { maximumFractionDigits: 0 });
  const dinero = (n) => "$ " + nf.format(Math.round(n || 0));
  const pct = (p, signo = true) => {
    if (p === null || !isFinite(p)) return "—";
    const s = (p * 100).toFixed(1).replace(".", ",");
    return (signo && p > 0 ? "+" : "") + s + " %";
  };
  const variacion = (a, b) => (b ? (a - b) / Math.abs(b) : null);
  const clase = (n) => (n < 0 ? "neg" : n > 0 ? "pos" : "");
  const esc = (s) => String(s).replace(/[&<>"']/g, (c) =>
    ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  const ultimoDia = (a, m) => new Date(a, m, 0).getDate(); // m 1..12
  const fechaNum = (a, m, d) => a * 10000 + m * 100 + d;
  const mesAnterior = (a, m, n = 1) => {
    let mm = m - n, aa = a;
    while (mm < 1) { mm += 12; aa -= 1; }
    return [aa, mm];
  };

  const estado = {
    datos: null,
    dias: null, cli: null, neto: null, dev: null,
    mensual: new Map(),           // "aaaa-m" -> {neto, dev}
    corte: null,                  // {a, m, d, num}
    anios: [],                    // años seleccionados
    hastaMes: 12,
    tab: "acum",
    grafica: null,
    firma: null,
    msCarga: 0,
    hora: null,
  };

  // ------------------------------------------------------------------
  // Datos
  // ------------------------------------------------------------------
  async function cargarDatos(silencioso = false) {
    const t0 = performance.now();
    const btn = $("btnGenerar");
    btn.disabled = true;
    try {
      const r = await fetch("api/datos", { cache: "no-store" });
      const j = await r.json();
      if (!r.ok || j.error) throw new Error(j.error || `Error ${r.status}`);
      estado.msCarga = Math.round(performance.now() - t0);
      estado.hora = new Date();
      prepararDatos(j);
      $("error").hidden = true;
      $("cargando").hidden = true;
      render();
      if (silencioso) toast("Datos actualizados desde el Excel");
    } catch (e) {
      $("cargando").hidden = true;
      const err = $("error");
      err.hidden = false;
      err.textContent = "No se pudo leer la base de datos.\n" + e.message;
    } finally {
      btn.disabled = false;
    }
  }

  function prepararDatos(j) {
    const anterior = estado.datos;
    estado.datos = j;
    estado.firma = j.firma;
    const n = j.registros.length;
    estado.dias = new Int32Array(n);
    estado.cli = new Int32Array(n);
    estado.neto = new Float64Array(n);
    estado.dev = new Float64Array(n);
    estado.mensual = new Map();
    for (let i = 0; i < n; i++) {
      const [dia, c, venta, dev] = j.registros[i];
      estado.dias[i] = dia;
      estado.cli[i] = c;
      estado.neto[i] = venta - dev;
      estado.dev[i] = dev;
      const k = Math.floor(dia / 100);
      const acc = estado.mensual.get(k) || { neto: 0, dev: 0 };
      acc.neto += venta - dev;
      acc.dev += dev;
      estado.mensual.set(k, acc);
    }
    const f = j.fecha_max;
    estado.corte = { a: Math.floor(f / 10000), m: Math.floor(f / 100) % 100, d: f % 100, num: f };

    if (!anterior) {
      const a = estado.corte.a;
      const min = Math.floor(j.fecha_min / 10000);
      estado.anios = [a - 1, a].filter((x) => x >= min);
      estado.hastaMes = estado.corte.m;
    }
  }

  function mes(a, m) { return estado.mensual.get(a * 100 + m) || { neto: 0, dev: 0 }; }

  function sumaRango(desde, hasta) {
    let neto = 0, dev = 0;
    const { dias } = estado;
    for (let i = 0; i < dias.length; i++) {
      const d = dias[i];
      if (d >= desde && d <= hasta) { neto += estado.neto[i]; dev += estado.dev[i]; }
    }
    return { neto, dev };
  }

  function porCliente(desde, hasta) {
    const out = new Float64Array(estado.datos.clientes.length);
    const { dias } = estado;
    for (let i = 0; i < dias.length; i++) {
      const d = dias[i];
      if (d >= desde && d <= hasta) out[estado.cli[i]] += estado.neto[i];
    }
    return out;
  }

  // Mueve un rango aaaammdd al año indicado (mismo mes y día).
  const aAnio = (num, anio) => anio * 10000 + (num % 10000);

  // ------------------------------------------------------------------
  // Render
  // ------------------------------------------------------------------
  function render() {
    renderPeriodos();
    renderControles();
    renderKpis();
    renderGrafica();
    renderTablaMeses();
    renderClientes();
    renderFuente();
  }

  function rangoComparable() {
    // Rango acumulado del año A (más reciente seleccionado) y del año B.
    const anios = [...estado.anios].sort((x, y) => y - x);
    const A = anios[0], B = anios[1] ?? null;
    const hm = estado.hastaMes;
    const finTeorico = fechaNum(A, hm, ultimoDia(A, hm));
    const finA = Math.min(finTeorico, estado.corte.num);
    const recortado = finA < finTeorico;
    return {
      A, B, hm, recortado,
      desdeA: fechaNum(A, 1, 1), hastaA: finA,
      desdeB: B ? fechaNum(B, 1, 1) : null,
      hastaB: B ? Math.min(aAnio(finA, B), fechaNum(B, hm, ultimoDia(B, hm))) : null,
    };
  }

  function etiquetaCorte(r) {
    const mesFin = Math.floor(r.hastaA / 100) % 100;
    const base = r.hm === 1 ? "Enero" : `Enero – ${MESES[r.hm - 1]}`;
    return r.recortado ? `${base} (al ${r.hastaA % 100} ${MES_CORTO[mesFin - 1].toLowerCase()})` : base;
  }

  function renderPeriodos() {
    const { a, m, d } = estado.corte;
    const tarjetas = [];
    // Periodo actual: del 1 al día de corte vs. mismo rango del año anterior.
    {
      const act = sumaRango(fechaNum(a, m, 1), fechaNum(a, m, d)).neto;
      const ant = sumaRango(fechaNum(a - 1, m, 1), fechaNum(a - 1, m, d)).neto;
      tarjetas.push({
        etq: "Periodo actual", activo: true,
        titulo: `${MESES[m - 1]} ${a}`, rango: `1 – ${d} ${MES_CORTO[m - 1].toLowerCase()}`,
        act, ant, ref: `Mismo periodo año anterior (1 – ${d} ${MES_CORTO[m - 1].toLowerCase()} ${a - 1})`,
      });
    }
    [["Mes anterior", 1], ["Segundo mes anterior", 2]].forEach(([etq, n]) => {
      const [aa, mm] = mesAnterior(a, m, n);
      const act = mes(aa, mm).neto, ant = mes(aa - 1, mm).neto;
      tarjetas.push({
        etq, titulo: `${MESES[mm - 1]} ${aa}`, rango: "mes completo", act, ant,
        ref: `Mismo periodo año anterior (${MESES[mm - 1].toLowerCase()} ${aa - 1})`,
      });
    });
    $("periodos").innerHTML = tarjetas.map((t) => {
      const v = variacion(t.act, t.ant), dif = t.act - t.ant;
      const pill = v === null ? "" :
        `<span class="pill ${v < 0 ? "neg" : "pos"}">${v < 0 ? "▼" : "▲"} ${pct(v)}</span>`;
      return `<div class="tarjeta periodo ${t.activo ? "activo" : ""}">
        <div class="etq">${t.etq}</div>${pill}
        <div class="rango">${t.titulo} <span>· ${t.rango}</span></div>
        <div class="monto">${dinero(t.act)}</div>
        <div class="pie">${t.ref}: <b>${dinero(t.ant)}</b><br>
          Diferencia: <span class="${clase(dif)}">${dinero(dif)}</span></div>
      </div>`;
    }).join("");
  }

  function colorAnio(anio) {
    const orden = [...estado.anios].sort((x, y) => y - x);
    const i = orden.indexOf(anio);
    return { fuerte: COLORES[i] || "#999", claro: COLORES_CLAROS[i] || "#ccc" };
  }

  function renderControles() {
    const min = Math.floor(estado.datos.fecha_min / 10000);
    const max = estado.corte.a;
    let html = "";
    for (let y = min; y <= max; y++) {
      const sel = estado.anios.includes(y);
      const st = sel ? `style="background:${colorAnio(y).fuerte}"` : "";
      html += `<button class="chip ${sel ? "sel" : ""}" data-anio="${y}" ${st}>${y}</button>`;
    }
    html += `<span class="contador">${estado.anios.length} de 4</span>`;
    $("anios").innerHTML = html;

    const sel = $("hastaMes");
    if (!sel.options.length) {
      sel.innerHTML = MESES.map((n, i) => `<option value="${i + 1}">${n}</option>`).join("");
    }
    sel.value = String(estado.hastaMes);
    $("actualizado").textContent =
      `Actualizado ${estado.hora.toLocaleTimeString("es-CO")} · ${estado.msCarga} ms`;
  }

  function renderKpis() {
    const r = rangoComparable();
    const a = sumaRango(r.desdeA, r.hastaA);
    const b = r.B ? sumaRango(r.desdeB, r.hastaB) : null;
    const etq = etiquetaCorte(r);
    const kpis = [
      { titulo: `Ventas netas ${r.A}`, valor: dinero(a.neto), sub: etq },
      r.B ? { titulo: `Ventas netas ${r.B}`, valor: dinero(b.neto), sub: etq }
        : { titulo: "Año de comparación", valor: "—", sub: "Seleccione un segundo año" },
    ];
    if (b) {
      const dif = a.neto - b.neto, v = variacion(a.neto, b.neto);
      kpis.push({
        titulo: `Variación ${r.A} vs ${r.B}`,
        valor: `<span class="${clase(dif)}">${dif < 0 ? "▼ " : dif > 0 ? "▲ " : ""}${dinero(dif)}</span>`,
        sub: `<span class="${clase(dif)}">${pct(v)}</span>`,
      });
    } else {
      kpis.push({ titulo: "Variación", valor: "—", sub: "" });
    }
    kpis.push({
      titulo: `Devoluciones ${r.A}`, valor: dinero(a.dev),
      sub: b ? `${r.B}: ${dinero(b.dev)}` : etq,
    });
    $("kpis").innerHTML = kpis.map((k) => `<div class="tarjeta kpi">
      <div class="titulo">${k.titulo}</div><div class="valor">${k.valor}</div>
      <div class="sub">${k.sub}</div></div>`).join("");
  }

  function mesParcial() {
    const { a, m, d } = estado.corte;
    return d < ultimoDia(a, m) ? { a, m, d } : null;
  }

  function renderGrafica() {
    const anios = [...estado.anios].sort((x, y) => x - y);
    const hm = estado.hastaMes;
    const parcial = mesParcial();
    const etiquetas = MES_CORTO.slice(0, hm).map((n, i) =>
      parcial && i + 1 === parcial.m && anios.includes(parcial.a) ? n + "*" : n);
    const datasets = anios.map((y) => {
      const c = colorAnio(y);
      const datos = [], colores = [];
      for (let m = 1; m <= hm; m++) {
        const futuro = fechaNum(y, m, 1) > estado.corte.num;
        datos.push(futuro ? null : mes(y, m).neto);
        colores.push(parcial && parcial.a === y && parcial.m === m ? c.claro : c.fuerte);
      }
      return { label: String(y), data: datos, backgroundColor: colores, borderRadius: 2,
        maxBarThickness: 26, categoryPercentage: 0.6, barPercentage: 0.9 };
    });
    $("leyendaGrafica").innerHTML = anios.map((y) =>
      `<span><i style="background:${colorAnio(y).fuerte}"></i>${y}</span>`).join("");

    const ejeY = (v) => "$ " + nf.format(v / 1e6) + " M";
    if (estado.grafica) {
      estado.grafica.data.labels = etiquetas;
      estado.grafica.data.datasets = datasets;
      estado.grafica.update();
      return;
    }
    estado.grafica = new Chart($("grafica"), {
      type: "bar",
      data: { labels: etiquetas, datasets },
      options: {
        responsive: true, maintainAspectRatio: false, animation: { duration: 250 },
        plugins: {
          legend: { display: false },
          tooltip: { callbacks: { label: (ctx) => `${ctx.dataset.label}: ${dinero(ctx.parsed.y)}` } },
        },
        scales: {
          x: { grid: { display: false }, ticks: { color: "#6b7686", font: { size: 11 } } },
          y: { beginAtZero: true, grid: { color: "#eef1f5" }, border: { display: false },
            ticks: { color: "#6b7686", font: { size: 10 }, callback: ejeY, maxTicksLimit: 6 } },
        },
      },
    });
  }

  function datosTablaMeses() {
    const anios = [...estado.anios].sort((x, y) => x - y);
    const A = anios[anios.length - 1], B = anios.length > 1 ? anios[anios.length - 2] : null;
    const filas = [];
    let acumA = 0, acumB = 0;
    const tot = {};
    anios.forEach((y) => (tot[y] = { neto: 0, dev: 0 }));
    for (let m = 1; m <= estado.hastaMes; m++) {
      const f = { m, netos: {}, devs: {} };
      anios.forEach((y) => {
        const v = fechaNum(y, m, 1) > estado.corte.num ? null : mes(y, m);
        f.netos[y] = v ? v.neto : null;
        f.devs[y] = v ? v.dev : null;
        if (v) { tot[y].neto += v.neto; tot[y].dev += v.dev; }
      });
      if (B !== null && f.netos[A] !== null) {
        acumA += f.netos[A]; acumB += f.netos[B] || 0;
        f.dif = f.netos[A] - (f.netos[B] || 0);
        f.var = variacion(f.netos[A], f.netos[B]);
        f.varAcum = variacion(acumA, acumB);
      }
      filas.push(f);
    }
    return { anios, A, B, filas, tot };
  }

  function renderTablaMeses() {
    const { anios, A, B, filas, tot } = datosTablaMeses();
    const parcial = mesParcial();
    let h = `<thead><tr><th class="izq">Mes</th>${anios.map((y) => `<th>Netas ${y}</th>`).join("")}`;
    if (B !== null) h += `<th>Variación ${A} vs ${B}</th><th>Var %</th><th>Var % acum.</th>`;
    h += `${anios.map((y) => `<th>Devol. ${y}</th>`).join("")}</tr></thead><tbody>`;
    filas.forEach((f) => {
      const marca = parcial && f.m === parcial.m && anios.includes(parcial.a) ? " *" : "";
      h += `<tr><td class="izq">${MESES[f.m - 1]}${marca}</td>`;
      h += anios.map((y) => `<td>${f.netos[y] === null ? "—" : dinero(f.netos[y])}</td>`).join("");
      if (B !== null) {
        h += f.dif === undefined ? "<td>—</td><td>—</td><td>—</td>" :
          `<td class="${clase(f.dif)}">${dinero(f.dif)}</td><td class="${clase(f.var)}">${pct(f.var)}</td>` +
          `<td class="${clase(f.varAcum)}">${pct(f.varAcum)}</td>`;
      }
      h += anios.map((y) => `<td>${f.devs[y] === null ? "—" : dinero(f.devs[y])}</td>`).join("");
      h += "</tr>";
    });
    h += `</tbody><tfoot><tr><td class="izq">Total</td>${anios.map((y) => `<td>${dinero(tot[y].neto)}</td>`).join("")}`;
    if (B !== null) {
      const dif = tot[A].neto - tot[B].neto, v = variacion(tot[A].neto, tot[B].neto);
      h += `<td class="${clase(dif)}">${dinero(dif)}</td><td class="${clase(v)}">${pct(v)}</td><td></td>`;
    }
    h += `${anios.map((y) => `<td>${dinero(tot[y].dev)}</td>`).join("")}</tr></tfoot>`;
    $("tablaMeses").innerHTML = h;

    let nota = "Ventas netas = ventas − devoluciones.";
    if (parcial && anios.includes(parcial.a) && estado.hastaMes >= parcial.m) {
      nota = `* ${MESES[parcial.m - 1]} ${parcial.a} con datos hasta el día ${parcial.d}; en la tabla y la gráfica se ` +
        "muestra el mes tal como viene en el archivo (las tarjetas superiores comparan el mismo rango de días). " + nota;
    }
    $("notaMeses").textContent = nota;
  }

  // ------------------------------------------------------------------
  // Top clientes
  // ------------------------------------------------------------------
  function periodosClientes() {
    const { a, m, d } = estado.corte;
    const r = rangoComparable();
    const comp = r.B ?? r.A - 1;
    const finA = r.hastaA;
    const mesFin = Math.floor(finA / 100) % 100;
    const [a1, m1] = mesAnterior(a, m, 1);
    const [a2, m2] = mesAnterior(a, m, 2);
    const rec = r.recortado ? ` (al ${finA % 100})` : "";
    return {
      acum: {
        etq: `Acumulado ${MES_CORTO[0]} – ${MES_CORTO[mesFin - 1]} ${r.A}${rec}`,
        A: r.A, B: comp, desdeA: r.desdeA, hastaA: finA,
        desdeB: fechaNum(comp, 1, 1), hastaB: Math.min(aAnio(finA, comp), fechaNum(comp, r.hm, ultimoDia(comp, r.hm))),
      },
      actual: {
        etq: `Periodo actual · ${MES_CORTO[m - 1]} ${a} (1–${d})`, A: a, B: a - 1,
        desdeA: fechaNum(a, m, 1), hastaA: fechaNum(a, m, d),
        desdeB: fechaNum(a - 1, m, 1), hastaB: fechaNum(a - 1, m, d),
      },
      ant: {
        etq: `Mes anterior · ${MES_CORTO[m1 - 1]} ${a1}`, A: a1, B: a1 - 1,
        desdeA: fechaNum(a1, m1, 1), hastaA: fechaNum(a1, m1, 31),
        desdeB: fechaNum(a1 - 1, m1, 1), hastaB: fechaNum(a1 - 1, m1, 31),
      },
      ant2: {
        etq: `Segundo mes anterior · ${MES_CORTO[m2 - 1]} ${a2}`, A: a2, B: a2 - 1,
        desdeA: fechaNum(a2, m2, 1), hastaA: fechaNum(a2, m2, 31),
        desdeB: fechaNum(a2 - 1, m2, 1), hastaB: fechaNum(a2 - 1, m2, 31),
      },
    };
  }

  function calcularClientes() {
    const periodos = periodosClientes();
    const p = periodos[estado.tab] || periodos.acum;
    const va = porCliente(p.desdeA, p.hastaA);
    const vb = porCliente(p.desdeB, p.hastaB);
    const nombres = estado.datos.clientes;
    const lista = [];
    for (let i = 0; i < nombres.length; i++) {
      if (va[i] !== 0 || vb[i] !== 0) lista.push({ i, nombre: nombres[i], a: va[i], b: vb[i] });
    }
    lista.sort((x, y) => y.a - x.a || y.b - x.b);
    const totA = lista.reduce((s, c) => s + c.a, 0);
    const totB = lista.reduce((s, c) => s + c.b, 0);
    const conA = lista.filter((c) => c.a > 0);
    const conB = lista.filter((c) => c.b > 0);
    const top = conA.slice(0, 20);
    const topB = [...conB].sort((x, y) => y.b - x.b).slice(0, 20);
    return {
      periodos, p, lista, top, totA, totB, conA, conB,
      nuevos: conA.filter((c) => c.b <= 0),
      perdidos: conB.filter((c) => c.a <= 0).sort((x, y) => y.b - x.b),
      concA: totA ? top.reduce((s, c) => s + c.a, 0) / totA : null,
      concB: totB ? topB.reduce((s, c) => s + c.b, 0) / totB : null,
    };
  }

  function listaCorta(clientes, n = 3) {
    if (!clientes.length) return "";
    const nombres = clientes.slice(0, n).map((c) => esc(c.nombre)).join(", ");
    return clientes.length > n ? `${nombres} y ${clientes.length - n} más` : nombres;
  }

  function renderClientes() {
    const r = calcularClientes();
    const { p } = r;
    $("tabsClientes").innerHTML = Object.entries(r.periodos).map(([k, v]) =>
      `<button class="tab ${k === estado.tab ? "sel" : ""}" data-tab="${k}">${esc(v.etq)}</button>`).join("");

    const cA = COLORES[0], cB = COLORES[1];
    $("statsClientes").innerHTML = [
      { t: `Concentración Top 20 · ${p.A}`, v: pct(r.concA, false), s: `${p.B}: ${pct(r.concB, false)}` },
      { t: "Clientes con compras", v: nf.format(r.conA.length), s: `${p.B}: ${nf.format(r.conB.length)}` },
      { t: "Clientes nuevos", v: nf.format(r.nuevos.length), s: `Sin compras en el mismo periodo ${p.B}`,
        l: listaCorta(r.nuevos) },
      { t: `Clientes sin compras en ${p.A}`, v: nf.format(r.perdidos.length), s: `Compraron en ${p.B}`,
        l: listaCorta(r.perdidos) },
    ].map((k) => `<div class="tarjeta kpi mini"><div class="titulo">${k.t}</div>
      <div class="valor">${k.v}</div><div class="sub">${k.s}</div>
      ${k.l ? `<div class="lista">${k.l}</div>` : ""}</div>`).join("");

    $("leyendaClientes").innerHTML =
      `<span><i style="background:${cA}"></i>${p.A}</span><span><i style="background:${cB}"></i>${p.B} (mismo periodo)</span>`;

    const max = Math.max(1, ...r.top.map((c) => Math.max(c.a, c.b)));
    let h = `<thead><tr><th>#</th><th class="izq">Cliente</th><th>Ventas ${p.A}</th><th>Ventas ${p.B}</th>` +
      `<th>Variación</th><th>Var %</th><th>Part. ${p.A}</th><th class="izq">Comparativo</th></tr></thead><tbody>`;
    r.top.forEach((c, n) => {
      const dif = c.a - c.b, v = variacion(c.a, c.b);
      const flecha = v === null ? "" : v < 0 ? "▼ " : "▲ ";
      const nuevo = c.b <= 0 ? '<span class="badge">Nuevo</span>' : "";
      h += `<tr><td class="rank">${n + 1}</td><td class="izq cliente" title="${esc(c.nombre)}">${esc(c.nombre)}${nuevo}</td>` +
        `<td class="fuerte">${dinero(c.a)}</td><td>${dinero(c.b)}</td>` +
        `<td class="${clase(dif)}">${dinero(dif)}</td><td class="${clase(v)}">${flecha}${pct(v)}</td>` +
        `<td>${pct(r.totA ? c.a / r.totA : null, false)}</td>` +
        `<td class="izq barras"><div style="width:${(Math.max(c.a, 0) / max) * 100}%;background:${cA}"></div>` +
        `<div style="width:${(Math.max(c.b, 0) / max) * 100}%;background:${cB}"></div></td></tr>`;
    });
    const topA = r.top.reduce((s, c) => s + c.a, 0), topB = r.top.reduce((s, c) => s + c.b, 0);
    const otrosN = r.lista.length - r.top.length;
    const oA = r.totA - topA, oB = r.totB - topB, oD = oA - oB, oV = variacion(oA, oB);
    h += `</tbody><tbody>`;
    if (otrosN > 0) {
      h += `<tr class="otros"><td></td><td class="izq">Otros ${nf.format(otrosN)} clientes</td><td>${dinero(oA)}</td>` +
        `<td>${dinero(oB)}</td><td class="${clase(oD)}">${dinero(oD)}</td><td class="${clase(oV)}">${pct(oV)}</td>` +
        `<td>${pct(r.totA ? oA / r.totA : null, false)}</td><td></td></tr>`;
    }
    const tD = r.totA - r.totB, tV = variacion(r.totA, r.totB);
    h += `<tr class="total"><td></td><td class="izq">Total ventas netas</td><td>${dinero(r.totA)}</td>` +
      `<td>${dinero(r.totB)}</td><td class="${clase(tD)}">${dinero(tD)}</td><td class="${clase(tV)}">${pct(tV)}</td>` +
      `<td></td><td></td></tr></tbody>`;
    $("tablaClientes").innerHTML = h;
    $("notaClientes").textContent =
      `Ranking por ventas netas (ventas − devoluciones) de ${p.A} en el periodo elegido, comparadas con el mismo ` +
      `periodo de ${p.B}. El mes en curso se compara con el mismo rango de días del año anterior.`;
    estado.ultimoClientes = r;
  }

  function renderFuente() {
    const f = estado.datos.fuente;
    const hojas = f.hojas.map((h) => `${h.hoja} (modo ${h.modo}, encabezado fila ${h.fila_encabezado}, ` +
      `columnas: ${Object.entries(h.columnas).map(([k, v]) => `${k}=“${v}”`).join(", ")})`).join(" · ");
    const avisos = f.avisos.length ? ` · Avisos: ${f.avisos.join(" | ")}` : "";
    $("fuente").textContent = `Fuente: ${f.ruta} · ${nf.format(f.filas)} movimientos · ` +
      `datos del ${fmtFecha(estado.datos.fecha_min)} al ${fmtFecha(estado.datos.fecha_max)} · ` +
      `lectura ${estado.datos.ms_lectura} ms · ${hojas}${avisos}`;
  }

  const fmtFecha = (n) => `${String(n % 100).padStart(2, "0")}/${String(Math.floor(n / 100) % 100).padStart(2, "0")}/${Math.floor(n / 10000)}`;

  // ------------------------------------------------------------------
  // CSV
  // ------------------------------------------------------------------
  const num = (n) => (n === null || n === undefined ? "" : String(Math.round(n * 100) / 100).replace(".", ","));
  const pctCsv = (p) => (p === null || p === undefined || !isFinite(p) ? "" : (p * 100).toFixed(1).replace(".", ","));
  const celda = (s) => (/[;"\n]/.test(String(s)) ? `"${String(s).replace(/"/g, '""')}"` : String(s));

  async function copiar(filas, descripcion) {
    const texto = filas.map((f) => f.map(celda).join(";")).join("\r\n");
    try {
      await navigator.clipboard.writeText(texto);
    } catch {
      const ta = document.createElement("textarea");
      ta.value = texto; document.body.appendChild(ta); ta.select();
      document.execCommand("copy"); ta.remove();
    }
    toast(`${descripcion} copiado (separado por ;) — péguelo en Excel`);
  }

  function csvMeses() {
    const { anios, A, B, filas, tot } = datosTablaMeses();
    const cab = ["Mes", ...anios.map((y) => `Netas ${y}`)];
    if (B !== null) cab.push(`Variación ${A} vs ${B}`, "Var %", "Var % acum.");
    cab.push(...anios.map((y) => `Devol. ${y}`));
    const out = [cab];
    filas.forEach((f) => {
      const fila = [MESES[f.m - 1], ...anios.map((y) => num(f.netos[y]))];
      if (B !== null) fila.push(num(f.dif), pctCsv(f.var), pctCsv(f.varAcum));
      fila.push(...anios.map((y) => num(f.devs[y])));
      out.push(fila);
    });
    const t = ["Total", ...anios.map((y) => num(tot[y].neto))];
    if (B !== null) t.push(num(tot[A].neto - tot[B].neto), pctCsv(variacion(tot[A].neto, tot[B].neto)), "");
    t.push(...anios.map((y) => num(tot[y].dev)));
    out.push(t);
    copiar(out, "Detalle mensual");
  }

  function csvClientes() {
    const r = estado.ultimoClientes;
    if (!r) return;
    const out = [["#", "Cliente", `Ventas ${r.p.A}`, `Ventas ${r.p.B}`, "Variación", "Var %", `Part. ${r.p.A} %`]];
    r.lista.forEach((c, n) => out.push([n + 1, c.nombre, num(c.a), num(c.b), num(c.a - c.b),
      pctCsv(variacion(c.a, c.b)), pctCsv(r.totA ? c.a / r.totA : null)]));
    out.push(["", "Total ventas netas", num(r.totA), num(r.totB), num(r.totA - r.totB), pctCsv(variacion(r.totA, r.totB)), ""]);
    copiar(out, `Clientes (${r.p.etq})`);
  }

  // ------------------------------------------------------------------
  // Utilidades e interacción
  // ------------------------------------------------------------------
  let toastTimer;
  function toast(msg) {
    const t = $("toast");
    t.textContent = msg; t.hidden = false;
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => (t.hidden = true), 3200);
  }

  $("anios").addEventListener("click", (e) => {
    const b = e.target.closest("[data-anio]");
    if (!b) return;
    const y = Number(b.dataset.anio);
    if (estado.anios.includes(y)) {
      if (estado.anios.length === 1) return toast("Debe quedar al menos un año seleccionado");
      estado.anios = estado.anios.filter((x) => x !== y);
    } else {
      if (estado.anios.length >= 4) return toast("Máximo 4 años: quite uno antes de agregar otro");
      estado.anios = [...estado.anios, y];
    }
    render();
  });
  $("hastaMes").addEventListener("change", (e) => { estado.hastaMes = Number(e.target.value); render(); });
  $("tabsClientes").addEventListener("click", (e) => {
    const b = e.target.closest("[data-tab]");
    if (!b) return;
    estado.tab = b.dataset.tab;
    renderClientes();
  });
  $("btnGenerar").addEventListener("click", () => cargarDatos());
  $("btnCsvMeses").addEventListener("click", () => estado.datos && csvMeses());
  $("btnCsvClientes").addEventListener("click", () => estado.datos && csvClientes());

  // Conexión automática: si el Excel cambia, se recargan los datos.
  async function sondear() {
    try {
      const r = await fetch("api/estado", { cache: "no-store" });
      const j = await r.json();
      if (estado.firma && j.firma !== estado.firma) await cargarDatos(true);
      else if (!estado.datos) await cargarDatos();
    } catch { /* servidor detenido: se reintenta en el próximo ciclo */ }
  }

  cargarDatos();
  setInterval(sondear, SONDEO_MS);
})();
