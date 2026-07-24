// ============================================================
// app.js — Dashboard MPL Egresos · Municipalidad Provincial de Lambayeque
// Versión 1.0 — Julio 2026
// Bloques: B1 KPIs · B2 Rubro · B3 Categoría · B4 Proyectos · B5 Ranking · B6 Histórico
// Fuente: MEF Consulta Amigable de Gastos — Solo Proyectos
// ============================================================

"use strict";

// ── Utilidades ───────────────────────────────────────────────
const $ = id => document.getElementById(id);

function fmtS(n) {
  if (n === null || isNaN(n)) return "—";
  if (Math.abs(n) >= 1e6) return "S/ " + (n / 1e6).toFixed(2) + " M";
  if (Math.abs(n) >= 1e3) return "S/ " + (n / 1e3).toFixed(1) + " K";
  return "S/ " + Math.round(n).toLocaleString("es-PE");
}

function fmtNum(n) {
  if (n === null || isNaN(n)) return "—";
  return "S/ " + Math.round(n).toLocaleString("es-PE");
}

function fmtCompacto(n) {
  if (!n && n !== 0) return "—";
  if (Math.abs(n) >= 1e6) return "S/ " + (n / 1e6).toFixed(1) + " M";
  if (Math.abs(n) >= 1e3) return "S/ " + (n / 1e3).toFixed(0) + " K";
  return "S/ " + Math.round(n).toLocaleString("es-PE");
}

function parseNum(s) {
  if (!s) return 0;
  return parseFloat(String(s).replace(/,/g, "")) || 0;
}

function fechaHoy() {
  const d = new Date();
  const meses = ["enero","febrero","marzo","abril","mayo","junio",
                  "julio","agosto","setiembre","octubre","noviembre","diciembre"];
  const dias  = ["domingo","lunes","martes","miércoles","jueves","viernes","sábado"];
  return dias[d.getDay()] + ", " + d.getDate() + " de " + meses[d.getMonth()] + " de " + d.getFullYear();
}

function semaforo(pct) {
  if (pct === null) return "#888888";
  if (pct >= 70)   return "var(--verde)";
  if (pct >= 40)   return "var(--amarillo-s)";
  return "var(--rojo-s)";
}

function barraHTML(pct) {
  if (pct === null) return '<span style="color:#888;font-size:.8em">N/A</span>';
  const color = semaforo(pct);
  const ancho = Math.min(pct, 100).toFixed(1);
  return `<div style="display:flex;align-items:center;gap:6px">
    <div style="flex:1;background:#e5e7eb;border-radius:4px;height:8px;min-width:60px">
      <div style="width:${ancho}%;background:${color};height:8px;border-radius:4px"></div>
    </div>
    <span style="font-size:.82em;font-weight:700;color:${color};min-width:40px">${pct.toFixed(1)}%</span>
  </div>`;
}

// ── Estado global ────────────────────────────────────────────
let datos = {};
let cargados = new Set();
let b6ChartInstance = null;
let B6_HIST = {};

const ARCHIVOS_ESPERADOS = ["rubro.xls", "categoria.xls", "proyecto.xls", "ranking.xls"];

// ── Parseo de archivos XLS (HTML disfrazado) ──────────────────
function extraerFilas(buffer) {
  const text = new TextDecoder("utf-8").decode(buffer);
  const trRe = /<tr[^>]*>([\s\S]*?)<\/tr>/gi;
  const tdRe = /<t[dh][^>]*>([\s\S]*?)<\/t[dh]>/gi;
  const rows = [];
  let trM;
  while ((trM = trRe.exec(text)) !== null) {
    const cells = [];
    let tdM;
    const inner = trM[1];
    tdRe.lastIndex = 0;
    while ((tdM = tdRe.exec(inner)) !== null) {
      const val = tdM[1]
        .replace(/<[^>]+>/g, "")
        .replace(/&nbsp;/g, "")
        .replace(/\r\n/g, "")
        .trim();
      cells.push(val);
    }
    if (cells.some(c => c)) rows.push(cells);
  }
  return rows;
}

// ── Estructura de columnas MEF Gastos ─────────────────────────
// Col 0: Descripción
// Col 1: PIA  Col 2: PIM  Col 3: Certificación
// Col 4: Compromiso Anual
// Col 5: Atención Compromiso Mensual (sub-col Ejecución)
// Col 6: Devengado (sub-col Ejecución)
// Col 7: Girado (sub-col Ejecución)
// Col 8: % Avance

function detectarTipo(rows) {
  for (const r of rows) {
    const d = (r[0] || "").trim();
    if (/^Rubro$/i.test(d))             return "rubro";
    if (/^Categor/i.test(d))            return "categoria";
    if (/^Proyecto$/i.test(d))          return "proyecto";
    if (/^Municipalidad$/i.test(d))     return "ranking";
  }
  return null;
}

function esCabecera(r) {
  const d = (r[0] || "").trim();
  return /^(Rubro|Categor|Proyecto|Municipalidad|Mes|Atenci|Total|Nivel|Gob\.|Departamento|Municipalidad\s+\d|Consulta)/i.test(d);
}

function esFilaDato(r) {
  if (!r || r.length < 7) return false;
  const d = (r[0] || "").trim();
  if (!d) return false;
  if (esCabecera(r)) return false;
  // Debe tener al menos un número en cols 1-7
  return [1,2,3,4,5,6,7].some(i => /[\d,]+/.test(r[i] || ""));
}

function parsearArchivo(buffer, nombre) {
  const rows = extraerFilas(buffer);
  const tipo = detectarTipo(rows);
  if (!tipo) {
    console.warn("[MPL-EG] No se pudo detectar tipo:", nombre);
    return null;
  }

  // Buscar fila totales de MPL (antes de los datos de detalle)
  let totalPIA = 0, totalPIM = 0, totalDev = 0, totalCert = 0,
      totalComp = 0, totalGirado = 0, avancePct = null;
  for (const r of rows) {
    if ((r[0] || "").includes("MUNICIPALIDAD PROVINCIAL DE LAMBAYEQUE") ||
        (r[0] || "").includes("140301-301238")) {
      totalPIA    = parseNum(r[1]);
      totalPIM    = parseNum(r[2]);
      totalCert   = parseNum(r[3]);
      totalComp   = parseNum(r[4]);
      totalDev    = parseNum(r[6]);
      totalGirado = parseNum(r[7]);
      avancePct   = totalPIM > 0 ? totalDev / totalPIM * 100 : null;
      break;
    }
  }
  // Para ranking, MPL no aparece en el header — calcular del total del dept.
  if (tipo === "ranking") {
    // Buscar fila Lambayeque departamento
    for (const r of rows) {
      if ((r[0] || "").includes("LAMBAYEQUE") && !r[0].includes("Municipalidad")) {
        totalPIM = parseNum(r[2]);
        totalDev = parseNum(r[6]);
        avancePct = totalPIM > 0 ? totalDev / totalPIM * 100 : null;
        break;
      }
    }
  }

  // Filas de detalle
  const registros = [];
  for (const r of rows) {
    if (!esFilaDato(r)) continue;
    registros.push({
      desc:    r[0] || "",
      pia:     parseNum(r[1]),
      pim:     parseNum(r[2]),
      cert:    parseNum(r[3]),
      comp:    parseNum(r[4]),
      acm:     parseNum(r[5]),
      dev:     parseNum(r[6]),
      girado:  parseNum(r[7]),
      pct:     parseNum(r[8]) || null
    });
  }

  // Recalcular totales de MPL desde rubro.xls para B1 si totalPIM=0
  if (tipo === "rubro" && totalPIM === 0 && registros.length) {
    totalPIA    = registros.reduce((s, r) => s + r.pia, 0);
    totalPIM    = registros.reduce((s, r) => s + r.pim, 0);
    totalCert   = registros.reduce((s, r) => s + r.cert, 0);
    totalComp   = registros.reduce((s, r) => s + r.comp, 0);
    totalDev    = registros.reduce((s, r) => s + r.dev, 0);
    totalGirado = registros.reduce((s, r) => s + r.girado, 0);
    avancePct   = totalPIM > 0 ? totalDev / totalPIM * 100 : null;
  }

  console.log(`[MPL-EG] Parseo OK: ${nombre} → tipo=${tipo} | registros=${registros.length} | PIM=${totalPIM} | Dev=${totalDev}`);

  return {
    tipo, nombre, registros,
    totalPIA, totalPIM, totalCert, totalComp, totalDev, totalGirado, avancePct
  };
}

// ── Carga automática de archivos ──────────────────────────────
async function autoCargar() {
  for (const nombre of ARCHIVOS_ESPERADOS) {
    try {
      const r = await fetch("xls/" + nombre + "?" + Date.now());
      if (!r.ok) continue;
      const buf = await r.arrayBuffer();
      const res = parsearArchivo(buf, nombre);
      if (res) {
        datos[res.tipo] = res;
        cargados.add(nombre);
        actualizarFileList();
        render();
      }
    } catch (e) { /* archivo no disponible */ }
  }
}

function procesarArchivos(files) {
  Array.from(files).forEach(file => {
    const reader = new FileReader();
    reader.onload = e => {
      const res = parsearArchivo(e.target.result, file.name);
      if (res) {
        datos[res.tipo] = res;
        cargados.add(file.name);
        actualizarFileList();
        render();
      } else {
        console.warn("[MPL-EG] No se pudo parsear:", file.name);
      }
    };
    reader.readAsArrayBuffer(file);
  });
}

function actualizarFileList() {
  const fl = $("fileList");
  if (!fl) return;
  fl.innerHTML = ARCHIVOS_ESPERADOS.map(f => {
    const ok = cargados.has(f);
    return `<div class="file-item">
      <span class="${ok ? "fi-ok" : "fi-wait"}">${ok ? "✓" : "○"}</span>
      <span class="fi-name">${f}</span>
    </div>`;
  }).join("");
}

// ── Render principal ──────────────────────────────────────────
function render() {
  const hoy = fechaHoy();
  ["b1fecha","b2fecha","b3fecha","b4fecha","b5fecha"].forEach(id => {
    const el = $(id); if (el) el.textContent = hoy;
  });
  renderB1();
  renderB2();
  renderB3();
  renderB4();
  renderB5();
  renderB6();
}

// ── B1 — KPIs + Donuts ──────────────────────────────────────
let donutInstances = {};

function renderDonut(canvasId, pct, colorFilled, colorEmpty) {
  const canvas = $(canvasId);
  if (!canvas) return;
  if (donutInstances[canvasId]) {
    donutInstances[canvasId].destroy();
    delete donutInstances[canvasId];
  }
  const safePct = Math.min(Math.max(pct || 0, 0), 100);
  donutInstances[canvasId] = new Chart(canvas, {
    type: "doughnut",
    data: {
      datasets: [{
        data: [safePct, 100 - safePct],
        backgroundColor: [colorFilled, colorEmpty],
        borderWidth: 0,
        hoverOffset: 0
      }]
    },
    options: {
      cutout: "72%",
      responsive: false,
      animation: { duration: 600 },
      plugins: { legend: { display: false }, tooltip: { enabled: false } }
    }
  });
}

function renderB1() {
  const d = datos.rubro;
  if (!d) return;

  const pim    = d.totalPIM;
  const cert   = d.totalCert;
  const comp   = d.totalComp;
  const dev    = d.totalDev;
  const porDev = pim - dev;

  const pctCert = pim > 0 ? cert / pim * 100 : 0;
  const pctDev  = pim > 0 ? dev  / pim * 100 : 0;
  const pctComp = pim > 0 ? comp / pim * 100 : 0;

  // ── Fila KPI cards ──
  $("kpi-pim").textContent     = fmtNum(pim);
  $("kpi-pim-sub").textContent = d.registros.length + " registros cargados";

  $("kpi-cert").textContent     = fmtNum(cert);
  $("kpi-cert-sub").textContent = pctCert.toFixed(1) + "%";

  $("kpi-comp").textContent     = fmtNum(comp);
  $("kpi-comp-sub").textContent = pctComp.toFixed(1) + "%";

  $("kpi-dev").textContent     = fmtNum(dev);
  $("kpi-dev-sub").textContent = pctDev.toFixed(1) + "%";

  $("kpi-pordev").textContent  = fmtNum(porDev);

  // ── Donuts ──
  // Donut Certificado (verde)
  renderDonut("donutCert", pctCert, "#1e7e34", "#a8d5b0");
  $("donutCertPct").textContent = pctCert.toFixed(1) + "%";
  $("legCert").textContent    = fmtNum(cert);
  $("legPorCert").textContent = fmtNum(pim - cert);

  // Donut Devengado (rojo)
  renderDonut("donutDev", pctDev, "#7a1219", "#e8c5c7");
  $("donutDevPct").textContent = pctDev.toFixed(1) + "%";
  $("legDev").textContent    = fmtNum(dev);
  $("legPorDev").textContent = fmtNum(porDev);

  // Donut Compromiso (dorado)
  renderDonut("donutComp", pctComp, "#f5c518", "#faeab0");
  $("donutCompPct").textContent = pctComp.toFixed(1) + "%";
  $("legComp").textContent    = fmtNum(comp);
  $("legPorComp").textContent = fmtNum(pim - comp);
}

// ── B2 — Rubro ───────────────────────────────────────────────
function renderB2() {
  const d = datos.rubro;
  const tbody = $("b2tbody");
  const tfoot = $("b2tfoot");
  if (!tbody) return;
  if (!d) { tbody.innerHTML = `<tr><td colspan="8" class="vacio">Carga rubro.xls para ver los datos.</td></tr>`; return; }

  tbody.innerHTML = d.registros.map(r => {
    const pct = r.pim > 0 ? r.dev / r.pim * 100 : null;
    const porDev = r.pim - r.dev;
    return `<tr>
      <td>${r.desc}</td>
      <td class="num">${fmtNum(r.pia)}</td>
      <td class="num">${fmtNum(r.pim)}</td>
      <td class="num">${fmtNum(r.cert)}</td>
      <td class="num">${fmtNum(r.dev)}</td>
      <td class="num" style="color:${porDev < 0 ? "var(--rojo-s)" : "inherit"}">${fmtNum(porDev)}</td>
      <td>${barraHTML(pct)}</td>
    </tr>`;
  }).join("");

  const tot = d;
  const pctTot = tot.totalPIM > 0 ? tot.totalDev / tot.totalPIM * 100 : null;
  tfoot.innerHTML = `<tr>
    <td style="font-weight:800">TOTAL MPL</td>
    <td class="num" style="color:var(--rojo-osc);font-weight:800">${fmtNum(tot.totalPIA)}</td>
    <td class="num" style="color:var(--rojo-osc);font-weight:800">${fmtNum(tot.totalPIM)}</td>
    <td class="num" style="color:var(--rojo-osc);font-weight:800">${fmtNum(tot.totalCert)}</td>
    <td class="num" style="color:var(--rojo-osc);font-weight:800">${fmtNum(tot.totalDev)}</td>
    <td class="num" style="color:var(--rojo-osc);font-weight:800">${fmtNum(tot.totalPIM - tot.totalDev)}</td>
    <td>${barraHTML(pctTot)}</td>
  </tr>`;
}

// ── B3 — Categoría Presupuestal ──────────────────────────────
function renderB3() {
  const d = datos.categoria;
  const tbody = $("b3tbody");
  const tfoot = $("b3tfoot");
  if (!tbody) return;
  if (!d) { tbody.innerHTML = `<tr><td colspan="7" class="vacio">Carga categoria.xls para ver los datos.</td></tr>`; return; }

  tbody.innerHTML = d.registros.map(r => {
    const pct = r.pim > 0 ? r.dev / r.pim * 100 : null;
    return `<tr>
      <td>${r.desc}</td>
      <td class="num">${fmtNum(r.pim)}</td>
      <td class="num">${fmtNum(r.cert)}</td>
      <td class="num">${fmtNum(r.comp)}</td>
      <td class="num">${fmtNum(r.dev)}</td>
      <td class="num">${fmtNum(r.girado)}</td>
      <td>${barraHTML(pct)}</td>
    </tr>`;
  }).join("");

  const tot = d;
  const pctTot = tot.totalPIM > 0 ? tot.totalDev / tot.totalPIM * 100 : null;
  tfoot.innerHTML = `<tr>
    <td style="font-weight:800">TOTAL</td>
    <td class="num" style="color:var(--rojo-osc);font-weight:800">${fmtNum(tot.totalPIM)}</td>
    <td class="num" style="color:var(--rojo-osc);font-weight:800">${fmtNum(tot.totalCert)}</td>
    <td class="num" style="color:var(--rojo-osc);font-weight:800">${fmtNum(tot.totalComp)}</td>
    <td class="num" style="color:var(--rojo-osc);font-weight:800">${fmtNum(tot.totalDev)}</td>
    <td class="num" style="color:var(--rojo-osc);font-weight:800">${fmtNum(tot.totalGirado)}</td>
    <td>${barraHTML(pctTot)}</td>
  </tr>`;
}

// ── B4 — Proyectos de Inversión ──────────────────────────────
function renderB4() {
  const d = datos.proyecto;
  const tbody = $("b4tbody");
  const tfoot = $("b4tfoot");
  if (!tbody) return;
  if (!d) { tbody.innerHTML = `<tr><td colspan="7" class="vacio">Carga proyecto.xls para ver los datos.</td></tr>`; return; }

  // Ordenar por devengado descendente
  const sorted = [...d.registros].sort((a, b) => b.dev - a.dev);

  tbody.innerHTML = sorted.map((r, i) => {
    const pct = r.pim > 0 ? r.dev / r.pim * 100 : null;
    // Extraer código y nombre del proyecto
    const m = r.desc.match(/^(\d+):\s*(.+)$/);
    const cod  = m ? m[1] : "";
    const nom  = m ? m[2] : r.desc;
    return `<tr>
      <td style="text-align:center;color:var(--texto-mut);font-size:.82em">${i + 1}</td>
      <td style="color:var(--texto-mut);font-size:.8em">${cod}</td>
      <td style="font-size:.85em">${nom}</td>
      <td class="num">${fmtNum(r.pim)}</td>
      <td class="num">${fmtNum(r.cert)}</td>
      <td class="num">${fmtNum(r.dev)}</td>
      <td>${barraHTML(pct)}</td>
    </tr>`;
  }).join("");

  const tot = d;
  const pctTot = tot.totalPIM > 0 ? tot.totalDev / tot.totalPIM * 100 : null;
  tfoot.innerHTML = `<tr>
    <td colspan="3" style="font-weight:800">TOTAL</td>
    <td class="num" style="color:var(--rojo-osc);font-weight:800">${fmtNum(tot.totalPIM)}</td>
    <td class="num" style="color:var(--rojo-osc);font-weight:800">${fmtNum(tot.totalCert)}</td>
    <td class="num" style="color:var(--rojo-osc);font-weight:800">${fmtNum(tot.totalDev)}</td>
    <td>${barraHTML(pctTot)}</td>
  </tr>`;
}

// ── B5 — Ranking de Municipalidades ──────────────────────────
function renderB5() {
  const d = datos.ranking;
  const tbody = $("b5tbody");
  const hl    = $("b5highlight");
  const nota  = $("b5nota");
  if (!tbody) return;
  if (!d) { tbody.innerHTML = `<tr><td colspan="6" class="vacio">Carga ranking.xls para ver los datos.</td></tr>`; return; }

  const muns = d.registros.map(r => {
    const m = r.desc.match(/^(\d+-\d+):\s*MUNICIPALIDAD\s+(.+)$/i);
    const nombre = m ? "MPL " + m[2] : r.desc;
    const esMPL  = r.desc.includes("140301-301238") || r.desc.toLowerCase().includes("provincial de lambayeque");
    return { ...r, nombre, esMPL, pct: r.pim > 0 ? r.dev / r.pim * 100 : null };
  });

  // Ordenar por % devengado desc
  const porPct = [...muns].sort((a, b) => (b.pct ?? -1) - (a.pct ?? -1));
  const mpl = porPct.find(m => m.esMPL);
  const posPct   = mpl ? porPct.indexOf(mpl) + 1 : null;
  const porMonto = [...muns].sort((a, b) => b.dev - a.dev);
  const posMonto = mpl ? porMonto.indexOf(mpl) + 1 : null;

  if (mpl && hl) {
    const pctStr = mpl.pct !== null ? mpl.pct.toFixed(1) + "%" : "N/A";
    hl.style.display = "";
    hl.innerHTML = `&#128269; La <strong>Municipalidad Provincial de Lambayeque</strong> ocupa el puesto
      <strong>${posPct}° de ${muns.length}</strong> en % de avance de gasto (${pctStr}).
      Por <strong>monto devengado absoluto</strong>, se ubica en el puesto
      <strong>${posMonto}° de ${muns.length}</strong> con ${fmtS(mpl.dev)}.`;
  }

  tbody.innerHTML = porPct.map((m, i) => {
    const pos = i + 1;
    const trC = m.esMPL ? 'class="mpl-row"' : "";
    const posStyle = pos <= 3
      ? 'style="color:var(--dorado-osc);font-weight:800;text-align:center"'
      : 'style="font-weight:700;text-align:center"';
    return `<tr ${trC}>
      <td ${posStyle}>${pos}°</td>
      <td style="font-weight:${m.esMPL ? "700" : "500"};font-size:.85em">${m.nombre}</td>
      <td class="num">${fmtNum(m.pim)}</td>
      <td class="num">${fmtNum(m.dev)}</td>
      <td>${barraHTML(m.pct)}</td>
    </tr>`;
  }).join("");

  if (nota) nota.style.display = "";
}

// ── B6 — Histórico Ene–Jul ────────────────────────────────────
function renderB6() {
  const años   = Object.keys(B6_HIST).map(Number).sort();
  const dev2026 = (() => {
    if (!datos.rubro) return null;
    // El archivo diario de rubro.xls tiene el devengado acumulado actual de MPL
    return datos.rubro.totalDev;
  })();

  if (dev2026 !== null) años.push(2026);

  const IDX_2026 = años.length - 1;

  const valores = años.map(a => a === 2026 ? dev2026 : (B6_HIST[a] ?? 0));

  // KPI cards históricas
  const kpiWrap = $("b6kpis");
  if (kpiWrap) {
    kpiWrap.innerHTML = años.map((a, i) => {
      const v = valores[i];
      const es2026 = (a === 2026);
      let variacion = "";
      if (i > 0 && valores[i - 1] > 0) {
        const delta = (v - valores[i - 1]) / valores[i - 1] * 100;
        const col   = delta >= 0 ? "#2a7d46" : "#c0392b";
        const sym   = delta >= 0 ? "▲" : "▼";
        variacion = `<div style="font-size:.75em;color:${col};margin-top:2px">${sym} ${Math.abs(delta).toFixed(1)}% vs ${a - 1}</div>`;
      }
      return `<div class="kpi-hist ${es2026 ? "kpi-hist-2026" : ""}">
        <div class="kpi-hist-label">${a}${es2026 ? " · Acum." : " · Ene–Jul"}</div>
        <div class="kpi-hist-val">${fmtCompacto(v)}</div>
        ${variacion}
      </div>`;
    }).join("");
  }

  // Gráfico
  const canvas = $("b6chart");
  if (!canvas) return;
  if (b6ChartInstance) { b6ChartInstance.destroy(); b6ChartInstance = null; }

  const bgColors    = años.map((a, i) => i === IDX_2026 ? "#FFC526" : "#9a1820");
  const borderColors = años.map((a, i) => i === IDX_2026 ? "#d9a000" : "#7a1219");
  const labelColors  = años.map((a, i) => i === IDX_2026 ? "#92400e" : "#7a1219");
  const tickColors   = años.map((a, i) => i === IDX_2026 ? "#92400e" : "#374151");

  const b6Labels = {
    id: "b6Labels",
    afterDatasetsDraw(chart) {
      const { ctx, data, scales: { y } } = chart;
      ctx.save();
      data.datasets[0].data.forEach((val, i) => {
        if (val === null || val === 0) return;
        const meta = chart.getDatasetMeta(0);
        const bar  = meta.data[i];
        ctx.fillStyle = labelColors[i];
        ctx.font = "bold 11px 'Barlow Condensed', sans-serif";
        ctx.textAlign = "center";
        ctx.fillText(fmtCompacto(val), bar.x, bar.y - 6);
      });
      ctx.restore();
    }
  };

  b6ChartInstance = new Chart(canvas, {
    type: "bar",
    data: {
      labels: años.map(String),
      datasets: [{
        data: valores,
        backgroundColor: bgColors,
        borderColor: borderColors,
        borderWidth: 1.5,
        borderRadius: 4
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: {
          callbacks: {
            label: ctx => "Devengado: " + fmtNum(ctx.parsed.y)
          }
        }
      },
      scales: {
        x: {
          grid: { display: false },
          ticks: {
            color: ctx => tickColors[ctx.index] || "#374151",
            font: { family: "'Barlow Condensed', sans-serif", size: 13, weight: "600" }
          }
        },
        y: {
          grid: { color: "#f3f4f6" },
          ticks: {
            color: "#6b7280",
            font: { family: "Barlow, sans-serif", size: 11 },
            callback: v => fmtCompacto(v)
          }
        }
      }
    },
    plugins: [b6Labels]
  });
}

// ── Event Listeners ───────────────────────────────────────────
$("dropzone").addEventListener("click", () => $("file").click());
$("dropzone").addEventListener("keydown", e => {
  if (e.key === "Enter" || e.key === " ") { e.preventDefault(); $("file").click(); }
});
$("file").addEventListener("change", e => {
  procesarArchivos(e.target.files); e.target.value = "";
});
$("clear").addEventListener("click", () => {
  datos = {}; cargados = new Set(); actualizarFileList(); render();
});

const dz = $("dropzone");
["dragenter","dragover"].forEach(ev =>
  dz.addEventListener(ev, e => { e.preventDefault(); dz.classList.add("drag"); }));
["dragleave","drop"].forEach(ev =>
  dz.addEventListener(ev, e => { e.preventDefault(); dz.classList.remove("drag"); }));
dz.addEventListener("drop", e => {
  if (e.dataTransfer.files.length) procesarArchivos(e.dataTransfer.files);
});

// ── PDF ───────────────────────────────────────────────────────
let pdfLibsCargadas = false;

async function cargarLibsPDF() {
  if (pdfLibsCargadas) return;
  await new Promise((res, rej) => {
    const s = document.createElement("script");
    s.src = "https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js";
    s.onload = res; s.onerror = rej;
    document.head.appendChild(s);
  });
  await new Promise((res, rej) => {
    const s = document.createElement("script");
    s.src = "https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js";
    s.onload = res; s.onerror = rej;
    document.head.appendChild(s);
  });
  pdfLibsCargadas = true;
}

async function exportarPDF() {
  const btn = document.querySelector(".btn-pdf");
  btn.textContent = "⏳ Generando...";
  btn.disabled = true;
  try {
    await cargarLibsPDF();
    const { jsPDF } = window.jspdf;
    const pdf = new jsPDF({ orientation: "landscape", unit: "mm", format: "a4" });
    const bloques = document.querySelectorAll(".slide");
    let primera = true;
    for (const bloque of bloques) {
      const canvas = await html2canvas(bloque, { scale: 1.5, useCORS: true });
      const imgData = canvas.toDataURL("image/jpeg", 0.9);
      const pw = pdf.internal.pageSize.getWidth();
      const ph = pdf.internal.pageSize.getHeight();
      const ratio = Math.min(pw / canvas.width, ph / canvas.height) * 96;
      const w = canvas.width * ratio;
      const h = canvas.height * ratio;
      if (!primera) pdf.addPage();
      pdf.addImage(imgData, "JPEG", (pw - w) / 2, (ph - h) / 2, w, h);
      primera = false;
    }
    const hoy = new Date();
    pdf.save(`MPL_Egresos_${hoy.getFullYear()}${String(hoy.getMonth()+1).padStart(2,"0")}${String(hoy.getDate()).padStart(2,"0")}.pdf`);
  } catch (e) {
    console.error("PDF error:", e);
    alert("Error al generar PDF. Intente de nuevo.");
  } finally {
    btn.textContent = "⬇ Descargar PDF";
    btn.disabled = false;
  }
}

// ── Inicio ────────────────────────────────────────────────────
fetch("data/historico_egresos.json?" + Date.now())
  .then(r => r.json())
  .then(data => {
    B6_HIST = data;
    renderB6();
  })
  .catch(() => console.warn("[MPL-EG] No se pudo cargar historico_egresos.json"));

autoCargar();
