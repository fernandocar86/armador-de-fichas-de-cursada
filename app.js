const $ = id => document.getElementById(id);
const files = {gen: null, con: null};
let estudiantes = [], datosCursada = '';

function msg(t, cls){ const m = $('msg'); m.textContent = t; m.className = cls || ''; }
function habilitar(){ $('go').disabled = !(files.gen && files.con); }
[['fgen','gen','ngen'],['fcon','con','ncon']].forEach(([id,k,n]) => {
  $(id).addEventListener('change', e => {
    files[k] = e.target.files[0] || null;
    $(n).textContent = files[k] ? files[k].name : 'Ningún archivo';
    $('result').hidden = true; habilitar();
  });
});

async function leerHoja(file){
  const wb = XLSX.read(await file.arrayBuffer(), {type: 'array'});
  const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], {defval: null});
  return rows.map(r => Object.fromEntries(Object.entries(r).map(([k,v]) => [k.trim(), typeof v === 'string' ? v.trim() : v])));
}

// "Parcial 1" -> "P1", "Tema 2" -> "T2" (igual que el script original: se quitan las minúsculas)
const abrev = s => String(s).replace(/[a-zñíóéáúü\s]/g, '');
const num = v => (v === null || v === '' || isNaN(Number(v))) ? null : Number(v);
const fmt = v => { const n = num(v); return n === null ? (v === null || v === '' ? '-' : String(v)) : String(Math.round(n * 100) / 100); };

function agruparConsignas(rows){
  const grupos = new Map();
  rows.forEach(r => {
    if (r['Parciales'] == null || r['Temas'] == null) return;
    const clave = r['Parciales'] + '|' + r['Temas'];
    if (!grupos.has(clave)) grupos.set(clave, {parcial: String(r['Parciales']), tema: String(r['Temas']), consignas: []});
    grupos.get(clave).consignas.push(String(r['Consignas']));
  });
  return [...grupos.values()].sort((a, b) => (a.parcial + ' ' + a.tema).localeCompare(b.parcial + ' ' + b.tema, 'es', {numeric: true}));
}

function armarEstudiantes(general, grupos){
  const mapa = new Map();
  const get = r => {
    const k = r['Apellido'] + ', ' + r['Nombre'];
    if (!mapa.has(k)) {
      mapa.set(k, {nombre: k, apellido: r['Apellido'], pila: r['Nombre'], resumen: [], evals: []});
      general.filter(x => x['Apellido'] === r['Apellido'] && x['Nombre'] === r['Nombre'])
        .forEach(x => mapa.get(k).resumen.push({prom: x['Promedio'], obs: x['Observaciones']}));
    }
    return mapa.get(k);
  };
  for (const g of grupos) {
    const id = abrev(g.parcial), tag = id + ';' + abrev(g.tema);
    if (general.length && !(id + ';Tema' in general[0]))
      throw new Error(`Falta la columna «${id};Tema» en la planilla de notas (parcial «${g.parcial}»).`);
    for (const r of general) {
      const prom = num(r['Promedio']);
      if (prom === null || prom < 4) continue;
      const t = String(r[id + ';Tema']).replace(/\s/g, '').toUpperCase();
      if (t !== tag.toUpperCase() && t !== abrev(g.tema).toUpperCase()) continue;
      get(r).evals.push({
        titulo: g.parcial + ' ' + g.tema,
        cabeza: g.consignas,
        valores: g.consignas.map((_, i) => r[`${id};C${i + 1}`]),
        total: r[id + ';Total']
      });
    }
  }
  return [...mapa.values()].sort((a, b) => a.nombre.localeCompare(b.nombre, 'es'));
}

$('go').addEventListener('click', async () => {
  try {
    msg('Procesando…');
    const [general, cons] = await Promise.all([leerHoja(files.gen), leerHoja(files.con)]);
    for (const c of ['Parciales', 'Temas', 'Consignas'])
      if (!cons.length || !(c in cons[0])) throw new Error(`Falta la columna «${c}» en la planilla de consignas.`);
    for (const c of ['Apellido', 'Nombre', 'Promedio'])
      if (!general.length || !(c in general[0])) throw new Error(`Falta la columna «${c}» en la planilla de notas.`);
    estudiantes = armarEstudiantes(general, agruparConsignas(cons));
    if (!estudiantes.length) throw new Error('No se encontró ningún estudiante con promedio de 4 o más y un parcial cargado.');
    datosCursada = `Año ${$('year').value}, Comisión ${$('docente').value}`;
    $('count').textContent = `${estudiantes.length} estudiantes`;
    $('names').innerHTML = '';
    estudiantes.forEach(e => { const li = document.createElement('li'); li.textContent = e.nombre; $('names').appendChild(li); });
    $('result').hidden = false;
    msg('');
  } catch (err) { $('result').hidden = true; msg(err.message, 'err'); }
});

function crearPdf(dosPorHoja){
  const {jsPDF} = window.jspdf;
  const doc = new jsPDF(dosPorHoja ? {unit: 'mm', format: 'a4'} : {unit: 'mm', format: [148.5, 210], orientation: 'landscape'});
  const W = 210, HS = 148.5, M = 10;
  let slot = -1, top = 0, y = 0;
  const nuevo = () => {
    slot++;
    if (dosPorHoja ? (slot % 2 === 0 && slot > 0) : slot > 0) doc.addPage();
    top = dosPorHoja ? (slot % 2) * HS : 0;
    y = top + M + 3;
  };
  const asegurar = h => { if (y + h > top + HS - M) nuevo(); };
  const par = (rotulo, texto) => {
    doc.setFont('helvetica', 'bold'); doc.setFontSize(10.5);
    const w = doc.getTextWidth(rotulo + ' ');
    doc.setFont('helvetica', 'normal');
    const lineas = doc.splitTextToSize(String(texto), W - 2 * M - w);
    asegurar(lineas.length * 5);
    doc.setFont('helvetica', 'bold'); doc.text(rotulo, M, y);
    doc.setFont('helvetica', 'normal'); doc.text(lineas, M + w, y);
    y += lineas.length * 5 + 1;
  };

  estudiantes.forEach(e => {
    nuevo();
    doc.setFont('helvetica', 'bold'); doc.setFontSize(16); doc.text(e.nombre, M, y); y += 6.5;
    doc.setFont('helvetica', 'normal'); doc.setFontSize(10.5); doc.text(datosCursada, M, y); y += 7;
    e.resumen.forEach(r => {
      par('Promedio:', fmt(r.prom));
      if (r.obs !== null && r.obs !== '') par('Observaciones:', r.obs);
    });
    y += 2;
    e.evals.forEach(ev => {
      const n = ev.cabeza.length, cw = (W - 2 * M) / n;
      doc.setFontSize(9);
      doc.setFont('helvetica', 'bold');
      const lin = ev.cabeza.map(c => doc.splitTextToSize(c, cw - 2));
      const hh = Math.max(...lin.map(l => l.length)) * 4 + 3;
      asegurar(8 + hh + 7 + 8);
      doc.setFontSize(12); doc.text(ev.titulo, M, y); y += 3;
      doc.setFontSize(9); doc.setDrawColor(150);
      // Primero se dibujan todas las celdas con fondo y después los textos: en jsPDF,
      // escribir texto cambia el color de relleno activo y las celdas siguientes salían negras.
      doc.setFillColor(232, 236, 245);
      ev.cabeza.forEach((c, i) => doc.rect(M + i * cw, y, cw, hh, 'FD'));
      doc.setTextColor(0, 0, 0);
      ev.cabeza.forEach((c, i) => doc.text(lin[i], M + i * cw + cw / 2, y + 4.5, {align: 'center'}));
      y += hh;
      doc.setFont('helvetica', 'normal');
      ev.valores.forEach((v, i) => {
        doc.rect(M + i * cw, y, cw, 7);
        doc.text(fmt(v), M + i * cw + cw / 2, y + 4.8, {align: 'center'});
      });
      y += 7 + 5;
      par('Total:', fmt(ev.total));
      y += 3;
    });
  });

  if (dosPorHoja) {
    doc.setDrawColor(150); doc.setLineDashPattern([2, 2], 0);
    for (let p = 1; p <= doc.getNumberOfPages(); p++) { doc.setPage(p); doc.line(5, HS, W - 5, HS); }
  }
  return doc;
}

const base = () => ($('year').value + $('docente').value).replace(/[^\wñáéíóú-]+/gi, '') || 'fichas';
$('dl5').addEventListener('click', () => crearPdf(false).save(base() + '.pdf'));
$('dl4').addEventListener('click', () => crearPdf(true).save(base() + '-A4.pdf'));
