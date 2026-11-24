'use strict';

const PDF_WORKER = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
// Cambiar la versión cuando cambie la lectura de PDFs, para descartar los datos guardados con la lectura anterior
const CLAVE_GUARDADO = 'buscador-loq:pdfs:v2';
const ORDEN_TIPOS = ['GC-MSMS', 'LC-MSMS', 'AGF'];
const RE_COL_ITEM = /mol[eé]cula|[ií]tem|compon|compuesto|analit|analyte|par[aá]metro|parameter|plaguicida|pesticid|sustancia|substance|nombre|name|naam/i;
const RE_COL_LOQ = /\bloq\b|l[ií]mite|limit|\blmr\b|\bmrl\b|valor|value|mg\s*\/\s*kg/i;

const estado = {
  pdfs: [],
  excel: null,
  resultados: []
};

/* ---------- Utilidades ---------- */

const $ = selector => document.querySelector(selector);

function esc(texto) {
  return String(texto ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

// Normaliza un nombre para comparar: sin tildes, minúsculas, guiones y espacios unificados
function normalizar(texto) {
  return String(texto ?? '')
    .normalize('NFD').replace(/[̀-ͯ]/g, '')
    .toLowerCase()
    .replace(/[‐-―]/g, '-')
    .replace(/(\d)\.(?=\d)/g, '$1,')
    .replace(/[^a-z0-9,.+\s-]/g, ' ')
    .replace(/[-\s]+/g, ' ')
    .trim();
}

function sinPuntuacion(texto) {
  return texto.replace(/[,.+]/g, ' ').replace(/\s+/g, ' ').trim();
}

function aNumero(valor) {
  if (typeof valor === 'number') return Number.isFinite(valor) ? valor : null;
  const m = String(valor ?? '').replace(',', '.').match(/\d+(?:\.\d+)?(?:e-?\d+)?/i);
  return m ? parseFloat(m[0]) : null;
}

function formatoNumero(n) {
  return n == null ? '—' : String(+n.toPrecision(10));
}

/* ---------- Lectura de los PDFs ---------- */

// "0.01", "0.01 r", "0.05 mr", "0.02 CS2", "Q 0.01"
function leerValor(texto) {
  const m = texto.match(/^(?:Q\s+)?(\d+(?:[.,]\d+)?)(?:\s+([A-Za-z0-9]{1,4}))?$/);
  return m ? { num: parseFloat(m[1].replace(',', '.')), nota: m[2] || '' } : null;
}

function limpiarNombre(texto) {
  return texto.replace(/\s*\*+\s*$/, '').replace(/\s+Q$/, '').replace(/\s+/g, ' ').trim();
}

function agruparFilas(items, tolerancia = 2.5) {
  const ordenados = [...items].sort((a, b) => b.y - a.y || a.x - b.x);
  const filas = [];
  for (const item of ordenados) {
    const fila = filas[filas.length - 1];
    if (fila && Math.abs(fila.y - item.y) <= tolerancia) fila.items.push(item);
    else filas.push({ y: item.y, items: [item] });
  }
  filas.forEach(f => f.items.sort((a, b) => a.x - b.x));
  return filas;
}

function detectarTipo(cabecera) {
  if (/\bGC[\s-]?MS/i.test(cabecera)) return 'GC-MSMS';
  if (/\bLC[\s-]?MS/i.test(cabecera)) return 'LC-MSMS';
  if (/\bAGF\b/.test(cabecera)) return 'AGF';
  return 'Otro';
}

// Páginas GC / LC: filas con varias columnas "nombre [Q] valor"
function parsearLista(cuerpo) {
  const salida = [];
  for (const fila of agruparFilas(cuerpo)) {
    let partes = [];
    let ultima = null;
    for (const item of fila.items) {
      if (item.s === 'Q') continue;
      // Marca suelta ("r", "m", "mr") que pertenece al valor anterior, no al nombre siguiente
      if (/^(m|r|mr|rm)$/.test(item.s) && !partes.length && ultima) {
        ultima.nota = (ultima.nota + ' ' + item.s).trim();
        continue;
      }
      const valor = leerValor(item.s);
      if (!valor) { partes.push(item.s); continue; }
      if (partes.length) {
        ultima = { nombre: limpiarNombre(partes.join(' ')), loq: valor.num, nota: valor.nota, lista: false };
        salida.push(ultima);
      }
      partes = [];
    }
  }
  return salida;
}

// Páginas AGF: tabla Componente | Q | Método | Límite, con encabezados en negrita y sub-items
function parsearTabla(cuerpo, columnas, fuenteNegrita) {
  const filas = agruparFilas(cuerpo.filter(i => i.x < columnas.q - 5)).map(f => ({
    y: f.y,
    s: f.items.map(i => i.s).join(' '),
    cab: f.items.some(i => i.f === fuenteNegrita),
    val: null
  }));

  for (const item of cuerpo) {
    if (item.x < columnas.lim - 5) continue;
    const valor = leerValor(item.s);
    if (!valor) continue;
    let mejor = null;
    for (const f of filas) {
      const d = Math.abs(f.y - item.y);
      if (d <= 8 && !f.val && (!mejor || d < mejor.d)) mejor = { f, d };
    }
    if (mejor) mejor.f.val = valor;
  }

  const salida = [];
  let cab = null;
  let subs = [];
  const cerrarBloque = () => {
    const loqCab = cab && cab.val;
    if (loqCab) salida.push({ nombre: limpiarNombre(cab.s), loq: loqCab.num, nota: loqCab.nota, lista: true });
    for (const sub of subs) {
      const v = sub.val || loqCab;
      if (v) salida.push({ nombre: limpiarNombre(sub.s), loq: v.num, nota: v.nota, lista: true });
    }
  };

  for (const f of filas) {
    if (f.cab) {
      cerrarBloque();
      cab = f;
      subs = [];
      continue;
    }
    const anterior = subs[subs.length - 1];
    if (anterior && !f.val && /[,-]$/.test(anterior.s)) {
      anterior.s += (anterior.s.endsWith('-') ? '' : ' ') + f.s;
    } else {
      subs.push(f);
    }
  }
  cerrarBloque();
  return salida;
}

// items: [{ s: texto, x, y, f: fuente }] de una página
function parsearPagina(crudos) {
  const items = crudos
    .map(i => ({ ...i, s: String(i.s).replace(/\s+/g, ' ').trim() }))
    .filter(i => i.s);

  const leyenda = items.filter(i => /^Q\s*:/.test(i.s)).sort((a, b) => b.y - a.y)[0];
  const pieY = leyenda ? leyenda.y + 1 : -Infinity;
  const cabComponente = items.find(i => /^(componente|component)$/i.test(i.s));
  const cabMetodo = cabComponente && items.find(i => /m[eé]todo|method/i.test(i.s) && Math.abs(i.y - cabComponente.y) < 10);
  const esTabla = Boolean(cabComponente && cabMetodo);
  const inicioLista = items.find(i => /^(lista de componentes|list of components)/i.test(i.s));
  const topeY = esTabla ? cabComponente.y : inicioLista ? inicioLista.y : Infinity;

  const cabecera = items.filter(i => i.y > topeY + 1).map(i => i.s).join(' ');
  const tipo = detectarTipo(cabecera);
  const cuerpo = items.filter(i => i.y < topeY - (esTabla ? 9 : 1) && i.y > pieY);

  if (!esTabla) return { tipo, entradas: parsearLista(cuerpo) };

  const filaCab = items.filter(i => Math.abs(i.y - cabComponente.y) < 10 && i.x > cabComponente.x);
  const colQ = filaCab.find(i => i.s === 'Q');
  const colLim = filaCab.filter(i => i.x > cabMetodo.x + 10).sort((a, b) => a.x - b.x)[0];
  const columnas = { q: colQ ? colQ.x : cabMetodo.x, lim: colLim ? colLim.x : Infinity };
  return { tipo, entradas: parsearTabla(cuerpo, columnas, cabComponente.f) };
}

async function leerPdf(archivo) {
  const datos = new Uint8Array(await archivo.arrayBuffer());
  const doc = await pdfjsLib.getDocument({ data: datos }).promise;
  const entradas = [];
  for (let p = 1; p <= doc.numPages; p++) {
    const pagina = await doc.getPage(p);
    const contenido = await pagina.getTextContent();
    const items = contenido.items.map(i => ({ s: i.str, x: i.transform[4], y: i.transform[5], f: i.fontName }));
    const { tipo, entradas: filas } = parsearPagina(items);
    filas.forEach(e => entradas.push({ n: e.nombre, t: tipo, l: e.loq, o: e.nota, p, li: e.lista ? 1 : 0 }));
  }
  await doc.destroy();
  return { nombre: archivo.name, tam: archivo.size, paginas: doc.numPages, entradas };
}

/* ---------- Búsqueda ---------- */

function clavesDe(nombre, esLista) {
  const claves = new Set();
  const agregar = t => {
    const n = normalizar(t);
    if (n.length > 1) claves.add(n);
  };
  const base = nombre.replace(/^(suma de|sum of)\s*:\s*/i, '');
  agregar(base);
  const sinParentesis = base.replace(/\([^)]*\)/g, ' ');
  agregar(sinParentesis);
  if (esLista) sinParentesis.split(/,\s+|\s+\/\s+/).forEach(agregar);
  return claves;
}

function construirIndice(pdfs) {
  return pdfs.flatMap(pdf => pdf.entradas.map(e => {
    const norm = normalizar(e.n);
    return { ...e, archivo: pdf.nombre, norm, normPlano: ' ' + sinPuntuacion(norm) + ' ', claves: clavesDe(e.n, e.li) };
  }));
}

function estadoDe(loqPdf, pedido) {
  if (loqPdf == null) return 'no-encontrado';
  if (pedido == null) return 'sin-loq';
  return loqPdf <= pedido * (1 + 1e-9) ? 'cumple' : 'no-cumple';
}

function buscarItem(indice, nombre, pedido) {
  const q = normalizar(nombre);
  let encontrados = indice.filter(e => e.claves.has(q));
  let parcial = false;
  if (!encontrados.length && q.length >= 4) {
    const qPlano = ' ' + sinPuntuacion(q) + ' ';
    encontrados = indice.filter(e => e.normPlano.includes(qPlano));
    parcial = encontrados.length > 0;
  }

  if (!encontrados.length) {
    return [{ molecula: nombre, tipo: '—', loqPdf: null, loqPedido: pedido, estado: 'no-encontrado', archivos: [] }];
  }

  // Un resultado por PDF + tipo + LOQ: se revisan todos los PDFs, no solo el primero que coincide
  const grupos = new Map();
  for (const e of encontrados) {
    const clave = e.archivo + '|' + e.t + '|' + e.l;
    const g = grupos.get(clave);
    if (!g) {
      grupos.set(clave, { tipo: e.t, loq: e.l, nota: e.o, nombrePdf: e.n, exacto: e.norm === q, pagina: e.p, archivos: [e.archivo] });
    } else if (!g.exacto && e.norm === q) {
      Object.assign(g, { nombrePdf: e.n, exacto: true, pagina: e.p, nota: e.o });
    }
  }

  const ordenPdf = new Map();
  indice.forEach(e => { if (!ordenPdf.has(e.archivo)) ordenPdf.set(e.archivo, ordenPdf.size); });
  const posicion = t => (ORDEN_TIPOS.includes(t) ? ORDEN_TIPOS.indexOf(t) : ORDEN_TIPOS.length);
  return [...grupos.values()]
    .sort((a, b) => ordenPdf.get(a.archivos[0]) - ordenPdf.get(b.archivos[0]) || posicion(a.tipo) - posicion(b.tipo) || a.loq - b.loq)
    .map(g => ({
      molecula: nombre,
      tipo: g.tipo,
      loqPdf: g.loq,
      nota: g.nota,
      loqPedido: pedido,
      estado: estadoDe(g.loq, pedido),
      nombrePdf: g.exacto ? '' : g.nombrePdf,
      pagina: g.pagina,
      archivos: g.archivos,
      parcial
    }));
}

/* ---------- Lectura del Excel ---------- */

function analizarHoja(filas) {
  const inicio = filas.findIndex(r => r.some(c => String(c).trim() !== ''));
  if (inicio < 0) return { cabeceras: [], datos: [], colItem: -1, colLoq: -1 };

  const primera = filas[inicio];
  const siguientes = filas.slice(inicio + 1, inicio + 20);
  const textoCelda = c => String(c ?? '').trim();
  const pareceDato = primera.some(c => typeof c === 'number' || RE_VALOR_EN_NOMBRE.test(textoCelda(c)));
  const pareceCabecera = primera.some(c => RE_COL_ITEM.test(textoCelda(c)) || RE_COL_LOQ.test(textoCelda(c))) ||
    primera.some((c, i) => textoCelda(c) && aNumero(c) == null && siguientes.some(r => typeof r[i] === 'number'));
  const tieneCabecera = !pareceDato && pareceCabecera;
  const datos = filas.slice(tieneCabecera ? inicio + 1 : inicio).filter(r => r.some(c => String(c).trim() !== ''));
  const nCol = Math.max(primera.length, ...datos.slice(0, 200).map(r => r.length));
  const cabeceras = Array.from({ length: nCol }, (_, i) =>
    (tieneCabecera && String(primera[i] ?? '').trim()) || `Columna ${XLSX.utils.encode_col(i)}`);

  const muestra = datos.slice(0, 40);
  const proporcionNumerica = i => {
    const llenas = muestra.filter(r => String(r[i] ?? '').trim() !== '');
    return llenas.length ? llenas.filter(r => aNumero(r[i]) != null && !/[a-z]{3,}/i.test(String(r[i]))).length / llenas.length : 0;
  };

  let colItem = tieneCabecera ? cabeceras.findIndex(h => RE_COL_ITEM.test(h)) : -1;
  let colLoq = tieneCabecera ? cabeceras.findIndex((h, i) => i !== colItem && RE_COL_LOQ.test(h)) : -1;
  if (colItem < 0) colItem = cabeceras.findIndex((_, i) => i !== colLoq && proporcionNumerica(i) < 0.5);
  if (colLoq < 0) colLoq = cabeceras.findIndex((_, i) => i !== colItem && proporcionNumerica(i) >= 0.6);
  return { cabeceras, datos, colItem: Math.max(colItem, 0), colLoq };
}

async function leerExcel(archivo) {
  const libro = XLSX.read(await archivo.arrayBuffer());
  const hojas = {};
  libro.SheetNames.forEach(n => {
    hojas[n] = XLSX.utils.sheet_to_json(libro.Sheets[n], { header: 1, defval: '', raw: true });
  });
  const hoja = libro.SheetNames.find(n => hojas[n].some(r => r.some(c => String(c).trim() !== ''))) || libro.SheetNames[0];
  return { nombre: archivo.name, hojas, hoja, ...analizarHoja(hojas[hoja]) };
}

// "Bendiocarb 0.02", "Bendiocarb: 0,02", "Bendiocarb <0.02 mg/kg" → nombre + valor.
// Solo se separa un número decimal al final, para no romper nombres como "S 421" o "Carbofurano-3-OH".
const RE_VALOR_EN_NOMBRE = /^(.*?[^\s:;=|–-])[\s:;=|–-]+[<≤]?\s*(\d*[.,]\d+)\s*(?:mg\s*\/\s*kg|ppm)?$/i;

function separarNombreValor(texto) {
  const m = texto.match(RE_VALOR_EN_NOMBRE);
  return m ? { nombre: m[1].trim(), valor: parseFloat(m[2].replace(',', '.')) } : { nombre: texto, valor: null };
}

function itemsDelExcel(excel) {
  if (!excel) return [];
  return excel.datos
    .map(r => {
      const { nombre, valor } = separarNombreValor(String(r[excel.colItem] ?? '').replace(/\s+/g, ' ').trim());
      const enColumna = excel.colLoq >= 0 ? aNumero(r[excel.colLoq]) : null;
      return { nombre, pedido: enColumna ?? valor };
    })
    .filter(i => i.nombre);
}

/* ---------- Guardado local ---------- */

function cargarGuardado() {
  try {
    const datos = JSON.parse(localStorage.getItem(CLAVE_GUARDADO) || '[]');
    if (Array.isArray(datos)) estado.pdfs = datos;
  } catch (e) {
    estado.pdfs = [];
  }
}

function guardar() {
  try {
    localStorage.setItem(CLAVE_GUARDADO, JSON.stringify(estado.pdfs));
  } catch (e) {
    avisar('No se pudieron guardar los PDFs en este navegador; seguirán cargados hasta que cierres la página.', 'error');
  }
}

/* ---------- Interfaz ---------- */

function avisar(texto, tipo = '') {
  const el = $('#estado');
  el.textContent = texto;
  el.className = 'estado' + (tipo ? ' estado--' + tipo : '');
}

function actualizarBoton() {
  $('#botonBuscar').disabled = !(estado.pdfs.length && itemsDelExcel(estado.excel).length);
}

function pintarPdfs() {
  const lista = $('#listaPdf');
  lista.innerHTML = estado.pdfs.map((pdf, i) => {
    const tipos = ORDEN_TIPOS.concat('Otro').filter(t => pdf.entradas.some(e => e.t === t));
    return `
      <li class="archivo">
        <span class="archivo__icono archivo__icono--pdf">PDF</span>
        <span class="archivo__info">
          <span class="archivo__nombre" title="${esc(pdf.nombre)}">${esc(pdf.nombre)}</span>
          <span class="archivo__meta">${pdf.entradas.length} componentes · ${tipos.map(t => t.replace('-MSMS', '')).join(' · ') || 'sin tipo'}</span>
        </span>
        <button type="button" class="archivo__quitar" data-quitar-pdf="${i}" aria-label="Quitar ${esc(pdf.nombre)}">&times;</button>
      </li>`;
  }).join('');
  actualizarBoton();
}

function pintarExcel() {
  const caja = $('#infoExcel');
  const excel = estado.excel;
  if (!excel) {
    caja.innerHTML = '';
    actualizarBoton();
    return;
  }
  const items = itemsDelExcel(excel);
  const opciones = (seleccion, conVacio) =>
    (conVacio ? `<option value="-1"${seleccion < 0 ? ' selected' : ''}>Sin LOQ</option>` : '') +
    excel.cabeceras.map((h, i) => `<option value="${i}"${i === seleccion ? ' selected' : ''}>${esc(h)}</option>`).join('');
  const nombresHojas = Object.keys(excel.hojas);
  const ejemplo = items.slice(0, 3).map(i => `${esc(i.nombre)}${i.pedido != null ? ` <b>${formatoNumero(i.pedido)}</b>` : ''}`).join(', ');

  caja.innerHTML = `
    <div class="archivo">
      <span class="archivo__icono archivo__icono--xls">XLS</span>
      <span class="archivo__info">
        <span class="archivo__nombre" title="${esc(excel.nombre)}">${esc(excel.nombre)}</span>
        <span class="archivo__meta">${items.length} items${ejemplo ? ' · ' + ejemplo + (items.length > 3 ? '…' : '') : ''}</span>
      </span>
      <button type="button" class="archivo__quitar" data-quitar-excel aria-label="Quitar Excel">&times;</button>
    </div>
    <div class="columnas">
      ${nombresHojas.length > 1 ? `
      <label class="campo"><span>Hoja</span>
        <select data-campo="hoja">${nombresHojas.map(n => `<option${n === excel.hoja ? ' selected' : ''}>${esc(n)}</option>`).join('')}</select>
      </label>` : ''}
      <label class="campo"><span>Columna de molécula</span>
        <select data-campo="colItem">${opciones(excel.colItem, false)}</select>
      </label>
      <label class="campo"><span>Columna de LOQ pedido</span>
        <select data-campo="colLoq">${opciones(excel.colLoq, true)}</select>
      </label>
    </div>`;
  actualizarBoton();
}

const TEXTO_ESTADO = {
  'cumple': '✅ Cumple',
  'no-cumple': '❌ No cumple',
  'no-encontrado': '❌ No encontrado',
  'sin-loq': '— Sin LOQ'
};

function pintarResultados() {
  const filas = estado.resultados;
  const porItem = new Map();
  filas.forEach(f => {
    const lista = porItem.get(f.indice) || [];
    lista.push(f);
    porItem.set(f.indice, lista);
  });
  const grupos = [...porItem.values()];
  const cuenta = condicion => grupos.filter(condicion).length;
  const cumplen = cuenta(g => g.some(f => f.estado === 'cumple'));
  const noEncontrados = cuenta(g => g[0].estado === 'no-encontrado');
  const noCumplen = cuenta(g => g.some(f => f.estado === 'no-cumple') && !g.some(f => f.estado === 'cumple'));

  $('#resumen').innerHTML = `
    <span class="dato"><b>${grupos.length}</b> buscados</span>
    <span class="dato dato--ok"><b>${cumplen}</b> cumplen</span>
    <span class="dato dato--mal"><b>${noCumplen}</b> no cumplen</span>
    <span class="dato dato--neutro"><b>${noEncontrados}</b> no encontrados</span>`;

  pintarTabla();
  $('#resultados').hidden = false;
}

// Filtra la tabla mientras se escribe: muestra las moléculas cuyo nombre contiene el texto (sin importar tildes ni mayúsculas)
function pintarTabla() {
  const q = normalizar($('#filtro').value);
  const todas = estado.resultados;
  const coincide = new Set();
  if (q) {
    todas.forEach(f => {
      if (normalizar(f.molecula).includes(q)) coincide.add(f.indice);
    });
  }
  const filas = q ? todas.filter(f => coincide.has(f.indice)) : todas;
  const totalItems = new Set(todas.map(f => f.indice)).size;
  $('#filtroCuenta').textContent = q ? `${coincide.size} de ${totalItems} moléculas` : '';

  if (!filas.length) {
    $('#cuerpoTabla').innerHTML = `<tr><td colspan="6" class="tabla__vacia">Ninguna molécula coincide con “${esc($('#filtro').value.trim())}”.</td></tr>`;
    return;
  }

  $('#cuerpoTabla').innerHTML = filas.map(f => {
    const primera = !f.continuacion;
    const detalle = [];
    if (f.nombrePdf) detalle.push(`En el PDF: ${esc(f.nombrePdf)}`);
    if (f.parcial) detalle.push('coincidencia parcial');
    const tipoClase = f.tipo === '—' ? 'vacio' : f.tipo.slice(0, 2).toLowerCase();
    return `
      <tr class="${primera ? 'fila--primera' : 'fila--cont'}">
        <td class="tabla__molecula">
          ${primera ? `<span class="molecula">${esc(f.molecula)}</span>` : ''}
          ${detalle.length ? `<span class="molecula__detalle">${detalle.join(' · ')}</span>` : ''}
        </td>
        <td><span class="tipo tipo--${tipoClase}">${esc(f.tipo)}</span></td>
        <td class="num">${formatoNumero(f.loqPdf)}${f.nota ? ` <small>${esc(f.nota)}</small>` : ''}</td>
        <td class="num">${formatoNumero(f.loqPedido)}</td>
        <td><span class="pildora pildora--${f.estado}">${TEXTO_ESTADO[f.estado]}</span></td>
        <td class="tabla__archivo">
          ${f.archivos.length ? `<span class="pdf" title="${esc(f.archivos.join(', '))}">${esc(f.archivos.join(', '))}</span>
          <span class="pdf__pagina">pág. ${f.pagina}</span>` : '<span class="pdf__pagina">—</span>'}
        </td>
      </tr>`;
  }).join('');
}

async function agregarPdfs(archivos) {
  const pdfs = [...archivos].filter(a => /\.pdf$/i.test(a.name) || a.type === 'application/pdf');
  if (!pdfs.length) return avisar('Solo se aceptan archivos PDF en este campo.', 'error');

  const errores = [];
  for (const [i, archivo] of pdfs.entries()) {
    avisar(`Leyendo ${archivo.name} (${i + 1}/${pdfs.length})…`, 'cargando');
    try {
      const pdf = await leerPdf(archivo);
      if (!pdf.entradas.length) errores.push(`${archivo.name}: no se encontraron componentes`);
      const previo = estado.pdfs.findIndex(p => p.nombre === pdf.nombre);
      if (previo >= 0) estado.pdfs[previo] = pdf;
      else estado.pdfs.push(pdf);
    } catch (e) {
      console.error(e);
      errores.push(`${archivo.name}: no se pudo leer`);
    }
  }
  guardar();
  pintarPdfs();
  if (errores.length) avisar(errores.join(' · '), 'error');
  else avisar(`${pdfs.length === 1 ? 'PDF cargado' : pdfs.length + ' PDFs cargados'}.`, 'ok');
}

async function cargarExcel(archivos) {
  const archivo = [...archivos].find(a => /\.(xlsx|xlsm|xls|csv)$/i.test(a.name));
  if (!archivo) return avisar('Sube un archivo Excel (.xlsx, .xls) o CSV.', 'error');
  try {
    avisar(`Leyendo ${archivo.name}…`, 'cargando');
    estado.excel = await leerExcel(archivo);
    pintarExcel();
    const n = itemsDelExcel(estado.excel).length;
    avisar(n ? `${n} items listos para buscar.` : 'No se encontraron items en el Excel.', n ? 'ok' : 'error');
  } catch (e) {
    console.error(e);
    avisar(`No se pudo leer ${archivo.name}.`, 'error');
  }
}

function buscar() {
  const items = itemsDelExcel(estado.excel);
  if (!estado.pdfs.length || !items.length) return;
  const indice = construirIndice(estado.pdfs);
  estado.resultados = items.flatMap((item, indiceItem) =>
    buscarItem(indice, item.nombre, item.pedido).map((fila, j) => ({ ...fila, indice: indiceItem, continuacion: j > 0 })));
  $('#filtro').value = '';
  pintarResultados();
  avisar(`Búsqueda completada: ${items.length} items.`, 'ok');
  $('#resultados').scrollIntoView({ behavior: 'smooth', block: 'start' });
}

/* ---------- Exportación a Excel con estilos (ExcelJS, se carga solo al exportar) ---------- */

const EXCELJS_URL = 'https://cdn.jsdelivr.net/npm/exceljs@4.4.0/dist/exceljs.min.js';
const MARCA = {
  naranja: 'FFEE7203',
  naranjaSuave: 'FFFFF4EA',
  grisOscuro: 'FF4A4A4F',
  grisTexto: 'FF6B6B70',
  borde: 'FFE8E6E3',
  blanco: 'FFFFFFFF'
};
const ESTILO_ESTADO = {
  'cumple': { fondo: 'FFE8F5EE', texto: 'FF16794A' },
  'no-cumple': { fondo: 'FFFDECEA', texto: 'FFB42318' },
  'no-encontrado': { fondo: 'FFFDECEA', texto: 'FFB42318' },
  'sin-loq': { fondo: 'FFF7F6F4', texto: 'FF6B6B70' }
};

let cargaExcelJS = null;
function cargarExcelJS() {
  if (window.ExcelJS) return Promise.resolve();
  cargaExcelJS = cargaExcelJS || new Promise((resolver, rechazar) => {
    const script = document.createElement('script');
    script.src = EXCELJS_URL;
    script.onload = resolver;
    script.onerror = () => {
      cargaExcelJS = null;
      rechazar(new Error('No se pudo cargar ExcelJS'));
    };
    document.head.appendChild(script);
  });
  return cargaExcelJS;
}

function fechaHoy(separador = '-', invertida = false) {
  const hoy = new Date();
  const partes = [hoy.getFullYear(), hoy.getMonth() + 1, hoy.getDate()].map(n => String(n).padStart(2, '0'));
  return (invertida ? partes.reverse() : partes).join(separador);
}

function descargar(blob, nombre) {
  const url = URL.createObjectURL(blob);
  const enlace = document.createElement('a');
  enlace.href = url;
  enlace.download = nombre;
  document.body.appendChild(enlace);
  enlace.click();
  enlace.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

async function exportar() {
  if (!estado.resultados.length) return;
  const boton = $('#botonExportar');
  boton.disabled = true;
  try {
    await cargarExcelJS();
    const blob = await crearExcel(estado.resultados);
    descargar(blob, `resultados-loq-${fechaHoy()}.xlsx`);
  } catch (e) {
    console.error(e);
    avisar('No se pudo generar el Excel. Revisa tu conexión a internet e inténtalo de nuevo.', 'error');
  } finally {
    boton.disabled = false;
  }
}

async function crearExcel(resultados) {
  const libro = new ExcelJS.Workbook();
  libro.creator = 'Normec Groen Agro Control Peru';
  libro.created = new Date();

  const hoja = libro.addWorksheet('Resultados', {
    views: [{ state: 'frozen', ySplit: 4, showGridLines: false }],
    pageSetup: { orientation: 'portrait', fitToPage: true, fitToWidth: 1, fitToHeight: 0 }
  });
  hoja.columns = [
    { key: 'molecula', width: 38 },
    { key: 'tipo', width: 13 },
    { key: 'loqNormec', width: 18 },
    { key: 'loqSolicitado', width: 20 },
    { key: 'cumple', width: 19 },
    { key: 'archivo', width: 48 }
  ];

  const borde = { style: 'thin', color: { argb: MARCA.borde } };
  const bordes = { top: borde, left: borde, bottom: borde, right: borde };
  const relleno = argb => ({ type: 'pattern', pattern: 'solid', fgColor: { argb } });

  // Fila 1: título con la marca
  hoja.mergeCells('A1:F1');
  const titulo = hoja.getCell('A1');
  titulo.value = 'Normec Groen Agro Control Peru  ·  Validación de LOQ';
  titulo.font = { name: 'Calibri', size: 15, bold: true, color: { argb: MARCA.blanco } };
  titulo.fill = relleno(MARCA.naranja);
  titulo.alignment = { vertical: 'middle', indent: 1 };
  hoja.getRow(1).height = 32;

  // Fila 2: resumen
  const porItem = new Map();
  resultados.forEach(f => porItem.set(f.indice, [...(porItem.get(f.indice) || []), f]));
  const grupos = [...porItem.values()];
  const cumplen = grupos.filter(g => g.some(f => f.estado === 'cumple')).length;
  const noEncontrados = grupos.filter(g => g[0].estado === 'no-encontrado').length;
  const noCumplen = grupos.filter(g => g.some(f => f.estado === 'no-cumple') && !g.some(f => f.estado === 'cumple')).length;
  hoja.mergeCells('A2:F2');
  const resumen = hoja.getCell('A2');
  resumen.value = `Generado el ${fechaHoy('/', true)}  ·  ${grupos.length} moléculas  ·  ${cumplen} cumplen  ·  ${noCumplen} no cumplen  ·  ${noEncontrados} no encontradas  ·  LOQ en mg/kg`;
  resumen.font = { name: 'Calibri', size: 10, color: { argb: MARCA.grisTexto } };
  resumen.fill = relleno(MARCA.naranjaSuave);
  resumen.alignment = { vertical: 'middle', indent: 1 };
  hoja.getRow(2).height = 22;

  hoja.getRow(3).height = 8;

  // Fila 4: cabecera
  const cabecera = hoja.getRow(4);
  cabecera.values = ['Molécula', 'Tipo', 'LOQ NORMEC', 'LOQ Solicitado', 'Cumple', 'Archivo'];
  cabecera.height = 24;
  cabecera.eachCell((celda, col) => {
    const izquierda = col === 1 || col === 6;
    celda.font = { name: 'Calibri', size: 11, bold: true, color: { argb: MARCA.blanco } };
    celda.fill = relleno(MARCA.grisOscuro);
    celda.alignment = { vertical: 'middle', horizontal: izquierda ? 'left' : 'center', indent: izquierda ? 1 : 0 };
    celda.border = { ...bordes, bottom: { style: 'medium', color: { argb: MARCA.naranja } } };
  });

  // Datos: una franja de color por molécula y su nombre en una sola celda combinada
  let franja = false;
  let inicioGrupo = 5;
  const combinarGrupo = fin => {
    if (fin > inicioGrupo) hoja.mergeCells(inicioGrupo, 1, fin, 1);
  };
  resultados.forEach((f, i) => {
    if (!f.continuacion) {
      if (i > 0) combinarGrupo(4 + i);
      inicioGrupo = 5 + i;
      franja = !franja;
    }
    const fila = hoja.addRow({
      molecula: f.molecula,
      tipo: f.tipo,
      loqNormec: f.loqPdf ?? '—',
      loqSolicitado: f.loqPedido ?? '—',
      cumple: TEXTO_ESTADO[f.estado],
      archivo: f.archivos.map(a => a.replace(/\.pdf$/i, '')).join(', ') || '—'
    });
    fila.height = 20;
    const fondo = franja ? MARCA.blanco : 'FFFBFAF9';
    fila.eachCell({ includeEmpty: true }, (celda, col) => {
      celda.font = { name: 'Calibri', size: 11, color: { argb: 'FF1C1C1E' }, bold: col === 1 };
      celda.fill = relleno(fondo);
      celda.border = bordes;
      const horizontal = col === 1 || col === 6 ? 'left' : col === 2 || col === 5 ? 'center' : 'right';
      celda.alignment = { vertical: 'middle', horizontal, indent: col === 1 || col === 6 ? 1 : 0 };
      if (col === 6) celda.font = { name: 'Calibri', size: 10, color: { argb: MARCA.grisTexto } };
      if ((col === 3 || col === 4) && typeof celda.value === 'number') celda.numFmt = '0.0##';
    });
    const tipo = fila.getCell(2);
    tipo.font = { name: 'Calibri', size: 10, bold: true, color: { argb: f.tipo === 'GC-MSMS' ? 'FF2C5A9E' : f.tipo === 'LC-MSMS' ? 'FF6340A8' : MARCA.grisOscuro } };
    const cumple = fila.getCell(5);
    const color = ESTILO_ESTADO[f.estado];
    cumple.fill = relleno(color.fondo);
    cumple.font = { name: 'Calibri', size: 11, bold: true, color: { argb: color.texto } };
  });
  combinarGrupo(4 + resultados.length);

  hoja.autoFilter = { from: { row: 4, column: 1 }, to: { row: 4 + resultados.length, column: 6 } };

  const buffer = await libro.xlsx.writeBuffer();
  return new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
}

function prepararZona(zona, input, alSoltar) {
  const caja = zona.querySelector('.zona__caja');
  input.addEventListener('change', () => {
    if (input.files.length) alSoltar(input.files);
    input.value = '';
  });
  ['dragenter', 'dragover'].forEach(ev => caja.addEventListener(ev, e => {
    e.preventDefault();
    caja.classList.add('zona__caja--encima');
  }));
  ['dragleave', 'drop'].forEach(ev => caja.addEventListener(ev, e => {
    e.preventDefault();
    caja.classList.remove('zona__caja--encima');
  }));
  caja.addEventListener('drop', e => {
    if (e.dataTransfer.files.length) alSoltar(e.dataTransfer.files);
  });
}

function iniciar() {
  if (!window.pdfjsLib || !window.XLSX) {
    avisar('No se pudieron cargar las librerías de lectura. Revisa tu conexión a internet y recarga la página.', 'error');
    return;
  }
  pdfjsLib.GlobalWorkerOptions.workerSrc = PDF_WORKER;

  cargarGuardado();
  pintarPdfs();
  pintarExcel();

  prepararZona($('#zonaPdf'), $('#inputPdf'), agregarPdfs);
  prepararZona($('#zonaExcel'), $('#inputExcel'), cargarExcel);

  $('#listaPdf').addEventListener('click', e => {
    const boton = e.target.closest('[data-quitar-pdf]');
    if (!boton) return;
    estado.pdfs.splice(Number(boton.dataset.quitarPdf), 1);
    guardar();
    pintarPdfs();
  });

  $('#infoExcel').addEventListener('click', e => {
    if (!e.target.closest('[data-quitar-excel]')) return;
    estado.excel = null;
    pintarExcel();
  });

  $('#infoExcel').addEventListener('change', e => {
    const campo = e.target.dataset.campo;
    if (!campo || !estado.excel) return;
    if (campo === 'hoja') {
      estado.excel.hoja = e.target.value;
      Object.assign(estado.excel, analizarHoja(estado.excel.hojas[e.target.value]));
    } else {
      estado.excel[campo] = Number(e.target.value);
    }
    pintarExcel();
  });

  $('#formulario').addEventListener('submit', e => {
    e.preventDefault();
    buscar();
  });

  $('#botonExportar').addEventListener('click', exportar);

  let pendiente = 0;
  $('#filtro').addEventListener('input', () => {
    cancelAnimationFrame(pendiente);
    pendiente = requestAnimationFrame(pintarTabla);
  });
}

document.addEventListener('DOMContentLoaded', iniciar);
