'use strict';
// Simulador mínimo de la API de Word (Office.js) para probar el script de la skill
// formatear-curs-clinic-argos sin abrir Word.
//
// Imita lo que el script usa: párrafos cargados con load()+sync() (los valores cargados
// son una "foto" que no cambia al escribir, salvo las propiedades que el propio script
// asigna, como en Office.js), insertParagraph, delete por referencia, rangos vivos que se
// desplazan al editar texto, search (incluido "^l"), getOoxml/insertOoxml, estilos,
// inlinePictures y campos TOC.

const NOMBRE = { Normal: 'Normal', Heading1: 'Título 1', Heading2: 'Título 2', Heading3: 'Título 3',
  Heading4: 'Título 4', Heading5: 'Título 5', ListBullet: 'Lista con viñetas', ListNumber: 'Lista con números',
  Toc1: 'TDC 1', Toc2: 'TDC 2', Toc3: 'TDC 3' };
const BUILTIN = Object.fromEntries(Object.entries(NOMBRE).map(([k, v]) => [v, k]));
const builtinDe = style => BUILTIN[style] || 'Other';

let sigId = 1;
class Nodo {
  constructor(spec = {}) {
    this.id = sigId++;
    this.text = spec.text || '';
    this.style = spec.style || 'Normal';
    this.font = Object.assign({}, spec.font || {});   // solo formato directo; lo demás viene del estilo
    this.pf = Object.assign({}, spec.pf || {});
    this.isListItem = !!spec.list;
    this.img = spec.img || 0;
    this.chars = (spec.chars || []).map(c => Object.assign({}, c));
    this.rangos = new Set();
    this.borrado = false;
  }
  clon() {
    return { text: this.text, style: this.style, font: Object.assign({}, this.font), pf: Object.assign({}, this.pf),
      list: this.isListItem, img: this.img, chars: this.chars.map(c => Object.assign({}, c)) };
  }
}

class Doc {
  constructor(specs) {
    this.nodos = specs.map(s => new Nodo(s));
    this.estilos = { 'Lista con viñetas': {}, 'MiniEspacio': {} };
    this.toc = { n: 1, actualizado: 0 };
    this.log = [];
  }
  idx(n) { const i = this.nodos.indexOf(n); if (i < 0) throw new Error('ItemNotFound: párrafo borrado'); return i; }
  vivo(n) { if (n.borrado) throw new Error('ItemNotFound: párrafo borrado'); }
  insertar(ref, spec, loc) {
    this.vivo(ref);
    const n = new Nodo(spec);
    const i = this.idx(ref);
    this.nodos.splice(loc === 'Before' ? i : i + 1, 0, n);
    return n;
  }
  borrar(n) { this.vivo(n); this.nodos.splice(this.idx(n), 1); n.borrado = true; }
  editar(n, s, e, t) {
    this.vivo(n);
    n.text = n.text.slice(0, s) + t + n.text.slice(e);
    const d = t.length - (e - s);
    const mueve = x => (x >= e ? x + d : x > s ? s + (x - s <= t.length ? x - s : t.length) : x);
    for (const r of n.rangos) { r.s = mueve(r.s); r.e = mueve(r.e); }
    for (const c of n.chars) { c.s = mueve(c.s); c.e = mueve(c.e); }
    n.chars = n.chars.filter(c => c.e > c.s);
  }
}

// ── rangos ──
class RangoTexto {
  constructor(doc, nodo, s, e) { this.doc = doc; this.nodo = nodo; this.s = s; this.e = e; nodo.rangos.add(this); }
  get font() {
    const r = this;
    const set = (k, v) => { r.doc.vivo(r.nodo); r.nodo.chars.push({ s: r.s, e: r.e, k, v }); };
    return new Proxy({}, { set(_, k, v) { set(k, v); return true; } });
  }
  set styleBuiltIn(v) { this.doc.vivo(this.nodo); this.nodo.chars.push({ s: this.s, e: this.e, k: 'estilo', v }); }
  getRange(loc = 'Whole') {
    if (loc === 'Start') return new RangoTexto(this.doc, this.nodo, this.s, this.s);
    if (loc === 'End' || loc === 'After') return new RangoTexto(this.doc, this.nodo, this.e, this.e);
    if (loc === 'Whole' || loc === 'Content') return new RangoTexto(this.doc, this.nodo, this.s, this.e);
    throw new Error('InvalidArgument: Range.getRange(' + loc + ')');
  }
  expandTo(o) {
    if (o instanceof RangoTexto && o.nodo === this.nodo)
      return new RangoTexto(this.doc, this.nodo, Math.min(this.s, o.s), Math.max(this.e, o.e));
    throw new Error('mock: expandTo entre párrafos distintos no soportado para rangos de texto');
  }
  insertText(t, loc) {
    if (loc !== 'Replace') throw new Error('mock: insertText ' + loc);
    this.doc.editar(this.nodo, this.s, this.e, t);
    this.e = this.s + t.length;
    return this;
  }
  delete() { this.doc.editar(this.nodo, this.s, this.e, ''); }
}

class RangoParrafos {
  constructor(doc, a, b) { this.doc = doc; this.a = a; this.b = b; }
  nodos() {
    const i = this.doc.idx(this.a), j = this.doc.idx(this.b);
    return this.doc.nodos.slice(Math.min(i, j), Math.max(i, j) + 1);
  }
  get font() {
    const r = this;
    return new Proxy({}, { set(_, k, v) { for (const n of r.nodos()) n.font[k] = v; return true; } });
  }
  expandTo(o) {
    if (o instanceof RangoParrafos) {
      const todos = [this.a, this.b, o.a, o.b].sort((x, y) => this.doc.idx(x) - this.doc.idx(y));
      return new RangoParrafos(this.doc, todos[0], todos[3]);
    }
    throw new Error('mock: expandTo mixto no soportado');
  }
  getOoxml() { return { value: JSON.stringify(this.nodos().map(n => n.clon())) }; }
}

// ── colecciones ──
class Coleccion {
  constructor(ctx, fabrica) { this.ctx = ctx; this.fabrica = fabrica; this._items = null; }
  load() { this.ctx.pendientes.push(this); return this; }
  cargar() { this._items = this.fabrica(); }
  get items() { if (!this._items) throw new Error('PropertyNotLoaded: items'); return this._items; }
  getFirst() {
    const l = this.fabrica();
    if (!l.length) throw new Error('ItemNotFound: getFirst');
    return l[0];
  }
}

function proxyParrafo(ctx, nodo, foto) {
  const doc = ctx.doc;
  const cache = foto ? { text: nodo.text, style: nodo.style, styleBuiltIn: builtinDe(nodo.style),
    isListItem: nodo.isListItem, font: { color: nodo.font.color, name: nodo.font.name } } : { font: {} };
  const leer = k => { if (!(k in cache)) throw new Error('PropertyNotLoaded: ' + k); return cache[k]; };
  const font = new Proxy({}, {
    get(_, k) { if (!(k in cache.font)) throw new Error('PropertyNotLoaded: font.' + String(k)); return cache.font[k]; },
    set(_, k, v) { doc.vivo(nodo); nodo.font[k] = v; cache.font[k] = v; return true; }
  });
  const p = {
    _nodo: nodo,
    get text() { return leer('text'); },
    get style() { return leer('style'); },
    // en Word, aplicar un estilo de párrafo quita el formato de fuente directo de todo el párrafo
    set style(v) { doc.vivo(nodo); nodo.style = v; nodo.font = {}; cache.style = v; },
    get styleBuiltIn() { return leer('styleBuiltIn'); },
    set styleBuiltIn(v) { doc.vivo(nodo); nodo.style = NOMBRE[v] || v; nodo.font = {}; cache.styleBuiltIn = v; },
    get isListItem() { return leer('isListItem'); },
    get font() { return font; },
    getRange(loc = 'Whole') {
      doc.vivo(nodo);
      if (loc === 'Whole') return new RangoParrafos(doc, nodo, nodo);
      if (loc === 'Content') return new RangoTexto(doc, nodo, 0, nodo.text.length);
      if (loc === 'Start') return new RangoTexto(doc, nodo, 0, 0);
      if (loc === 'End') return new RangoTexto(doc, nodo, nodo.text.length, nodo.text.length);
      throw new Error('mock: Paragraph.getRange ' + loc);
    },
    insertParagraph(t, loc) {
      const n = doc.insertar(nodo, { text: t, style: nodo.style, font: nodo.font, pf: nodo.pf }, loc);
      return proxyParrafo(ctx, n, false);
    },
    delete() { doc.borrar(nodo); },
    detachFromList() { doc.vivo(nodo); nodo.isListItem = false; },
    insertOoxml(xml, loc) {
      if (loc !== 'Replace') throw new Error('mock: insertOoxml ' + loc);
      const specs = JSON.parse(xml);
      const i = doc.idx(nodo);
      const nuevos = specs.map(s => new Nodo(s));
      doc.nodos.splice(i, 1, ...nuevos);
      nodo.borrado = true;
    },
    search(str, opts = {}) {
      return new Coleccion(ctx, () => {
        doc.vivo(nodo);
        const res = [];
        const t = nodo.text;
        if (str === '^l') {
          for (let i = 0; i < t.length; i++) if (t[i] === '\u000b') res.push(new RangoTexto(doc, nodo, i, i + 1));
          return res;
        }
        if (str.length > 255) throw new Error('InvalidArgument: search > 255');
        const hay = opts.matchCase ? t : t.toLowerCase(), aguja = opts.matchCase ? str : str.toLowerCase();
        for (let i = hay.indexOf(aguja); i >= 0; i = hay.indexOf(aguja, i + 1))
          res.push(new RangoTexto(doc, nodo, i, i + str.length));
        return res;
      });
    },
    get inlinePictures() {
      return new Coleccion(ctx, () => { doc.vivo(nodo); return Array.from({ length: nodo.img }, () => ({ width: 100 })); });
    },
  };
  for (const k of ['leftIndent', 'firstLineIndent', 'spaceBefore', 'spaceAfter', 'lineSpacing', 'alignment'])
    Object.defineProperty(p, k, { set(v) { doc.vivo(nodo); nodo.pf[k] = v; }, get() { throw new Error('mock: leer ' + k); } });
  return p;
}

function contexto(doc, opciones = {}) {
  const ctx = { doc, pendientes: [], syncs: 0 };
  const estilo = n => {
    const existe = n in doc.estilos && !doc.estilos[n].fantasma;
    const reg = doc.estilos[n] || (doc.estilos[n] = { fantasma: true });
    return {
      load() { return this; }, isNullObject: !existe,
      font: new Proxy({}, { set(_, k, v) { reg['font.' + k] = v; return true; } }),
      paragraphFormat: new Proxy({}, { set(_, k, v) { reg['pf.' + k] = v; return true; } }),
    };
  };
  ctx.document = {
    getStyles: () => ({ getByNameOrNullObject: estilo }),
    body: {
      get paragraphs() { return new Coleccion(ctx, () => doc.nodos.map(n => proxyParrafo(ctx, n, true))); },
      get inlinePictures() {
        return new Coleccion(ctx, () => doc.nodos.flatMap(n => Array.from({ length: n.img }, () => ({ width: 100 }))));
      },
      get fields() {
        return {
          getByTypes(tipos) {
            if (opciones.sinCampos) throw new Error('ApiNotFound: fields');
            return new Coleccion(ctx, () => tipos.includes('TOC') && doc.toc.n
              ? [{ updateResult() { doc.toc.actualizado++; } }] : []);
          }
        };
      },
    },
  };
  ctx.sync = async () => {
    ctx.syncs++;
    const p = ctx.pendientes; ctx.pendientes = [];
    for (const c of p) c.cargar();
  };
  return ctx;
}

async function ejecutar(codigo, doc, opciones = {}) {
  global.Office = { context: { requirements: { isSetSupported: () => opciones.api15 !== false } } };
  global.Word = {};
  const AsyncFunction = Object.getPrototypeOf(async function () {}).constructor;
  const ctx = contexto(doc, opciones);
  const log = [];
  const consola = { log: (...a) => log.push(a.join(' ')) };
  const fn = new AsyncFunction('context', 'console', codigo);
  const ret = await fn(ctx, consola);
  return { ret: ret !== undefined ? ret : log.join('\n'), syncs: ctx.syncs };
}

function volcar(doc) {
  return doc.nodos.map(n => {
    const f = Object.entries(n.font).map(([k, v]) => k + '=' + v).join(',');
    const pf = Object.entries(n.pf).map(([k, v]) => k.replace(/(Indent|Spac(e|ing))/, '') + '=' + v).join(',');
    const ch = n.chars.map(c => `{${c.s}-${c.e} ${c.k}=${c.v}}`).join('');
    return `${n.style.padEnd(17)}|${n.isListItem ? ' [lista]' : ''}${n.img ? ' [IMG]' : ''} ${n.text.replace(/\u000b/g, '⏎')} ${ch}` +
      `\n${' '.repeat(17)}|   font{${f}} pf{${pf}}`;
  }).join('\n');
}

module.exports = { Doc, Nodo, ejecutar, volcar };
