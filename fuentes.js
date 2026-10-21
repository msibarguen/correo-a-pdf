/*!
 * Correo a PDF — fuentes TrueType incrustadas (v0.2.0)
 *
 * Lee archivos .ttf (tablas head, hhea, hmtx, cmap, OS/2, post) y los incrusta en
 * pdf-lib como fuentes CID (Type0 / CIDFontType2, Identity-H) con ToUnicode, de
 * modo que el texto del PDF es vectorial, seleccionable y buscable.
 * No depende de fontkit.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory(require('pdf-lib'));
  else root.CorreoFuentes = factory(root.PDFLib);
})(typeof self !== 'undefined' ? self : this, function (PDFLib) {
  'use strict';
  const { PDFName, PDFHexString, PDFString } = PDFLib;

  // Familias disponibles (archivos en /fuentes) y estilos: r, b, i, bi
  const FAMILIAS = ['carlito', 'caladea', 'libsans', 'libserif', 'libmono', 'dejavu'];
  const ESTILOS = ['r', 'b', 'i', 'bi'];
  const NOMBRE_PS = {
    carlito: 'Carlito', caladea: 'Caladea', libsans: 'LiberationSans', libserif: 'LiberationSerif',
    libmono: 'LiberationMono', dejavu: 'DejaVuSans'
  };

  // ------------------------------------------------------------------ lector TTF

  function leerTTF(bytes) {
    const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    const u16 = o => dv.getUint16(o), i16 = o => dv.getInt16(o), u32 = o => dv.getUint32(o);
    const numTablas = u16(4);
    const tablas = {};
    for (let i = 0; i < numTablas; i++) {
      const o = 12 + i * 16;
      const tag = String.fromCharCode(bytes[o], bytes[o + 1], bytes[o + 2], bytes[o + 3]);
      tablas[tag] = { off: u32(o + 8), len: u32(o + 12) };
    }
    const req = t => { if (!tablas[t]) throw new Error('Fuente sin tabla ' + t); return tablas[t].off; };

    const head = req('head');
    const unidades = u16(head + 18);
    const bbox = [i16(head + 36), i16(head + 38), i16(head + 40), i16(head + 42)];
    const hhea = req('hhea');
    const ascenso = i16(hhea + 4), descenso = i16(hhea + 6);
    const nMetricas = u16(hhea + 34);
    const numGlifos = u16(req('maxp') + 4);

    const hmtx = req('hmtx');
    const anchos = new Uint16Array(numGlifos);
    let ultimo = 0;
    for (let g = 0; g < numGlifos; g++) {
      if (g < nMetricas) ultimo = u16(hmtx + g * 4);
      anchos[g] = ultimo;
    }

    let altoMayus = Math.round(ascenso * 0.7), peso = 400;
    if (tablas['OS/2']) {
      const os2 = tablas['OS/2'].off;
      peso = u16(os2 + 4);
      if (u16(os2) >= 2 && tablas['OS/2'].len >= 90) altoMayus = i16(os2 + 88);
    }
    let anguloItalica = 0;
    if (tablas.post) anguloItalica = dv.getInt32(tablas.post.off + 4) / 65536;

    // cmap: preferir (3,10) formato 12, luego (3,1) formato 4
    const cmap = new Map();
    const cm = req('cmap');
    const n = u16(cm + 2);
    let sub12 = null, sub4 = null;
    for (let i = 0; i < n; i++) {
      const p = u16(cm + 4 + i * 8), e = u16(cm + 6 + i * 8), off = cm + u32(cm + 8 + i * 8);
      const fmt = u16(off);
      if (fmt === 12 && p === 3 && e === 10) sub12 = off;
      if (fmt === 4 && p === 3 && (e === 1 || e === 0)) sub4 = off;
    }
    if (sub12 != null) {
      const grupos = u32(sub12 + 12);
      for (let i = 0; i < grupos; i++) {
        const o = sub12 + 16 + i * 12;
        const ini = u32(o), fin = u32(o + 4), g0 = u32(o + 8);
        for (let c = ini; c <= fin; c++) cmap.set(c, g0 + (c - ini));
      }
    } else if (sub4 != null) {
      const segX2 = u16(sub4 + 6), seg = segX2 / 2;
      const finOff = sub4 + 14, iniOff = finOff + segX2 + 2, deltaOff = iniOff + segX2, rangoOff = deltaOff + segX2;
      for (let s = 0; s < seg; s++) {
        const fin = u16(finOff + s * 2), ini = u16(iniOff + s * 2);
        const delta = i16(deltaOff + s * 2), rango = u16(rangoOff + s * 2);
        for (let c = ini; c <= fin && c !== 0xFFFF; c++) {
          let g;
          if (rango === 0) g = (c + delta) & 0xFFFF;
          else {
            const dir = rangoOff + s * 2 + rango + (c - ini) * 2;
            g = u16(dir);
            if (g !== 0) g = (g + delta) & 0xFFFF;
          }
          if (g) cmap.set(c, g);
        }
      }
    }
    return { unidades, bbox, ascenso, descenso, altoMayus, peso, anguloItalica, numGlifos, anchos, cmap };
  }

  // ------------------------------------------------------------------ incrustación en pdf-lib

  function etiquetaSubconjunto(clave) {
    let h = 0;
    for (const ch of clave) h = (h * 31 + ch.charCodeAt(0)) >>> 0;
    let s = '';
    for (let i = 0; i < 6; i++) { s += String.fromCharCode(65 + (h % 26)); h = Math.floor(h / 26) + 7; }
    return s;
  }

  function cmapToUnicode(inverso) {
    const lineas = [
      '/CIDInit /ProcSet findresource begin', '12 dict begin', 'begincmap',
      '/CIDSystemInfo << /Registry (Adobe) /Ordering (UCS) /Supplement 0 >> def',
      '/CMapName /Adobe-Identity-UCS def', '/CMapType 2 def',
      '1 begincodespacerange', '<0000> <FFFF>', 'endcodespacerange'
    ];
    const pares = Array.from(inverso.entries()).sort((a, b) => a[0] - b[0]);
    for (let i = 0; i < pares.length; i += 100) {
      const bloque = pares.slice(i, i + 100);
      lineas.push(bloque.length + ' beginbfchar');
      bloque.forEach(([g, cp]) => {
        let u;
        if (cp > 0xFFFF) {
          const v = cp - 0x10000;
          u = ((0xD800 + (v >> 10)).toString(16).padStart(4, '0') + (0xDC00 + (v & 0x3FF)).toString(16).padStart(4, '0'));
        } else u = cp.toString(16).padStart(4, '0');
        lineas.push('<' + g.toString(16).padStart(4, '0') + '> <' + u + '>');
      });
      lineas.push('endbfchar');
    }
    lineas.push('endcmap', 'CMapName currentdict /CMap defineresource pop', 'end', 'end');
    return lineas.join('\n');
  }

  /** Gestor de fuentes para un documento pdf-lib. */
  class GestorFuentes {
    constructor(doc, bytesPorCara) {
      this.doc = doc;
      this.bytes = bytesPorCara;      // { 'carlito-r': Uint8Array, ... }
      this.info = new Map();          // cara -> datos TTF
      this.incrustadas = new Map();   // cara -> { ref, usados:Set }
    }

    datos(cara) {
      if (!this.info.has(cara)) {
        const b = this.bytes[cara];
        if (!b) throw new Error('Fuente no disponible: ' + cara);
        this.info.set(cara, leerTTF(b));
      }
      return this.info.get(cara);
    }

    tieneGlifo(cara, cp) {
      return this.datos(cara).cmap.has(cp);
    }

    /** Devuelve la referencia de fuente (la crea al primer uso). La ToUnicode se escribe al final. */
    referencia(cara) {
      let e = this.incrustadas.get(cara);
      if (e) return e;
      const d = this.datos(cara);
      const ctx = this.doc.context;
      const escala = 1000 / d.unidades;
      const nombre = etiquetaSubconjunto(cara) + '+' + NOMBRE_PS[cara.split('-')[0]] +
        { r: '', b: '-Bold', i: '-Italic', bi: '-BoldItalic' }[cara.split('-')[1]];
      const italica = /i$/.test(cara);

      const archivo = ctx.flateStream(this.bytes[cara], { Length1: this.bytes[cara].length });
      const archivoRef = ctx.register(archivo);
      const descriptor = ctx.obj({
        Type: 'FontDescriptor',
        FontName: nombre,
        Flags: 32 + (italica ? 64 : 0) + (cara.startsWith('libmono') ? 1 : 0),
        FontBBox: d.bbox.map(v => Math.round(v * escala)),
        ItalicAngle: d.anguloItalica || (italica ? -12 : 0),
        Ascent: Math.round(d.ascenso * escala),
        Descent: Math.round(d.descenso * escala),
        CapHeight: Math.round(d.altoMayus * escala),
        StemV: /b/.test(cara.split('-')[1]) ? 120 : 80,
        FontFile2: archivoRef
      });
      const descRef = ctx.register(descriptor);
      const anchos = Array.from(d.anchos, w => Math.round(w * escala));
      const cid = ctx.obj({
        Type: 'Font', Subtype: 'CIDFontType2', BaseFont: nombre,
        CIDSystemInfo: { Registry: PDFString.of('Adobe'), Ordering: PDFString.of('Identity'), Supplement: 0 },
        FontDescriptor: descRef, DW: 1000, W: [0, anchos], CIDToGIDMap: 'Identity'
      });
      const cidRef = ctx.register(cid);
      const ref = ctx.nextRef();
      e = { ref, cidRef, nombre, usados: new Map() };
      this.incrustadas.set(cara, e);
      return e;
    }

    /** Convierte texto en la cadena hexadecimal de glifos y registra el mapeo a Unicode. */
    codificar(cara, texto) {
      const d = this.datos(cara);
      const e = this.referencia(cara);
      let hex = '';
      for (const ch of texto) {
        const cp = ch.codePointAt(0);
        const g = d.cmap.get(cp) || 0;
        if (g && !e.usados.has(g)) e.usados.set(g, cp);
        hex += g.toString(16).padStart(4, '0');
      }
      return hex;
    }

    anchoTexto(cara, texto, tam) {
      const d = this.datos(cara);
      let w = 0;
      for (const ch of texto) w += d.anchos[d.cmap.get(ch.codePointAt(0)) || 0];
      return w * tam / d.unidades;
    }

    /** Cierra las fuentes: escribe los objetos Type0 con su ToUnicode. Llamar antes de doc.save(). */
    finalizar() {
      const ctx = this.doc.context;
      for (const [, e] of this.incrustadas) {
        const tu = ctx.flateStream(cmapToUnicode(e.usados));
        const tuRef = ctx.register(tu);
        ctx.assign(e.ref, ctx.obj({
          Type: 'Font', Subtype: 'Type0', BaseFont: e.nombre, Encoding: 'Identity-H',
          DescendantFonts: [e.cidRef], ToUnicode: tuRef
        }));
      }
    }
  }

  /** Dibuja texto con una fuente incrustada (operadores de bajo nivel). */
  function dibujarTexto(pagina, gestor, cara, texto, x, y, tam, color, nombresPorPagina) {
    if (!texto) return;
    const e = gestor.referencia(cara);
    let porPagina = nombresPorPagina.get(pagina);
    if (!porPagina) { porPagina = new Map(); nombresPorPagina.set(pagina, porPagina); }
    let nombre = porPagina.get(cara);
    if (!nombre) {
      nombre = pagina.node.newFontDictionary('CPF', e.ref);
      porPagina.set(cara, nombre);
    }
    const hex = gestor.codificar(cara, texto);
    const {
      pushGraphicsState, popGraphicsState, beginText, endText, setFontAndSize, setTextMatrix,
      showText, setFillingRgbColor
    } = PDFLib;
    pagina.pushOperators(
      pushGraphicsState(),
      setFillingRgbColor(color[0], color[1], color[2]),
      beginText(),
      setFontAndSize(nombre, tam),
      setTextMatrix(1, 0, 0, 1, x, y),
      showText(PDFHexString.of(hex)),
      endText(),
      popGraphicsState()
    );
  }

  return { FAMILIAS, ESTILOS, leerTTF, GestorFuentes, dibujarTexto };
});
