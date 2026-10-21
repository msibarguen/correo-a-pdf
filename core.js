/*!
 * Correo a PDF — motor de armado (core)
 * Prototipo v0.1.0 — Unidad de Transformación Digital, Rama Judicial
 *
 * Arma un único PDF con:
 *   1. Carátula (datos del correo, índice, estado de cada anexo, hashes SHA-256)
 *   2. Cuerpo del correo (texto) y enlaces encontrados (clicables)
 *   3. Anexos PDF e imágenes, unidos en orden
 * Los archivos que no son PDF ni imagen NO se convierten: se alertan y se
 * registran en la carátula.
 *
 * Funciona en el navegador (complemento de Outlook) y en Node (pruebas).
 * No envía datos a ningún servidor: todo ocurre en memoria, en el equipo.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) {
    module.exports = factory(require('pdf-lib'), require('crypto').webcrypto);
  } else {
    root.CorreoPDF = factory(root.PDFLib, root.crypto);
  }
})(typeof self !== 'undefined' ? self : this, function (PDFLib, webcrypto) {
  'use strict';

  const {
    PDFDocument, StandardFonts, rgb, PDFName, PDFString, PDFHexString, PDFNull
  } = PDFLib;

  const VERSION = '0.1.0';
  const LIMITE_BYTES = 100 * 1024 * 1024; // 100 MB por correo
  const CARTA = [612, 792];
  const MARGEN = 50;
  const ZONA_HORARIA = 'America/Bogota';

  const COLOR = {
    texto: rgb(0.12, 0.12, 0.14),
    suave: rgb(0.40, 0.40, 0.44),
    linea: rgb(0.78, 0.78, 0.80),
    alerta: rgb(0.70, 0.12, 0.10),
    ok: rgb(0.08, 0.45, 0.25),
    enlace: rgb(0.05, 0.30, 0.65),
    fondoAlerta: rgb(1, 0.95, 0.93)
  };

  // ---------------------------------------------------------------------------
  // Utilidades
  // ---------------------------------------------------------------------------

  async function sha256(bytes) {
    const buf = await webcrypto.subtle.digest('SHA-256', bytes);
    return Array.from(new Uint8Array(buf)).map(b => b.toString(16).padStart(2, '0')).join('');
  }

  function tamanoLegible(n) {
    if (n == null || isNaN(n)) return '—';
    if (n < 1024) return n + ' B';
    if (n < 1024 * 1024) return (n / 1024).toFixed(1) + ' KB';
    return (n / (1024 * 1024)).toFixed(2) + ' MB';
  }

  function extension(nombre) {
    const m = /\.([^.]+)$/.exec(nombre || '');
    return m ? m[1].toLowerCase() : '';
  }

  function fechaLegible(fecha, zona) {
    if (!fecha) return '—';
    const d = fecha instanceof Date ? fecha : new Date(fecha);
    if (isNaN(d)) return String(fecha);
    try {
      return new Intl.DateTimeFormat('es-CO', {
        timeZone: zona, weekday: 'long', year: 'numeric', month: 'long', day: 'numeric',
        hour: '2-digit', minute: '2-digit', second: '2-digit', hour12: true
      }).format(d) + ' (hora de Colombia)';
    } catch (e) {
      return d.toISOString();
    }
  }

  function partesFecha(fecha, zona) {
    const d = fecha instanceof Date ? fecha : new Date(fecha);
    const f = new Intl.DateTimeFormat('en-CA', {
      timeZone: zona, year: 'numeric', month: '2-digit', day: '2-digit',
      hour: '2-digit', minute: '2-digit', hour12: false
    }).formatToParts(isNaN(d) ? new Date() : d);
    const p = {};
    f.forEach(x => { p[x.type] = x.value; });
    return p;
  }

  function slug(texto, max) {
    const s = (texto || '')
      .normalize('NFD').replace(/[̀-ͯ]/g, '')
      .replace(/[^A-Za-z0-9]+/g, '-').replace(/^-+|-+$/g, '')
      .slice(0, max || 40).replace(/-+$/g, '');
    return s || 'sin-asunto';
  }

  // Número de radicado de 23 dígitos, con o sin separadores.
  function detectarRadicados(...textos) {
    const encontrados = new Set();
    const re = /(?<![\d])(\d[\d .\-]{21,45}\d)(?![\d])/g;
    textos.filter(Boolean).forEach(t => {
      let m;
      while ((m = re.exec(t)) !== null) {
        const digitos = m[1].replace(/\D/g, '');
        if (digitos.length === 23) encontrados.add(digitos);
      }
    });
    return Array.from(encontrados).slice(0, 5);
  }

  function formatoRadicado(d) {
    // 5-2-2-3-4-5-2 (departamento/municipio, entidad, especialidad, despacho, año, consecutivo, recurso)
    return [d.slice(0, 5), d.slice(5, 7), d.slice(7, 9), d.slice(9, 12),
      d.slice(12, 16), d.slice(16, 21), d.slice(21, 23)].join('-');
  }

  // Desenvuelve enlaces protegidos por Microsoft Defender (Safe Links).
  function desenvolverEnlace(url) {
    try {
      const u = new URL(url);
      if (/safelinks\.protection\.outlook\.com$/i.test(u.hostname) && u.searchParams.get('url')) {
        return { url: u.searchParams.get('url'), safelinks: true };
      }
    } catch (e) { /* no es URL válida */ }
    return { url, safelinks: false };
  }

  function normalizarEnlaces(lista) {
    const vistos = new Set();
    const salida = [];
    (lista || []).forEach(item => {
      const entrada = typeof item === 'string' ? { url: item } : item;
      if (!entrada || !entrada.url) return;
      const limpio = entrada.url.trim().replace(/[).,;:!?>\]]+$/, '');
      if (!/^https?:\/\//i.test(limpio)) return;
      const d = desenvolverEnlace(limpio);
      if (vistos.has(d.url)) return;
      vistos.add(d.url);
      salida.push({ url: d.url, etiqueta: entrada.etiqueta || '', safelinks: d.safelinks });
    });
    return salida;
  }

  /** Extrae enlaces de un texto plano. */
  function enlacesDeTexto(texto) {
    const re = /https?:\/\/[^\s<>"'\]]+/gi;
    return (texto || '').match(re) || [];
  }

  // ---------------------------------------------------------------------------
  // Detección del tipo real del archivo (por contenido, no solo por extensión)
  // ---------------------------------------------------------------------------

  const NOMBRES_TIPO = {
    doc: 'Documento de Word', docx: 'Documento de Word', docm: 'Documento de Word',
    odt: 'Documento de texto (OpenDocument)', rtf: 'Documento RTF', txt: 'Archivo de texto',
    xls: 'Hoja de Excel', xlsx: 'Hoja de Excel', xlsm: 'Hoja de Excel', csv: 'Archivo CSV',
    ods: 'Hoja de cálculo (OpenDocument)',
    ppt: 'Presentación de PowerPoint', pptx: 'Presentación de PowerPoint',
    zip: 'Archivo comprimido', rar: 'Archivo comprimido', '7z': 'Archivo comprimido',
    mp3: 'Audio', wav: 'Audio', ogg: 'Audio', m4a: 'Audio', opus: 'Audio', amr: 'Audio',
    mp4: 'Video', avi: 'Video', mov: 'Video', mkv: 'Video', wmv: 'Video',
    msg: 'Correo adjunto', eml: 'Correo adjunto',
    tif: 'Imagen TIFF', tiff: 'Imagen TIFF', heic: 'Imagen HEIC (iPhone)', heif: 'Imagen HEIF',
    xml: 'Archivo XML', html: 'Página web', htm: 'Página web', p7z: 'Archivo firmado', p7m: 'Archivo firmado'
  };

  function detectarTipo(bytes, nombre) {
    const b = bytes;
    const ext = extension(nombre);
    const empieza = (...sig) => sig.every((v, i) => b[i] === v);

    // %PDF puede aparecer dentro del primer KB
    const limite = Math.min(b.length - 4, 1024);
    for (let i = 0; i <= limite; i++) {
      if (b[i] === 0x25 && b[i + 1] === 0x50 && b[i + 2] === 0x44 && b[i + 3] === 0x46) {
        return { categoria: 'pdf', tipo: 'pdf', descripcion: 'PDF' };
      }
    }
    if (empieza(0x89, 0x50, 0x4E, 0x47)) return { categoria: 'imagen', tipo: 'png', descripcion: 'Imagen PNG' };
    if (empieza(0xFF, 0xD8, 0xFF)) return { categoria: 'imagen', tipo: 'jpg', descripcion: 'Imagen JPG' };
    if (empieza(0x47, 0x49, 0x46, 0x38)) return { categoria: 'imagen-convertible', tipo: 'gif', descripcion: 'Imagen GIF' };
    if (empieza(0x42, 0x4D)) return { categoria: 'imagen-convertible', tipo: 'bmp', descripcion: 'Imagen BMP' };
    if (empieza(0x52, 0x49, 0x46, 0x46) && b[8] === 0x57 && b[9] === 0x45 && b[10] === 0x42 && b[11] === 0x50) {
      return { categoria: 'imagen-convertible', tipo: 'webp', descripcion: 'Imagen WEBP' };
    }

    let descripcion = NOMBRES_TIPO[ext] || (ext ? 'Archivo .' + ext : 'Archivo sin extensión');
    let motivo = 'Formato no soportado: solo se unen PDF e imágenes (JPG, PNG, GIF, BMP, WEBP).';
    if (ext === 'pdf') {
      descripcion = 'Archivo con extensión .pdf';
      motivo = 'Tiene extensión .pdf pero su contenido no es un PDF (posible archivo renombrado o dañado).';
    } else if (['jpg', 'jpeg', 'png'].includes(ext)) {
      motivo = 'Tiene extensión de imagen pero su contenido no corresponde (posible archivo dañado).';
    }
    return { categoria: 'no-soportado', tipo: ext || 'desconocido', descripcion, motivo };
  }

  /** Clasificación rápida solo por nombre, para mostrar alertas antes de generar. */
  function prediccionPorNombre(nombre) {
    const ext = extension(nombre);
    if (ext === 'pdf') return { soportado: true, descripcion: 'PDF' };
    if (['jpg', 'jpeg', 'png', 'gif', 'bmp', 'webp'].includes(ext)) return { soportado: true, descripcion: 'Imagen' };
    return { soportado: false, descripcion: NOMBRES_TIPO[ext] || (ext ? 'Archivo .' + ext : 'Archivo') };
  }

  // Orientación EXIF de un JPEG (1 = normal). Devuelve 1 si no la encuentra.
  function orientacionExif(b) {
    try {
      if (b[0] !== 0xFF || b[1] !== 0xD8) return 1;
      let i = 2;
      while (i < b.length - 4) {
        if (b[i] !== 0xFF) return 1;
        const marcador = b[i + 1];
        const largo = (b[i + 2] << 8) | b[i + 3];
        if (marcador === 0xE1 && b[i + 4] === 0x45 && b[i + 5] === 0x78 && b[i + 6] === 0x69 && b[i + 7] === 0x66) {
          const t = i + 10; // inicio TIFF
          const little = b[t] === 0x49;
          const u16 = o => little ? (b[o] | (b[o + 1] << 8)) : ((b[o] << 8) | b[o + 1]);
          const u32 = o => little
            ? (b[o] | (b[o + 1] << 8) | (b[o + 2] << 16) | (b[o + 3] << 24)) >>> 0
            : ((b[o] << 24) | (b[o + 1] << 16) | (b[o + 2] << 8) | b[o + 3]) >>> 0;
          const ifd = t + u32(t + 4);
          const n = u16(ifd);
          for (let k = 0; k < n; k++) {
            const e = ifd + 2 + k * 12;
            if (u16(e) === 0x0112) return u16(e + 8) || 1;
          }
          return 1;
        }
        if (marcador === 0xDA) return 1;
        i += 2 + largo;
      }
    } catch (e) { /* ignorar */ }
    return 1;
  }

  // ---------------------------------------------------------------------------
  // Escritor de texto con saltos de línea y de página
  // ---------------------------------------------------------------------------

  class Escritor {
    constructor(doc, fuentes) {
      this.doc = doc;
      this.f = fuentes;
      this.anchoUtil = CARTA[0] - 2 * MARGEN;
      this.paginas = [];
      this.cacheCaracteres = new Map();
      this.nuevaPagina();
    }

    nuevaPagina() {
      this.pagina = this.doc.addPage(CARTA);
      this.paginas.push(this.pagina);
      this.y = CARTA[1] - MARGEN;
    }

    asegurar(alto) {
      if (this.y - alto < MARGEN + 20) this.nuevaPagina();
    }

    // Reemplaza caracteres que la fuente estándar (WinAnsi) no puede dibujar.
    limpiar(texto, fuente) {
      const sustitutos = { '→': '->', '←': '<-', '✓': 'v', '✔': 'v', ' ': ' ', '​': '', '﻿': '', '\t': '    ' };
      let salida = '';
      for (const ch of Array.from(String(texto == null ? '' : texto))) {
        if (ch === '\n') { salida += ch; continue; }
        if (sustitutos[ch] !== undefined) { salida += sustitutos[ch]; continue; }
        const clave = fuente.name + ch;
        let ok = this.cacheCaracteres.get(clave);
        if (ok === undefined) {
          try { fuente.encodeText(ch); ok = true; } catch (e) { ok = false; }
          this.cacheCaracteres.set(clave, ok);
        }
        salida += ok ? ch : (ch.codePointAt(0) < 32 ? '' : '?');
      }
      return salida;
    }

    partir(texto, fuente, tam, ancho) {
      const lineas = [];
      const parrafos = this.limpiar(texto, fuente).replace(/\r\n?/g, '\n').split('\n');
      for (const p of parrafos) {
        if (p.trim() === '') { lineas.push(''); continue; }
        let actual = '';
        for (const palabra of p.split(/ +/)) {
          const prueba = actual ? actual + ' ' + palabra : palabra;
          if (fuente.widthOfTextAtSize(prueba, tam) <= ancho) { actual = prueba; continue; }
          if (actual) lineas.push(actual);
          // palabra más larga que la línea (p. ej. URL): cortar por caracteres
          if (fuente.widthOfTextAtSize(palabra, tam) > ancho) {
            let trozo = '';
            for (const ch of palabra) {
              if (fuente.widthOfTextAtSize(trozo + ch, tam) > ancho) { lineas.push(trozo); trozo = ch; } else trozo += ch;
            }
            actual = trozo;
          } else {
            actual = palabra;
          }
        }
        lineas.push(actual);
      }
      return lineas;
    }

    texto(texto, o = {}) {
      const fuente = o.negrita ? this.f.negrita : (o.mono ? this.f.mono : this.f.normal);
      const tam = o.tam || 10;
      const sangria = o.sangria || 0;
      const alto = tam * 1.35;
      const lineas = this.partir(texto, fuente, tam, this.anchoUtil - sangria);
      const cajas = [];
      for (const l of lineas) {
        this.asegurar(alto);
        this.y -= alto;
        if (l) {
          this.pagina.drawText(l, { x: MARGEN + sangria, y: this.y, size: tam, font: fuente, color: o.color || COLOR.texto });
          cajas.push({ pagina: this.pagina, x: MARGEN + sangria, y: this.y - 2, w: fuente.widthOfTextAtSize(l, tam), h: tam + 3 });
        }
      }
      return cajas;
    }

    // Etiqueta en negrita a la izquierda y valor ajustado a la derecha
    campo(etiqueta, valor, anchoEtiqueta = 120) {
      const tam = 9.5;
      const alto = tam * 1.35;
      const lineas = this.partir(valor || '—', this.f.normal, tam, this.anchoUtil - anchoEtiqueta);
      this.asegurar(alto);
      lineas.forEach((l, i) => {
        if (i > 0) this.asegurar(alto);
        this.y -= alto;
        if (i === 0) {
          this.pagina.drawText(this.limpiar(etiqueta, this.f.negrita), { x: MARGEN, y: this.y, size: tam, font: this.f.negrita, color: COLOR.suave });
        }
        this.pagina.drawText(l, { x: MARGEN + anchoEtiqueta, y: this.y, size: tam, font: this.f.normal, color: COLOR.texto });
      });
    }

    espacio(h) { this.y -= h; }

    linea() {
      this.asegurar(10);
      this.y -= 6;
      this.pagina.drawLine({ start: { x: MARGEN, y: this.y }, end: { x: CARTA[0] - MARGEN, y: this.y }, thickness: 0.6, color: COLOR.linea });
      this.y -= 6;
    }

    titulo(texto) {
      this.asegurar(40);
      this.espacio(8);
      this.texto(texto, { negrita: true, tam: 11.5 });
      this.linea();
    }

    recuadroAlerta(lineasTexto) {
      const tam = 9.5;
      const alto = tam * 1.4;
      const todas = [];
      lineasTexto.forEach((t, i) => this.partir(t, i === 0 ? this.f.negrita : this.f.normal, tam, this.anchoUtil - 20)
        .forEach(l => todas.push({ l, negrita: i === 0 })));
      const altoCaja = todas.length * alto + 14;
      this.asegurar(altoCaja + 6);
      const yTope = this.y - 4;
      this.pagina.drawRectangle({
        x: MARGEN, y: yTope - altoCaja, width: this.anchoUtil, height: altoCaja,
        color: COLOR.fondoAlerta, borderColor: COLOR.alerta, borderWidth: 0.8
      });
      let y = yTope - 7;
      todas.forEach(({ l, negrita }) => {
        y -= alto;
        this.pagina.drawText(l, { x: MARGEN + 10, y: y + 3, size: tam, font: negrita ? this.f.negrita : this.f.normal, color: COLOR.alerta });
      });
      this.y = yTope - altoCaja - 8;
    }
  }

  function agregarEnlace(doc, caja, url) {
    const anotacion = doc.context.obj({
      Type: 'Annot', Subtype: 'Link',
      Rect: [caja.x, caja.y, caja.x + caja.w, caja.y + caja.h],
      Border: [0, 0, 0],
      A: { Type: 'Action', S: 'URI', URI: PDFString.of(url) }
    });
    caja.pagina.node.addAnnot(doc.context.register(anotacion));
  }

  function agregarMarcadores(doc, entradas) {
    if (!entradas.length) return;
    const ctx = doc.context;
    const raiz = ctx.nextRef();
    const refs = entradas.map(() => ctx.nextRef());
    entradas.forEach((e, i) => {
      const pagina = doc.getPage(e.indicePagina);
      const dict = {
        Title: PDFHexString.fromText(e.titulo),
        Parent: raiz,
        Dest: [pagina.ref, PDFName.of('XYZ'), PDFNull, PDFNull, PDFNull]
      };
      if (i > 0) dict.Prev = refs[i - 1];
      if (i < refs.length - 1) dict.Next = refs[i + 1];
      ctx.assign(refs[i], ctx.obj(dict));
    });
    ctx.assign(raiz, ctx.obj({ Type: 'Outlines', First: refs[0], Last: refs[refs.length - 1], Count: refs.length }));
    doc.catalog.set(PDFName.of('Outlines'), raiz);
    doc.catalog.set(PDFName.of('PageMode'), PDFName.of('UseOutlines'));
  }

  // ---------------------------------------------------------------------------
  // Preparación de anexos
  // ---------------------------------------------------------------------------

  async function prepararAnexos(anexos, opciones, avisar) {
    const resultado = [];
    for (let i = 0; i < anexos.length; i++) {
      const a = anexos[i];
      avisar(`Revisando anexo ${i + 1} de ${anexos.length}: ${a.nombre}`);
      const r = {
        nombre: a.nombre || ('anexo-' + (i + 1)),
        tamano: a.tamano != null ? a.tamano : (a.bytes ? a.bytes.length : null),
        descripcion: '', estado: '', motivo: '', sha256: '', paginas: 0, recurso: null
      };

      if (a.tipo === 'nube') {
        r.estado = 'enlace';
        r.descripcion = 'Archivo compartido en la nube';
        r.motivo = 'Se incluye como enlace; su contenido no fue descargado.';
        r.url = a.url || '';
        r.tamano = a.tamano ? a.tamano : null;
        resultado.push(r);
        continue;
      }
      if (a.tipo === 'item') {
        r.estado = 'no_convertido';
        r.descripcion = 'Correo o elemento adjunto';
        r.motivo = 'Los correos adjuntos no se convierten en esta versión. Conserve el original.';
        resultado.push(r);
        continue;
      }
      if (a.error || !a.bytes) {
        r.estado = 'no_convertido';
        r.descripcion = prediccionPorNombre(r.nombre).descripcion;
        r.motivo = 'No se pudo leer el archivo desde Outlook' + (a.error ? ': ' + a.error : '.');
        resultado.push(r);
        continue;
      }

      r.sha256 = await sha256(a.bytes);
      const tipo = detectarTipo(a.bytes, r.nombre);
      r.descripcion = tipo.descripcion;

      if (a.enLinea && !opciones.incluirImagenesEnLinea) {
        r.estado = 'omitido';
        r.motivo = 'Imagen incrustada en el cuerpo del correo (firma, logo). No se incluyó.';
        resultado.push(r);
        continue;
      }

      try {
        if (tipo.categoria === 'pdf') {
          const pdf = await PDFDocument.load(a.bytes, { updateMetadata: false });
          r.paginas = pdf.getPageCount();
          if (r.paginas === 0) throw new Error('El PDF no tiene páginas.');
          r.recurso = { clase: 'pdf', pdf };
          r.estado = 'incluido';
        } else if (tipo.categoria === 'imagen' || tipo.categoria === 'imagen-convertible') {
          let bytesImg = a.bytes;
          let formato = tipo.tipo;
          const orientacion = formato === 'jpg' ? orientacionExif(a.bytes) : 1;
          if (tipo.categoria === 'imagen-convertible' || orientacion > 1) {
            if (!opciones.convertirImagen) {
              if (tipo.categoria === 'imagen-convertible') throw new Error('Conversión de ' + formato.toUpperCase() + ' no disponible.');
            } else {
              const destino = formato === 'jpg' ? 'jpg' : 'png';
              bytesImg = await opciones.convertirImagen(a.bytes, formato, destino);
              formato = destino;
            }
          }
          r.recurso = { clase: 'imagen', bytes: bytesImg, formato };
          r.paginas = 1;
          r.estado = 'incluido';
        } else {
          r.estado = 'no_convertido';
          r.motivo = tipo.motivo;
        }
      } catch (e) {
        r.estado = 'no_convertido';
        const msg = String(e && e.message || e);
        if (/encrypt/i.test(msg) || (e && e.constructor && /Encrypted/.test(e.constructor.name))) {
          r.motivo = 'PDF protegido o cifrado: no se puede unir. Conserve el original o solicite una copia sin protección.';
        } else if (tipo.categoria === 'pdf') {
          r.motivo = 'PDF dañado o no legible: ' + msg;
        } else {
          r.motivo = 'No se pudo procesar la imagen: ' + msg;
        }
        r.recurso = null;
      }
      resultado.push(r);
    }
    return resultado;
  }

  // ---------------------------------------------------------------------------
  // Carátula + cuerpo + enlaces
  // ---------------------------------------------------------------------------

  async function construirFrente(datos, anexos, enlaces, radicados, rangos, totalPaginas, opciones) {
    const doc = await PDFDocument.create();
    const fuentes = {
      normal: await doc.embedFont(StandardFonts.Helvetica),
      negrita: await doc.embedFont(StandardFonts.HelveticaBold),
      mono: await doc.embedFont(StandardFonts.Courier)
    };
    const e = new Escritor(doc, fuentes);
    const m = datos.meta || {};
    const marcas = [];

    // Encabezado
    e.texto('RAMA JUDICIAL DEL PODER PÚBLICO', { negrita: true, tam: 8.5, color: COLOR.suave });
    e.texto(opciones.nombreDespacho || 'Secretaría', { tam: 8.5, color: COLOR.suave });
    e.espacio(6);
    e.texto('CORREO ELECTRÓNICO RECIBIDO Y SUS ANEXOS', { negrita: true, tam: 15 });
    e.linea();
    marcas.push({ titulo: 'Carátula', pagina: e.paginas.length - 1 });

    e.campo('Asunto', m.asunto || '(sin asunto)');
    e.campo('De', m.de || '—');
    e.campo('Para', (m.para || []).join('; ') || '—');
    if (m.cc && m.cc.length) e.campo('CC', m.cc.join('; '));
    e.campo('Recibido', fechaLegible(m.recibido, opciones.zonaHoraria));
    if (m.idMensaje) e.campo('ID del mensaje', m.idMensaje);
    e.campo('Radicado detectado', radicados.length
      ? radicados.map(formatoRadicado).join('  |  ') + '  (verificar)'
      : 'No se detectó un número de radicado de 23 dígitos.');

    const noConvertidos = anexos.filter(a => a.estado === 'no_convertido');
    if (noConvertidos.length) {
      e.espacio(4);
      e.recuadroAlerta([
        `ATENCIÓN: ${noConvertidos.length} archivo(s) recibido(s) NO se convirtieron y NO están incluidos en este PDF.`,
        'Revise la lista de anexos abajo. Conserve el correo original y gestione esos archivos por separado.'
      ]);
    }

    // Índice
    e.titulo('ÍNDICE DEL DOCUMENTO');
    e.texto(`Carátula y cuerpo del correo .......... páginas 1 a ${rangos.finFrente}`, { tam: 9.5 });
    anexos.forEach((a, i) => {
      if (a.estado === 'incluido') {
        const r = rangos.porAnexo[i];
        e.texto(`Anexo ${i + 1}: ${a.nombre} .......... ${r[0] === r[1] ? 'página ' + r[0] : 'páginas ' + r[0] + ' a ' + r[1]}`, { tam: 9.5 });
      }
    });
    e.texto(`Total de páginas: ${totalPaginas}`, { tam: 9.5, negrita: true });

    // Detalle de anexos
    e.titulo(`ARCHIVOS RECIBIDOS (${anexos.length})`);
    e.texto('El SHA-256 es la huella digital de cada archivo tal como llegó en el correo; permite verificar que un archivo es el mismo recibido.', { tam: 7.5, color: COLOR.suave });
    e.espacio(4);
    if (!anexos.length) e.texto('El correo no traía archivos adjuntos.', { tam: 9.5, color: COLOR.suave });
    anexos.forEach((a, i) => {
      e.asegurar(48);
      e.texto(`${i + 1}. ${a.nombre}`, { negrita: true, tam: 9.5 });
      e.texto(`Tipo: ${a.descripcion || '—'}   ·   Tamaño: ${tamanoLegible(a.tamano)}`, { tam: 8.5, sangria: 12, color: COLOR.suave });
      let estado, color;
      if (a.estado === 'incluido') {
        const r = rangos.porAnexo[i];
        estado = `Incluido en este PDF (${r[0] === r[1] ? 'página ' + r[0] : 'páginas ' + r[0] + ' a ' + r[1]})`;
        color = COLOR.ok;
      } else if (a.estado === 'enlace') {
        estado = 'Archivo en la nube: incluido como enlace (ver sección de enlaces)';
        color = COLOR.enlace;
      } else if (a.estado === 'omitido') {
        estado = 'Omitido: ' + a.motivo;
        color = COLOR.suave;
      } else {
        estado = 'NO CONVERTIDO: ' + a.motivo;
        color = COLOR.alerta;
      }
      e.texto(estado, { tam: 8.5, sangria: 12, color, negrita: a.estado === 'no_convertido' });
      if (a.sha256) e.texto('SHA-256: ' + a.sha256, { tam: 7, sangria: 12, mono: true, color: COLOR.suave });
      e.espacio(3);
    });


    // Cuerpo
    e.nuevaPagina();
    marcas.push({ titulo: 'Cuerpo del correo', pagina: e.paginas.length - 1 });
    e.titulo('CUERPO DEL CORREO');
    const cuerpo = (datos.cuerpoTexto || '').replace(/\n{3,}/g, '\n\n').trim();
    e.texto(cuerpo || '(El correo no tiene texto en el cuerpo.)', { tam: 10 });

    // Enlaces
    e.espacio(10);
    e.titulo(`ENLACES ENCONTRADOS EN EL CORREO (${enlaces.length})`);
    marcas.push({ titulo: 'Enlaces del correo', pagina: e.paginas.length - 1 });
    if (!enlaces.length) {
      e.texto('El correo no contiene enlaces.', { tam: 9.5, color: COLOR.suave });
    } else {
      e.texto('Los enlaces NO fueron abiertos ni descargados; su contenido no está incluido en este PDF. Haga clic para abrirlos.', { tam: 8.5, color: COLOR.suave });
      e.espacio(4);
      enlaces.forEach((l, i) => {
        e.asegurar(30);
        if (l.etiqueta) e.texto(`${i + 1}. ${l.etiqueta}`, { tam: 9, negrita: true });
        const cajas = e.texto((l.etiqueta ? '   ' : `${i + 1}. `) + l.url, { tam: 9, color: COLOR.enlace });
        cajas.forEach(c => agregarEnlace(doc, c, l.url));
        if (l.safelinks) e.texto('(Enlace original, extraído de la protección Safe Links de Microsoft)', { tam: 7.5, sangria: 12, color: COLOR.suave });
        e.espacio(2);
      });
    }

    // Pie de página en las páginas generadas
    const generado = fechaLegible(opciones.ahora || new Date(), opciones.zonaHoraria);
    e.paginas.forEach((p, i) => {
      const pie = `Generado con Correo a PDF v${VERSION} (prototipo) · ${generado} · Página ${i + 1} de ${totalPaginas}`;
      p.drawText(e.limpiar(pie, fuentes.normal), { x: MARGEN, y: 28, size: 7, font: fuentes.normal, color: COLOR.suave });
    });

    return { doc, paginas: e.paginas.length, marcas };
  }

  // ---------------------------------------------------------------------------
  // Función principal
  // ---------------------------------------------------------------------------

  /**
   * @param {object} datos
   *   meta: { asunto, de, para[], cc[], recibido (Date|string), idMensaje }
   *   cuerpoTexto: string
   *   enlaces: [string | {url, etiqueta}]
   *   anexos: [{ nombre, tamano, tipo: 'archivo'|'nube'|'item', enLinea, bytes: Uint8Array, url, error }]
   * @param {object} opciones
   *   incluirImagenesEnLinea, convertirImagen(bytes, formatoOrigen, destino) -> Uint8Array,
   *   nombreDespacho, zonaHoraria, ahora, alProgresar(msg)
   */
  async function generarPdf(datos, opciones = {}) {
    opciones = Object.assign({ zonaHoraria: ZONA_HORARIA, incluirImagenesEnLinea: false }, opciones);
    const avisar = opciones.alProgresar || (() => {});
    const anexosEntrada = datos.anexos || [];

    const totalBytes = anexosEntrada.reduce((s, a) => s + (a.bytes ? a.bytes.length : 0), 0);
    if (totalBytes > LIMITE_BYTES) {
      throw new Error(`El correo pesa ${tamanoLegible(totalBytes)} y supera el límite de ${tamanoLegible(LIMITE_BYTES)} de esta versión.`);
    }

    const anexos = await prepararAnexos(anexosEntrada, opciones, avisar);

    // Enlaces: los del cuerpo + archivos en la nube
    const enlaces = normalizarEnlaces([
      ...(datos.enlaces || []),
      ...enlacesDeTexto(datos.cuerpoTexto),
      ...anexos.filter(a => a.estado === 'enlace' && a.url).map(a => ({ url: a.url, etiqueta: 'Archivo en la nube: ' + a.nombre }))
    ]);
    const radicados = detectarRadicados(datos.meta && datos.meta.asunto, datos.cuerpoTexto);

    const paginasAnexos = anexos.reduce((s, a) => s + (a.estado === 'incluido' ? a.paginas : 0), 0);

    // Dos pasadas: la carátula necesita conocer sus propias páginas
    avisar('Armando carátula y cuerpo del correo…');
    let paginasFrente = 2;
    let frente;
    for (let intento = 0; intento < 4; intento++) {
      const rangos = { finFrente: paginasFrente, porAnexo: {} };
      let cursor = paginasFrente + 1;
      anexos.forEach((a, i) => {
        if (a.estado === 'incluido') { rangos.porAnexo[i] = [cursor, cursor + a.paginas - 1]; cursor += a.paginas; }
      });
      frente = await construirFrente(datos, anexos, enlaces, radicados, rangos, paginasFrente + paginasAnexos, opciones);
      frente.rangos = rangos;
      if (frente.paginas === paginasFrente) break;
      paginasFrente = frente.paginas;
    }

    const doc = frente.doc;
    const marcadores = frente.marcas.map(mk => ({ titulo: mk.titulo, indicePagina: mk.pagina }));

    for (let i = 0; i < anexos.length; i++) {
      const a = anexos[i];
      if (a.estado !== 'incluido') continue;
      avisar(`Uniendo anexo ${i + 1}: ${a.nombre}`);
      const indiceInicio = doc.getPageCount();
      if (a.recurso.clase === 'pdf') {
        const paginas = await doc.copyPages(a.recurso.pdf, a.recurso.pdf.getPageIndices());
        paginas.forEach(p => doc.addPage(p));
      } else {
        const img = a.recurso.formato === 'png' ? await doc.embedPng(a.recurso.bytes) : await doc.embedJpg(a.recurso.bytes);
        const horizontal = img.width > img.height;
        const [pw, ph] = horizontal ? [CARTA[1], CARTA[0]] : CARTA;
        const margen = 36;
        const escala = Math.min((pw - 2 * margen) / img.width, (ph - 2 * margen) / img.height, 1.5);
        const w = img.width * escala, h = img.height * escala;
        const pagina = doc.addPage([pw, ph]);
        pagina.drawImage(img, { x: (pw - w) / 2, y: (ph - h) / 2, width: w, height: h });
      }
      marcadores.push({ titulo: `Anexo ${i + 1}: ${a.nombre}`, indicePagina: indiceInicio });
    }

    agregarMarcadores(doc, marcadores);

    const m = datos.meta || {};
    doc.setTitle(m.asunto || 'Correo electrónico');
    doc.setSubject('Correo de ' + (m.de || 'remitente desconocido') + ' con sus anexos');
    doc.setAuthor('Correo a PDF v' + VERSION);
    doc.setCreator('Correo a PDF v' + VERSION + ' — Rama Judicial (prototipo)');
    doc.setProducer('pdf-lib');
    if (radicados.length) doc.setKeywords(radicados.map(formatoRadicado));
    doc.setCreationDate(opciones.ahora || new Date());

    avisar('Guardando PDF…');
    const bytes = await doc.save();
    const huella = await sha256(bytes);

    const p = partesFecha(m.recibido || new Date(), opciones.zonaHoraria);
    const nombreArchivo = `${p.year}${p.month}${p.day}_${p.hour}${p.minute}_` +
      `${radicados[0] || 'sin-radicado'}_${slug(m.asunto, 40)}.pdf`;

    return {
      bytes,
      nombreArchivo,
      sha256: huella,
      totalPaginas: doc.getPageCount(),
      radicados: radicados.map(formatoRadicado),
      enlaces,
      anexos: anexos.map(a => ({
        nombre: a.nombre, descripcion: a.descripcion, tamano: a.tamano, estado: a.estado,
        motivo: a.motivo, sha256: a.sha256, paginas: a.estado === 'incluido' ? a.paginas : 0
      }))
    };
  }

  return {
    VERSION, LIMITE_BYTES,
    generarPdf, prediccionPorNombre, detectarTipo, detectarRadicados, formatoRadicado,
    normalizarEnlaces, enlacesDeTexto, tamanoLegible, orientacionExif
  };
});
