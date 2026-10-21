/*!
 * Correo a PDF — motor de armado (core)
 * Prototipo v0.3.0 — Unidad de Transformación Digital, Rama Judicial
 *
 * Arma un único PDF (tamaño oficio) con:
 *   1. El correo tal como se ve: encabezado tipo Outlook y cuerpo HTML redibujado
 *      como texto vectorial (tablas, negritas, firmas e imágenes en su sitio)
 *   2. Anexos PDF e imágenes, unidos en orden
 *   3. Hoja de control (temporal): índice, estado de cada archivo y huellas SHA-256
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

  const Fuentes = (typeof CorreoFuentes !== 'undefined') ? CorreoFuentes : require('./fuentes.js');

  const VERSION = '0.3.0';
  const LIMITE_BYTES = 100 * 1024 * 1024; // 100 MB por correo
  const PAGINA = [612, 936]; // oficio colombiano: 21,6 x 33 cm
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
      this.anchoUtil = PAGINA[0] - 2 * MARGEN;
      this.paginas = [];
      this.cacheCaracteres = new Map();
      this.nuevaPagina();
    }

    nuevaPagina() {
      this.pagina = this.doc.addPage(PAGINA);
      this.paginas.push(this.pagina);
      this.y = PAGINA[1] - MARGEN;
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
      this.pagina.drawLine({ start: { x: MARGEN, y: this.y }, end: { x: PAGINA[0] - MARGEN, y: this.y }, thickness: 0.6, color: COLOR.linea });
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

      if (a.enCuerpo) {
        r.estado = 'en_cuerpo';
        r.motivo = 'Imagen del cuerpo del correo: se muestra en su lugar dentro del correo.';
        resultado.push(r);
        continue;
      }

      if (a.enLinea && !opciones.incluirImagenesEnLinea) {
        r.estado = 'omitido';
        r.motivo = 'Marcada por Outlook como imagen incrustada, pero no aparece en el cuerpo del correo. No se incluyó.';
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

  // ---------------------------------------------------------------------------
  // Encabezado tipo Outlook (se maqueta con el mismo motor que el cuerpo)
  // ---------------------------------------------------------------------------

  const esc = s => String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

  function fechaOutlook(fecha, zona) {
    const d = fecha instanceof Date ? fecha : new Date(fecha);
    if (isNaN(d)) return String(fecha || '');
    const dia = new Intl.DateTimeFormat('es-CO', { timeZone: zona, weekday: 'short' }).format(d).replace(/\./g, '');
    const p = partesFecha(d, zona);
    return `${dia.charAt(0).toUpperCase()}${dia.slice(1)} ${p.day}/${p.month}/${p.year} ${p.hour === '24' ? '00' : p.hour}:${p.minute}`;
  }

  const ICONO_CLIP = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAADE0lEQVR42u2XTYscVRSG31O37p2eqvGjDQr9AwTtH6DgRogRIYEgSiHOpqph6N3gwqWasmUI4g8QUjDUNLMsFJIouEwLLl0JvRyGQWlQWwxY33XvcWGPNEoy6S835l3WF89565xTbwGLizzPEwAI/zfRedVBEFz2ff/pueNLy17Adqvb7dLZ2dltAFeUUs+toyrrUS7yPM9KkkSfnp5+0Wq1rhpjkKYp/ycAnueJJEl0r9d7VSl1vaoqQ0TfTKfTMgxDCwBv3IH9/f0tY8yAiFhr/atS6t08z/V4PLZnECs11kXn2fO8p1zX/anVarl5nt8cDocfzlfe7/dlp9Ph8XjM3W6XJpMJdTodPRgMzFqaUCnFzJxbluUCyACw7/sfuK5rp2n6VRRF3296Cv71umzbPhBCAMDbvV7vSymllef5d0T0/M7OzqU8z+/FcTwKw9B6mBP2suRN09zPssyxLKtm5tBxHJRleQvAW47jPJtl2WcARpPJRAAwKzXhg+5VSkkAd5n5h6qqaiK6D+DnPM9rZv5jbVPwIDEzmPlJAC0AEoCYuSqJyNo4wEyGiHije2CTegzwGGAtU7DKF3EVgIaZebYDBDMbZl44HdnLQksp23Vdg5l7RGTNFo9ZNKIt4gABABHlzExa6wNm1o7j7Liu62RZdkcI8QmA9l/GYH2bcGtriwE0Wmsw82Ui4jiOP9JaXzXGDLIsuxHH8ZtE9KKU8gmtNQAU6wgk6Pf7MoqiJgiCG0qpsGkaMsZcPzo6uvvPYoIg+FYp9UpZllNjzAvHx8fT81CzNMD5A3Z3d9tSyh+llE7TNAWAT5VSUV3XjTHmJaXUe03TvLa9vW2laToYDocfz+DrlRyYD6a+71+zbfu2mCWRqqp+B6ABPKOUIiEEiqL4uq7rd8qyLJIkuXBEH7ljz5NNEASvSynfN8ZcUUoJADDGoCiKXyzL+vzk5ORgNBo1F1m/MMC8EwCwt7f3hhDi5aqqwMy1EOLW4eHhb/OvbSOr82ExPAxDe9Gilv6v8zxPtNvtv2GiKGo2VvUm9SddQGucT33L4QAAAABJRU5ErkJggg==';

  function tamanoOutlook(bytes) {
    if (!(bytes > 0)) return '0 KB';
    if (bytes < 1024 * 1024) return Math.max(1, Math.round(bytes / 1024)) + ' KB';
    return (bytes / (1024 * 1024)).toFixed(1).replace('.', ',') + ' MB';
  }

  /** Línea de adjuntos como la muestra Outlook: "2 archivos adjuntos (222 KB)" y los nombres. */
  function htmlAdjuntos(adjuntos) {
    if (!adjuntos || !adjuntos.length) return '';
    const total = adjuntos.reduce((s, a) => s + (a.tamano || 0), 0);
    const titulo = adjuntos.length === 1 ? '1 archivo adjunto' : `${adjuntos.length} archivos adjuntos`;
    return `<div style="margin-top:14px;font-size:11px;color:#424242">
<div style="margin-bottom:3px"><img src="${ICONO_CLIP}" width="10" height="10" style="width:10px;height:10px;vertical-align:-1px;margin-right:5px">${titulo} (${tamanoOutlook(total)})</div>
<div>${adjuntos.map(a => esc(a.nombre) + ';').join(' ')}</div></div>`;
  }

  function htmlEncabezado(meta, zona, adjuntos) {
    const m = meta || {};
    const fila = (etq, val) => val ? `<tr><td style="font-weight:700;padding:2px 12px 2px 2px;vertical-align:top;white-space:nowrap">${etq}</td>` +
      `<td style="padding:2px 0;vertical-align:top">${esc(val)}</td></tr>` : '';
    return `<div style="font-family:'Segoe UI',sans-serif;color:#242424;font-size:12px">
<div style="border-top:1px solid #c8c8c8;border-bottom:1px solid #c8c8c8;padding:9px 0 9px 2px;margin-bottom:12px;font-size:15px;font-weight:700;line-height:1.35">${esc(m.asunto || '(sin asunto)')}</div>
<table style="border-collapse:collapse;font-size:12px;line-height:1.4">
${fila('Desde', m.de)}${fila('Fecha', m.recibido ? fechaOutlook(m.recibido, zona || ZONA_HORARIA) : '')}${fila('Para', (m.para || []).join('; '))}${fila('CC', (m.cc || []).join('; '))}
</table>${htmlAdjuntos(adjuntos)}<div style="height:18px"></div></div>`;
  }

  // ---------------------------------------------------------------------------
  // Composición del correo en páginas (texto vectorial con fuentes incrustadas)
  // ---------------------------------------------------------------------------

  const M = { sup: 40, inf: 46, izq: 40, der: 40 };
  const ANCHO_PT = PAGINA[0] - M.izq - M.der;   // 532 pt
  const ANCHO_PX = Math.round(ANCHO_PT / 0.75); // ancho de maquetación en píxeles CSS
  const ALTO_UTIL = PAGINA[1] - M.sup - M.inf;

  function aplanar(segmentos) {
    const flujo = [];
    let desfase = 0;
    segmentos.forEach(seg => {
      const k = (ANCHO_PT / (seg.anchoPx || ANCHO_PX)) * (seg.escala || 1);
      seg.items.forEach(it => {
        const n = Object.assign({}, it);
        n.x = it.x * k; n.y = it.y * k + desfase;
        if (it.w != null) n.w = it.w * k;
        if (it.h != null) n.h = it.h * k;
        if (it.tam != null) n.tam = it.tam * k;
        if (it.ancho != null) n.ancho = it.ancho * k;
        n.sup = it.sup * k + desfase; n.inf = it.inf * k + desfase;
        flujo.push(n);
      });
      desfase += seg.altoPx * k;
    });
    return { flujo, alto: desfase };
  }

  function calcularCortes(flujo, alto) {
    const intervalos = flujo
      .filter(it => (it.t === 'texto' || it.t === 'img') && (it.inf - it.sup) < ALTO_UTIL * 0.9)
      .map(it => [it.sup, it.inf]);
    // Fin real del contenido (evita una última página en blanco)
    const finContenido = flujo.reduce((mx, it) => it.t === 'enlace' ? mx : Math.max(mx, it.inf), 0);
    const total = Math.min(alto, finContenido + 4) || alto;
    const cortes = [];
    let ini = 0;
    while (ini < total - 0.5) {
      let c = ini + ALTO_UTIL;
      if (c >= total) { cortes.push([ini, total]); break; }
      for (let i = 0; i < 60; i++) {
        const choques = intervalos.filter(([a, b]) => a < c - 0.01 && b > c + 0.01);
        if (!choques.length) break;
        const nc = Math.min(...choques.map(([a]) => a));
        if (nc <= ini + ALTO_UTIL * 0.25) { c = ini + ALTO_UTIL; break; }
        c = nc;
      }
      cortes.push([ini, c]);
      ini = c;
    }
    if (!cortes.length) cortes.push([0, Math.max(total, 1)]);
    return cortes;
  }

  async function componerCorreo(doc, segmentos, imagenes, gestor) {
    const { pushGraphicsState, popGraphicsState, rectangle, clip, endPath } = PDFLib;
    const { flujo, alto } = aplanar(segmentos);
    // Enlazar cada palabra decorada con la siguiente palabra decorada (para líneas continuas)
    let previo = null;
    flujo.forEach(it => {
      if (it.t !== 'texto') return;
      if (previo && (previo.sub || previo.tach) && ((it.sub && previo.sub) || (it.tach && previo.tach))) previo.siguiente = it;
      previo = it;
    });
    const cortes = calcularCortes(flujo, alto);
    const nombres = new Map();
    const embebidas = new Map();
    const paginas = [];

    async function imagen(clave) {
      if (embebidas.has(clave)) return embebidas.get(clave);
      const im = imagenes.get(clave);
      let e = null;
      if (im) {
        try { e = im.formato === 'jpg' ? await doc.embedJpg(im.bytes) : await doc.embedPng(im.bytes); } catch (err) { e = null; }
      }
      embebidas.set(clave, e);
      return e;
    }

    for (const [ini, fin] of cortes) {
      const pagina = doc.addPage(PAGINA);
      paginas.push(pagina);
      const tope = PAGINA[1] - M.sup;
      const Y = y => tope - (y - ini);
      pagina.pushOperators(pushGraphicsState(), rectangle(M.izq - 3, tope - (fin - ini) - 1, ANCHO_PT + 6, (fin - ini) + 2), clip(), endPath());
      for (const it of flujo) {
        if (it.t === 'enlace' || it.inf <= ini || it.sup >= fin) continue;
        if (it.t === 'rect') {
          if (it.guiones && (it.w > it.h * 2 || it.h > it.w * 2)) {
            const horizontal = it.w >= it.h;
            const grosor = horizontal ? it.h : it.w;
            const patron = it.guiones === 1 ? [grosor, grosor] : [grosor * 3, grosor * 2];
            pagina.drawLine({
              start: { x: M.izq + it.x + (horizontal ? 0 : grosor / 2), y: Y(it.y + (horizontal ? grosor / 2 : 0)) },
              end: { x: M.izq + it.x + (horizontal ? it.w : grosor / 2), y: Y(it.y + (horizontal ? grosor / 2 : it.h)) },
              thickness: grosor, color: rgb(...it.color), dashArray: patron
            });
          } else {
            pagina.drawRectangle({ x: M.izq + it.x, y: Y(it.y + it.h), width: it.w, height: it.h, color: rgb(...it.color) });
          }
        } else if (it.t === 'img') {
          const e = await imagen(it.clave);
          if (e) pagina.drawImage(e, { x: M.izq + it.x, y: Y(it.y + it.h), width: it.w, height: it.h });
        } else if (it.t === 'texto') {
          Fuentes.dibujarTexto(pagina, gestor, it.cara, it.txt, M.izq + it.x, Y(it.y), it.tam, it.color, nombres);
          const grosor = Math.max(0.45, it.tam * 0.055);
          // Subrayado/tachado continuo hasta la siguiente palabra decorada de la misma línea
          let anchoLinea = it.ancho != null ? it.ancho : gestor.anchoTexto(it.cara, it.txt.trimEnd(), it.tam);
          const sig = it.siguiente;
          if (sig && /\s$/.test(it.txt) && Math.abs(sig.y - it.y) < 0.6 && sig.x > it.x && sig.x - (it.x + anchoLinea) < it.tam * 0.8) {
            anchoLinea = sig.x - it.x;
          }
          if (it.sub) pagina.drawRectangle({ x: M.izq + it.x, y: Y(it.y) - it.tam * 0.11 - grosor, width: anchoLinea, height: grosor, color: rgb(...it.sub) });
          if (it.tach && !(sig && sig.tach)) anchoLinea = it.ancho != null ? it.ancho : anchoLinea;
          if (it.tach) pagina.drawRectangle({ x: M.izq + it.x, y: Y(it.y) + it.tam * 0.27, width: anchoLinea, height: grosor, color: rgb(...it.tach) });
        }
      }
      pagina.pushOperators(popGraphicsState());
      // Enlaces clicables
      flujo.forEach(it => {
        if (it.t !== 'enlace' || it.sup < ini - 0.5 || it.sup >= fin) return;
        const caja = { pagina, x: M.izq + it.x, y: Y(Math.min(it.inf, fin)), w: it.w, h: Math.min(it.inf, fin) - it.sup };
        agregarEnlace(doc, caja, it.url);
      });
    }
    return paginas.length;
  }

  // ---------------------------------------------------------------------------
  // Hoja de control (temporal, al final)
  // ---------------------------------------------------------------------------

  async function construirHojaControl(datos, anexos, enlaces, radicados, rangos, opciones) {
    const doc = await PDFDocument.create();
    const fuentes = {
      normal: await doc.embedFont(StandardFonts.Helvetica),
      negrita: await doc.embedFont(StandardFonts.HelveticaBold),
      mono: await doc.embedFont(StandardFonts.Courier)
    };
    const e = new Escritor(doc, fuentes);
    const m = datos.meta || {};
    const pags = r => r[0] === r[1] ? 'página ' + r[0] : 'páginas ' + r[0] + ' a ' + r[1];

    e.texto('RAMA JUDICIAL DEL PODER PÚBLICO', { negrita: true, tam: 8.5, color: COLOR.suave });
    e.texto(opciones.nombreDespacho || 'Secretaría', { tam: 8.5, color: COLOR.suave });
    e.espacio(6);
    e.texto('HOJA DE CONTROL DEL DOCUMENTO', { negrita: true, tam: 15 });
    e.texto('Información de verificación generada automáticamente (versión de prueba).', { tam: 8.5, color: COLOR.suave });
    e.linea();

    e.campo('Asunto', m.asunto || '(sin asunto)');
    e.campo('De', m.de || '—');
    e.campo('Para', (m.para || []).join('; ') || '—');
    if (m.cc && m.cc.length) e.campo('CC', m.cc.join('; '));
    e.campo('Recibido', fechaLegible(m.recibido, opciones.zonaHoraria));
    if (m.idMensaje) e.campo('ID del mensaje', m.idMensaje);
    e.campo('Radicado detectado', radicados.length
      ? radicados.map(formatoRadicado).join('  |  ') + '  (verificar)'
      : 'No se detectó un número de radicado de 23 dígitos.');
    e.campo('Fuente del correo', datos.origen === 'eml' ? 'Mensaje original completo (MIME)' : 'Cuerpo y adjuntos leídos por separado');

    const noConvertidos = anexos.filter(a => a.estado === 'no_convertido');
    if (noConvertidos.length) {
      e.espacio(4);
      e.recuadroAlerta([
        `ATENCIÓN: ${noConvertidos.length} archivo(s) recibido(s) NO se convirtieron y NO están incluidos en este PDF.`,
        'Revise la lista de archivos. Conserve el correo original y gestione esos archivos por separado.'
      ]);
    }
    if (datos.externas) {
      e.espacio(2);
      e.texto(`Nota: ${datos.externas} imagen(es) externa(s) del cuerpo (alojadas en internet) no se incluyeron; aparecen como un recuadro.`, { tam: 8.5, color: COLOR.suave });
    }

    e.titulo('ÍNDICE DEL DOCUMENTO');
    e.texto(`Correo electrónico .......... ${pags(rangos.correo)}`, { tam: 9.5 });
    anexos.forEach((a, i) => { if (a.estado === 'incluido') e.texto(`Anexo ${rangos.numero[i]}: ${a.nombre} .......... ${pags(rangos.porAnexo[i])}`, { tam: 9.5 }); });
    e.texto(`Hoja de control .......... ${pags(rangos.hoja)}`, { tam: 9.5 });
    e.texto(`Total de páginas: ${rangos.total}`, { tam: 9.5, negrita: true });

    e.titulo(`ARCHIVOS RECIBIDOS EN EL CORREO (${anexos.length})`);
    e.texto('El SHA-256 es la huella digital de cada archivo tal como llegó; permite verificar que un archivo es el mismo recibido.', { tam: 7.5, color: COLOR.suave });
    e.espacio(4);
    if (!anexos.length) e.texto('El correo no traía archivos adjuntos.', { tam: 9.5, color: COLOR.suave });
    anexos.forEach((a, i) => {
      e.asegurar(48);
      e.texto(`${i + 1}. ${a.nombre}`, { negrita: true, tam: 9.5 });
      e.texto(`Tipo: ${a.descripcion || '—'}   ·   Tamaño: ${tamanoLegible(a.tamano)}`, { tam: 8.5, sangria: 12, color: COLOR.suave });
      let estado, color;
      if (a.estado === 'incluido') { estado = `Incluido como anexo ${rangos.numero[i]} (${pags(rangos.porAnexo[i])})`; color = COLOR.ok; }
      else if (a.estado === 'en_cuerpo') { estado = 'Imagen del cuerpo: se muestra en su lugar dentro del correo'; color = COLOR.ok; }
      else if (a.estado === 'enlace') { estado = 'Archivo en la nube: incluido como enlace'; color = COLOR.enlace; }
      else if (a.estado === 'omitido') { estado = 'Omitido: ' + a.motivo; color = COLOR.suave; }
      else { estado = 'NO CONVERTIDO: ' + a.motivo; color = COLOR.alerta; }
      e.texto(estado, { tam: 8.5, sangria: 12, color, negrita: a.estado === 'no_convertido' });
      if (a.sha256) e.texto('SHA-256: ' + a.sha256, { tam: 7, sangria: 12, mono: true, color: COLOR.suave });
      e.espacio(3);
    });

    e.titulo(`ENLACES DEL CORREO (${enlaces.length})`);
    if (!enlaces.length) e.texto('El correo no contiene enlaces.', { tam: 9.5, color: COLOR.suave });
    else {
      e.texto('Los enlaces no fueron abiertos ni descargados; su contenido no está incluido en este PDF.', { tam: 8.5, color: COLOR.suave });
      e.espacio(3);
      enlaces.forEach((l, i) => {
        e.asegurar(24);
        if (l.etiqueta) e.texto(`${i + 1}. ${l.etiqueta}`, { tam: 9, negrita: true });
        const cajas = e.texto((l.etiqueta ? '   ' : `${i + 1}. `) + l.url, { tam: 8.5, color: COLOR.enlace });
        cajas.forEach(c => agregarEnlace(doc, c, l.url));
        if (l.safelinks) e.texto('(enlace original, extraído de la protección Safe Links de Microsoft)', { tam: 7.5, sangria: 12, color: COLOR.suave });
        e.espacio(1);
      });
    }
    return { doc, paginas: e.paginas.length };
  }

  // ---------------------------------------------------------------------------
  // Función principal
  // ---------------------------------------------------------------------------

  /**
   * @param {object} datos
   *   meta, segmentos [{items, altoPx, escala, anchoPx}], imagenes Map(clave→{bytes,formato}),
   *   fuentes {cara: Uint8Array}, anexos [...], enlaces [...], cuerpoTexto, externas, origen
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
    const enlaces = normalizarEnlaces([
      ...(datos.enlaces || []),
      ...anexos.filter(a => a.estado === 'enlace' && a.url).map(a => ({ url: a.url, etiqueta: 'Archivo en la nube: ' + a.nombre }))
    ]);
    const m = datos.meta || {};
    const radicados = detectarRadicados(m.asunto, datos.cuerpoTexto);

    avisar('Dibujando el correo…');
    const doc = await PDFDocument.create();
    const gestor = new Fuentes.GestorFuentes(doc, datos.fuentes || {});
    const paginasCorreo = await componerCorreo(doc, datos.segmentos || [], datos.imagenes || new Map(), gestor);

    // Anexos
    const marcadores = [{ titulo: 'Correo electrónico', indicePagina: 0 }];
    const rangos = { correo: [1, paginasCorreo], porAnexo: {}, numero: {} };
    let numero = 0;
    for (let i = 0; i < anexos.length; i++) {
      const a = anexos[i];
      if (a.estado !== 'incluido') continue;
      numero++;
      avisar(`Uniendo anexo ${numero}: ${a.nombre}`);
      const inicio = doc.getPageCount();
      if (a.recurso.clase === 'pdf') {
        const paginas = await doc.copyPages(a.recurso.pdf, a.recurso.pdf.getPageIndices());
        paginas.forEach(p => doc.addPage(p));
      } else {
        const img = a.recurso.formato === 'png' ? await doc.embedPng(a.recurso.bytes) : await doc.embedJpg(a.recurso.bytes);
        const horizontal = img.width > img.height;
        const [pw, ph] = horizontal ? [PAGINA[1], PAGINA[0]] : PAGINA;
        const margen = 36;
        const escala = Math.min((pw - 2 * margen) / img.width, (ph - 2 * margen) / img.height, 1.5);
        const w = img.width * escala, h = img.height * escala;
        const pagina = doc.addPage([pw, ph]);
        pagina.drawImage(img, { x: (pw - w) / 2, y: (ph - h) / 2, width: w, height: h });
      }
      rangos.porAnexo[i] = [inicio + 1, doc.getPageCount()];
      rangos.numero[i] = numero;
      marcadores.push({ titulo: `Anexo ${numero}: ${a.nombre}`, indicePagina: inicio });
    }

    // Hoja de control (dos pasadas: necesita conocer su propio número de páginas)
    avisar('Armando la hoja de control…');
    const inicioHoja = doc.getPageCount() + 1;
    let paginasHoja = 1, hoja;
    for (let intento = 0; intento < 4; intento++) {
      rangos.hoja = [inicioHoja, inicioHoja + paginasHoja - 1];
      rangos.total = inicioHoja + paginasHoja - 1;
      hoja = await construirHojaControl(datos, anexos, enlaces, radicados, rangos, opciones);
      if (hoja.paginas === paginasHoja) break;
      paginasHoja = hoja.paginas;
    }
    const copiadas = await doc.copyPages(hoja.doc, hoja.doc.getPageIndices());
    copiadas.forEach(p => doc.addPage(p));
    marcadores.push({ titulo: 'Hoja de control', indicePagina: inicioHoja - 1 });

    // Encabezado y pie en las páginas propias (correo y hoja de control); los anexos no se tocan
    const total = doc.getPageCount();
    const helv = await doc.embedFont(StandardFonts.Helvetica);
    const gris = rgb(0.25, 0.25, 0.27), grisSuave = COLOR.suave;
    const ahora = opciones.ahora || new Date();
    const pf = partesFecha(ahora, opciones.zonaHoraria);
    const fechaImpresion = `${pf.day}/${pf.month}/${pf.year.slice(-2)}, ${pf.hour === '24' ? '00' : pf.hour}:${pf.minute}`;
    const tituloCuenta = opciones.nombreCuenta ? `Correo: ${opciones.nombreCuenta} - Outlook` : '';
    const enlace = opciones.enlaceCorreo || '';
    const recortar = (txt, tam, ancho) => {
      if (helv.widthOfTextAtSize(txt, tam) <= ancho) return txt;
      let s = txt;
      while (s.length > 1 && helv.widthOfTextAtSize(s + '...', tam) > ancho) s = s.slice(0, -1);
      return s + '...';
    };
    const marcarPagina = (i) => {
      const p = doc.getPage(i);
      const { width, height } = p.getSize();
      // Encabezado: fecha y hora de generación · cuenta de Outlook
      p.drawText(fechaImpresion, { x: M.izq - 18, y: height - 20, size: 7.5, font: helv, color: gris });
      if (tituloCuenta) {
        const t = recortar(tituloCuenta, 7.5, width - 2 * M.izq - 120);
        p.drawText(t, { x: (width - helv.widthOfTextAtSize(t, 7.5)) / 2, y: height - 20, size: 7.5, font: helv, color: gris });
      }
      // Pie: enlace al mensaje · número de página
      const num = `Página ${i + 1} de ${total}`;
      const anchoNum = helv.widthOfTextAtSize(num, 7.5);
      p.drawText(num, { x: width - M.der + 18 - anchoNum, y: 26, size: 7.5, font: helv, color: gris });
      if (enlace) {
        const t = recortar(enlace, 6.5, width - M.izq - M.der + 36 - anchoNum - 16);
        p.drawText(t, { x: M.izq - 18, y: 26, size: 6.5, font: helv, color: gris });
        agregarEnlace(doc, { pagina: p, x: M.izq - 18, y: 24, w: helv.widthOfTextAtSize(t, 6.5), h: 9 }, enlace);
      }
      p.drawText(`Generado con Correo a PDF v${VERSION} · Rama Judicial (prototipo)`, { x: M.izq - 18, y: 15, size: 6.5, font: helv, color: grisSuave });
    };
    for (let i = 0; i < paginasCorreo; i++) marcarPagina(i);
    for (let i = inicioHoja - 1; i < total; i++) marcarPagina(i);

    gestor.finalizar();
    agregarMarcadores(doc, marcadores);
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
      bytes, nombreArchivo, sha256: huella, totalPaginas: total, paginasCorreo,
      radicados: radicados.map(formatoRadicado), enlaces,
      anexos: anexos.map(a => ({
        nombre: a.nombre, descripcion: a.descripcion, tamano: a.tamano, estado: a.estado,
        motivo: a.motivo, sha256: a.sha256, paginas: a.estado === 'incluido' ? a.paginas : 0
      }))
    };
  }

  return {
    VERSION, LIMITE_BYTES, PAGINA, ANCHO_PX,
    generarPdf, htmlEncabezado, htmlAdjuntos, tamanoOutlook, fechaOutlook, prediccionPorNombre, detectarTipo, detectarRadicados, formatoRadicado,
    normalizarEnlaces, enlacesDeTexto, tamanoLegible, orientacionExif, calcularCortes
  };
});
