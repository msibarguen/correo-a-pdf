/*!
 * Correo a PDF — lector de correos EML / MIME (v0.2.0)
 *
 * Obtiene del mensaje original: el HTML (o texto) del cuerpo y todas las partes
 * adjuntas con su Content-ID, para poder ubicar las imágenes del cuerpo en su sitio.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.CorreoMime = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  const latin1 = bytes => {
    let s = '';
    for (let i = 0; i < bytes.length; i += 8192) s += String.fromCharCode.apply(null, bytes.subarray(i, i + 8192));
    return s;
  };
  const deLatin1 = s => { const b = new Uint8Array(s.length); for (let i = 0; i < s.length; i++) b[i] = s.charCodeAt(i) & 0xFF; return b; };

  function decodificarCharset(bytes, charset) {
    const cs = (charset || 'utf-8').toLowerCase().replace(/^"|"$/g, '');
    const alias = { 'us-ascii': 'utf-8', ascii: 'utf-8', 'iso-8859-1': 'windows-1252', latin1: 'windows-1252', 'cp1252': 'windows-1252' };
    try { return new TextDecoder(alias[cs] || cs).decode(bytes); }
    catch (e) { return new TextDecoder('utf-8').decode(bytes); }
  }

  function base64(s) {
    let limpio = s.replace(/[^A-Za-z0-9+/]/g, '');
    if (limpio.length % 4 === 1) limpio = limpio.slice(0, -1);
    limpio += '='.repeat((4 - limpio.length % 4) % 4);
    if (typeof atob === 'function') return deLatin1(atob(limpio));
    return new Uint8Array(Buffer.from(limpio, 'base64'));
  }

  function quotedPrintable(s) {
    const t = s.replace(/=\r?\n/g, '');
    const out = [];
    for (let i = 0; i < t.length; i++) {
      if (t[i] === '=' && /^[0-9A-Fa-f]{2}$/.test(t.substr(i + 1, 2))) { out.push(parseInt(t.substr(i + 1, 2), 16)); i += 2; }
      else out.push(t.charCodeAt(i) & 0xFF);
    }
    return new Uint8Array(out);
  }

  // RFC 2047: =?utf-8?B?...?= / =?iso-8859-1?Q?...?=
  function palabrasCodificadas(s) {
    if (!s) return s;
    return s.replace(/\?=\s+=\?/g, '?==?').replace(/=\?([^?]+)\?([BbQq])\?([^?]*)\?=/g, (m, cs, enc, txt) => {
      try {
        const bytes = enc.toUpperCase() === 'B' ? base64(txt) : quotedPrintable(txt.replace(/_/g, ' '));
        return decodificarCharset(bytes, cs);
      } catch (e) { return m; }
    });
  }

  function leerEncabezados(bloque) {
    const h = {};
    bloque.replace(/\r?\n[ \t]+/g, ' ').split(/\r?\n/).forEach(l => {
      const i = l.indexOf(':');
      if (i > 0) {
        const k = l.slice(0, i).trim().toLowerCase();
        h[k] = h[k] ? h[k] + ', ' + l.slice(i + 1).trim() : l.slice(i + 1).trim();
      }
    });
    return h;
  }

  // valor; clave=valor; clave*=utf-8''valor (RFC 2231)
  function parametros(valor) {
    const partes = (valor || '').match(/("[^"]*"|[^;])+/g) || [];
    const r = { valor: (partes.shift() || '').trim().toLowerCase(), p: {} };
    const continuaciones = {};
    partes.forEach(p => {
      const i = p.indexOf('=');
      if (i < 0) return;
      let k = p.slice(0, i).trim().toLowerCase();
      let v = p.slice(i + 1).trim().replace(/^"|"$/g, '');
      const m = /^(.+?)\*(\d+)?\*?$/.exec(k);
      if (m && /\*/.test(k)) {
        const base = m[1];
        (continuaciones[base] = continuaciones[base] || []).push({ n: +(m[2] || 0), v, enc: /\*$/.test(k) });
        return;
      }
      r.p[k] = v;
    });
    Object.entries(continuaciones).forEach(([k, lista]) => {
      lista.sort((a, b) => a.n - b.n);
      let txt = lista.map(x => x.v).join('');
      if (lista[0].enc) {
        const m = /^([^']*)'[^']*'(.*)$/.exec(txt);
        if (m) {
          try {
            const bytes = deLatin1(m[2].replace(/%([0-9A-Fa-f]{2})/g, (_, h) => String.fromCharCode(parseInt(h, 16))));
            txt = decodificarCharset(bytes, m[1] || 'utf-8');
          } catch (e) { /* dejar */ }
        }
      }
      r.p[k] = txt;
    });
    return r;
  }

  function dividir(raw) {
    const m = /\r?\n\r?\n/.exec(raw);
    if (!m) return { encabezados: leerEncabezados(raw), cuerpo: '' };
    return { encabezados: leerEncabezados(raw.slice(0, m.index)), cuerpo: raw.slice(m.index + m[0].length) };
  }

  function contenido(parte, encabezados) {
    const te = (encabezados['content-transfer-encoding'] || '7bit').toLowerCase().trim();
    if (te === 'base64') return base64(parte);
    if (te === 'quoted-printable') return quotedPrintable(parte);
    return deLatin1(parte);
  }

  /**
   * @param {Uint8Array|string} eml  Mensaje completo
   * @returns {{ html: string|null, texto: string|null, partes: Array }}
   *   partes: { nombre, tipo, bytes, contentId, enLinea, esCorreo }
   */
  function analizar(eml) {
    const raw = typeof eml === 'string' ? eml : latin1(eml);
    const resultado = { html: null, texto: null, partes: [], encabezados: {} };

    function recorrer(bloque, nivel) {
      const { encabezados, cuerpo } = dividir(bloque);
      if (nivel === 0) resultado.encabezados = encabezados;
      const ct = parametros(encabezados['content-type'] || 'text/plain; charset=us-ascii');
      const cd = parametros(encabezados['content-disposition'] || '');
      const nombre = palabrasCodificadas(cd.p.filename || ct.p.name || '');
      const contentId = (encabezados['content-id'] || '').trim().replace(/^<|>$/g, '');

      if (ct.valor.startsWith('multipart/') && ct.p.boundary) {
        const b = '--' + ct.p.boundary;
        const trozos = cuerpo.split(new RegExp('\\r?\\n?' + b.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '(?:--)?[ \\t]*\\r?\\n?'));
        trozos.slice(1).forEach(t => { if (t.trim()) recorrer(t, nivel + 1); });
        return;
      }

      const esAdjunto = cd.valor === 'attachment' || (!!nombre && !ct.valor.startsWith('text/'));
      if (!esAdjunto && ct.valor === 'text/html' && resultado.html == null) {
        resultado.html = decodificarCharset(contenido(cuerpo, encabezados), ct.p.charset);
        return;
      }
      if (!esAdjunto && ct.valor === 'text/plain' && resultado.texto == null && cd.valor !== 'attachment') {
        resultado.texto = decodificarCharset(contenido(cuerpo, encabezados), ct.p.charset);
        return;
      }
      if (ct.valor === 'message/rfc822') {
        const interno = dividir(cuerpo).encabezados;
        resultado.partes.push({
          nombre: nombre || ((palabrasCodificadas(interno.subject || 'correo adjunto')) + '.eml'),
          tipo: ct.valor, bytes: deLatin1(cuerpo), contentId, enLinea: false, esCorreo: true
        });
        return;
      }
      if (nivel === 0 && !esAdjunto) return;
      resultado.partes.push({
        nombre: nombre || ('adjunto' + (resultado.partes.length + 1)),
        tipo: ct.valor,
        bytes: contenido(cuerpo, encabezados),
        contentId,
        enLinea: cd.valor === 'inline' || (!!contentId && cd.valor !== 'attachment'),
        esCorreo: false
      });
    }

    recorrer(raw, 0);
    return resultado;
  }

  return { analizar, palabrasCodificadas };
});
