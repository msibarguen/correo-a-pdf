/*!
 * Correo a PDF — motor de renderizado HTML → elementos vectoriales (v0.2.0)
 *
 * El navegador maqueta el correo en un marco aislado (sin scripts ni recursos
 * externos) usando fuentes libres con las mismas medidas que las de Microsoft.
 * Luego se toma la posición exacta de cada palabra, imagen, borde, fondo y enlace
 * para redibujarlos en el PDF como texto vectorial real.
 * Solo funciona en el navegador.
 */
(function (root) {
  'use strict';

  // Mapa de familias tipográficas del correo → familia libre equivalente
  const MAPA_FUENTES = {
    calibri: 'carlito', 'calibri light': 'carlito', aptos: 'carlito', 'aptos display': 'carlito', candara: 'carlito',
    carlito: 'carlito', corbel: 'carlito',
    cambria: 'caladea', 'cambria math': 'caladea', caladea: 'caladea', georgia: 'caladea', 'book antiqua': 'caladea',
    'palatino linotype': 'caladea', garamond: 'caladea',
    arial: 'libsans', helvetica: 'libsans', 'helvetica neue': 'libsans', 'arial nova': 'libsans', 'segoe ui': 'libsans',
    'segoe ui semibold': 'libsans', 'segoe ui light': 'libsans', 'liberation sans': 'libsans', 'century gothic': 'libsans',
    montserrat: 'libsans', 'open sans': 'libsans', roboto: 'libsans', lato: 'libsans', 'trebuchet ms': 'libsans',
    'gill sans mt': 'libsans', 'franklin gothic book': 'libsans', 'arial narrow': 'libsans', 'system-ui': 'libsans',
    '-apple-system': 'libsans', 'sans-serif': 'libsans', 'microsoft sans serif': 'libsans',
    'times new roman': 'libserif', times: 'libserif', 'liberation serif': 'libserif', serif: 'libserif',
    'bookman old style': 'libserif', 'century schoolbook': 'libserif',
    'courier new': 'libmono', courier: 'libmono', consolas: 'libmono', 'lucida console': 'libmono', monospace: 'libmono',
    'liberation mono': 'libmono',
    verdana: 'dejavu', tahoma: 'dejavu', 'dejavu sans': 'dejavu', 'arial unicode ms': 'dejavu',
    symbol: 'dejavu', wingdings: 'dejavu', webdings: 'dejavu'
  };
  // Viñetas de Word/Outlook escritas con fuentes de símbolos
  const SIMBOLOS = {
    symbol: { '·': '•', '': '•', 'o': '◦', '§': '▪', '': '▪' },
    wingdings: { '§': '▪', '': '▪', 'Ø': '➢', '': '➢', 'ü': '✓', '': '✓', 'v': '❖', 'q': '❑', 'n': '■', 'l': '●' }
  };

  const esperar = ms => new Promise(r => setTimeout(r, ms));
  const cuadro = w => new Promise(r => (w.requestAnimationFrame || setTimeout)(() => r(), 16));

  function color(txt) {
    const m = /rgba?\(([^)]+)\)/.exec(txt || '');
    if (!m) return null;
    const p = m[1].split(/[\s,/]+/).filter(Boolean).map(Number);
    const a = p.length > 3 ? p[3] : 1;
    return { rgb: [p[0] / 255, p[1] / 255, p[2] / 255], a };
  }

  function escaparHtml(s) {
    return String(s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  }

  function textoAHtml(texto) {
    const enlazado = escaparHtml(texto || '').replace(/https?:\/\/[^\s<>"']+/g, u => `<a href="${u}">${u}</a>`);
    return `<div style="font-family:Calibri,sans-serif;font-size:11pt;white-space:pre-wrap;word-wrap:break-word">${enlazado}</div>`;
  }

  function claveFamilia(listaCss) {
    const familias = (listaCss || '').split(',').map(f => f.trim().replace(/^["']|["']$/g, '').toLowerCase());
    for (const f of familias) {
      if (f.startsWith('cpdf-')) return f.slice(5);
      if (MAPA_FUENTES[f]) return MAPA_FUENTES[f];
    }
    return 'libsans';
  }

  function primeraFamilia(listaCss) {
    return ((listaCss || '').split(',')[0] || '').trim().replace(/^["']|["']$/g, '').toLowerCase();
  }

  // ------------------------------------------------------------------ imágenes del cuerpo

  /**
   * Busca, para un src de <img>, el adjunto que le corresponde.
   * candidatos: [{ clave, nombre, contentId, ancho, alto }]
   */
  function resolverImagen(src, candidatos, usadas) {
    const s = (src || '').trim();
    if (/^cid:/i.test(s)) {
      const id = decodeURIComponent(s.slice(4)).replace(/^<|>$/g, '').toLowerCase();
      const antesArroba = id.split('@')[0];
      let c = candidatos.find(x => x.contentId && x.contentId.toLowerCase() === id);
      if (!c) c = candidatos.find(x => x.contentId && x.contentId.toLowerCase().split('@')[0] === antesArroba);
      if (!c) c = candidatos.find(x => !usadas.has(x.clave) && x.nombre && x.nombre.toLowerCase() === antesArroba);
      return c || null;
    }
    // URLs internas de Outlook que incluyen el nombre o el id del adjunto
    if (!/^data:/i.test(s)) {
      const sl = s.toLowerCase();
      const c = candidatos.find(x => !usadas.has(x.clave) &&
        ((x.idOutlook && sl.includes(encodeURIComponent(x.idOutlook).toLowerCase())) ||
         (x.contentId && sl.includes(x.contentId.toLowerCase()))));
      if (c) return c;
    }
    return null;
  }

  // ------------------------------------------------------------------ limpieza del HTML

  function limpiarHtml(html, candidatos, diag) {
    const doc = new DOMParser().parseFromString(html || '', 'text/html');
    doc.querySelectorAll('script,noscript,iframe,frame,frameset,object,embed,applet,link,meta,base,title,audio,video,source,track')
      .forEach(n => n.remove());
    doc.querySelectorAll('*').forEach(el => {
      for (const at of Array.from(el.attributes)) {
        const n = at.name.toLowerCase();
        if (n.startsWith('on') || n === 'background' || n === 'srcset') el.removeAttribute(at.name);
        if ((n === 'href' || n === 'src') && /^\s*javascript:/i.test(at.value)) el.removeAttribute(at.name);
        if (n === 'style' && /url\s*\(/i.test(at.value)) el.setAttribute('style', at.value.replace(/url\s*\([^)]*\)/gi, 'none'));
      }
    });
    doc.querySelectorAll('style').forEach(st => {
      st.textContent = st.textContent.replace(/@import[^;]+;/gi, '').replace(/url\s*\([^)]*\)/gi, 'none');
    });

    const usadas = new Set();
    const pendientes = [];
    let externas = 0;
    doc.querySelectorAll('img').forEach((img, i) => {
      const src = img.getAttribute('src') || '';
      const c = resolverImagen(src, candidatos, usadas);
      diag.push({ src: src.startsWith('data:') ? 'data:…' : src.replace(/\?.*$/, '?…').slice(0, 160), resuelta: c ? c.nombre : null });
      if (c) {
        usadas.add(c.clave);
        img.setAttribute('src', c.url);
        img.setAttribute('data-cpdf-img', c.clave);
      } else if (/^data:image\//i.test(src)) {
        img.setAttribute('data-cpdf-img', 'data-' + i);
      } else if (/^(cid:|blob:)/i.test(src) || !/^https?:/i.test(src)) {
        pendientes.push(img); // referencia interna no resuelta: se intentará por tamaño
      } else {
        externas++;
        sustituirExterna(doc, img);
      }
    });

    // Segundo intento: emparejar referencias internas no resueltas por proporción de la imagen
    pendientes.forEach(img => {
      const w = parseFloat(img.getAttribute('width')), h = parseFloat(img.getAttribute('height'));
      const libres = candidatos.filter(x => !usadas.has(x.clave));
      let c = null;
      if (w && h && libres.length) {
        const r = w / h;
        c = libres.slice().sort((a, b) => Math.abs(a.ancho / a.alto - r) - Math.abs(b.ancho / b.alto - r))[0];
        if (Math.abs(c.ancho / c.alto - r) > 0.08 * r) c = null;
      } else if (libres.length === 1) c = libres[0];
      if (c) {
        usadas.add(c.clave);
        img.setAttribute('src', c.url);
        img.setAttribute('data-cpdf-img', c.clave);
        const d = diag.find(x => x.resuelta === null); if (d) d.resuelta = c.nombre + ' (por tamaño)';
      } else {
        externas++;
        sustituirExterna(doc, img, 'imagen no encontrada');
      }
    });

    const estilos = Array.from(doc.querySelectorAll('style')).map(s => s.textContent).join('\n');
    const atributosBody = Array.from((doc.body || {}).attributes || []).filter(a => !/^on/i.test(a.name))
      .map(a => `${a.name}="${escaparHtml(a.value)}"`).join(' ');
    return { estilos, cuerpo: doc.body ? doc.body.innerHTML : '', atributosBody, usadas, externas };
  }

  function sustituirExterna(doc, img, texto) {
    const w = parseFloat(img.getAttribute('width')) || 120, h = parseFloat(img.getAttribute('height')) || 24;
    const span = doc.createElement('span');
    span.setAttribute('style', `display:inline-block;width:${Math.min(w, 600)}px;height:${Math.max(Math.min(h, 400), 16)}px;` +
      'border:1px dashed #9a9a9a;color:#7a7a7a;font:8pt Arial,sans-serif;overflow:hidden;vertical-align:bottom;box-sizing:border-box;padding:2px');
    span.textContent = '[' + (texto || 'imagen externa no incluida') + ']';
    img.replaceWith(span);
  }

  // ------------------------------------------------------------------ marco de maquetación

  async function crearMarco(contenidoHtml, anchoPx) {
    const marco = document.createElement('iframe');
    marco.setAttribute('sandbox', 'allow-same-origin');
    marco.setAttribute('aria-hidden', 'true');
    marco.style.cssText = `position:absolute;left:-30000px;top:0;width:${anchoPx}px;height:1px;border:0;`;
    const listo = new Promise(r => marco.addEventListener('load', r, { once: true }));
    marco.srcdoc = contenidoHtml;
    document.body.appendChild(marco);
    await listo;
    return marco;
  }

  function documentoMarco(lim, anchoPx) {
    const csp = "default-src 'none'; img-src blob: data:; font-src blob: data:; style-src 'unsafe-inline'";
    return `<!doctype html><html><head><meta charset="utf-8">
<meta http-equiv="Content-Security-Policy" content="${csp}">
<style>${lim.estilos}</style>
<style id="cpdf-base">
html,body{margin:0!important;padding:0!important;background:#fff!important}
body{width:${anchoPx}px;font-family:Calibri,sans-serif;font-size:11pt;word-wrap:break-word;overflow-wrap:break-word}
*{font-kerning:none!important;font-variant-ligatures:none!important;font-feature-settings:"liga" 0,"clig" 0,"kern" 0!important;
  text-rendering:optimizeSpeed!important;-webkit-font-smoothing:auto!important;font-synthesis:none!important}
img{max-width:100%!important}
</style></head><body ${lim.atributosBody}>${lim.cuerpo}</body></html>`;
  }

  // Sustituye las viñetas automáticas (::marker) por texto real medible
  function prepararListas(doc, win) {
    const romano = n => { const v = [[1000, 'm'], [900, 'cm'], [500, 'd'], [400, 'cd'], [100, 'c'], [90, 'xc'], [50, 'l'], [40, 'xl'], [10, 'x'], [9, 'ix'], [5, 'v'], [4, 'iv'], [1, 'i']]; let s = ''; v.forEach(([k, r]) => { while (n >= k) { s += r; n -= k; } }); return s; };
    const alfa = n => { let s = ''; while (n > 0) { n--; s = String.fromCharCode(97 + (n % 26)) + s; n = Math.floor(n / 26); } return s; };
    doc.querySelectorAll('li').forEach(li => {
      const cs = win.getComputedStyle(li);
      if (cs.display !== 'list-item' || cs.listStyleType === 'none') return;
      const lista = li.parentElement;
      let n = 1;
      if (lista && lista.tagName === 'OL') {
        const inicio = parseInt(lista.getAttribute('start') || '1', 10) || 1;
        const hermanos = Array.from(lista.children).filter(x => x.tagName === 'LI');
        n = inicio + hermanos.indexOf(li);
        if (li.hasAttribute('value')) n = parseInt(li.getAttribute('value'), 10) || n;
      }
      const t = cs.listStyleType;
      const marca = t === 'disc' ? '•' : t === 'circle' ? '◦' : t === 'square' ? '▪' :
        t === 'decimal' ? n + '.' : t === 'decimal-leading-zero' ? String(n).padStart(2, '0') + '.' :
        t === 'lower-alpha' || t === 'lower-latin' ? alfa(n) + '.' : t === 'upper-alpha' || t === 'upper-latin' ? alfa(n).toUpperCase() + '.' :
        t === 'lower-roman' ? romano(n) + '.' : t === 'upper-roman' ? romano(n).toUpperCase() + '.' : '•';
      if (cs.position === 'static') li.style.position = 'relative';
      li.style.listStyleType = 'none';
      const span = doc.createElement('span');
      span.setAttribute('data-cpdf-marca', '1');
      span.style.cssText = 'position:absolute;right:100%;margin-right:0.45em;white-space:nowrap;text-decoration:none;font-weight:inherit';
      span.textContent = marca;
      li.insertBefore(span, li.firstChild);
    });
  }

  // Asigna a cada elemento su familia libre equivalente; devuelve las familias usadas
  function mapearFuentes(doc, win) {
    const usadas = new Set(['dejavu']);
    const elementos = [doc.body, ...doc.body.querySelectorAll('*')];
    const asignacion = elementos.map(el => {
      const cs = win.getComputedStyle(el);
      const lista = cs.fontFamily;
      return { el, clave: claveFamilia(lista), primera: primeraFamilia(lista) };
    });
    asignacion.forEach(({ el, clave, primera }) => {
      usadas.add(clave);
      el.style.setProperty('font-family', `"cpdf-${clave}", "cpdf-dejavu"`, 'important');
      const tabla = SIMBOLOS[primera];
      if (tabla) {
        el.childNodes.forEach(n => {
          if (n.nodeType === 3) n.data = Array.from(n.data).map(ch => tabla[ch] || (ch.charCodeAt(0) >= 0xF000 && ch.charCodeAt(0) <= 0xF0FF ? '•' : ch)).join('');
        });
      }
    });
    return usadas;
  }

  async function cargarFuentes(doc, familias, cargarFuente) {
    const bytes = {};
    const reglas = [];
    const urls = [];
    for (const fam of familias) {
      for (const [est, peso, estilo] of [['r', 400, 'normal'], ['b', 700, 'normal'], ['i', 400, 'italic'], ['bi', 700, 'italic']]) {
        const cara = `${fam}-${est}`;
        const b = await cargarFuente(cara);
        bytes[cara] = b;
        const url = URL.createObjectURL(new Blob([b], { type: 'font/ttf' }));
        urls.push(url);
        reglas.push(`@font-face{font-family:"cpdf-${fam}";src:url(${url}) format("truetype");font-weight:${peso};font-style:${estilo};font-display:block}`);
      }
    }
    const st = doc.createElement('style');
    st.textContent = reglas.join('\n');
    doc.head.appendChild(st);
    const cargas = [];
    familias.forEach(fam => ['400 normal', '700 normal', '400 italic', '700 italic'].forEach(v => {
      const [p, e] = v.split(' ');
      cargas.push(doc.fonts.load(`${e} ${p} 16px "cpdf-${fam}"`).catch(() => null));
    }));
    await Promise.all(cargas);
    await doc.fonts.ready;
    return { bytes, urls };
  }

  // ------------------------------------------------------------------ recolección de elementos

  function caraDe(cs) {
    const fam = claveFamilia(cs.fontFamily);
    const negrita = parseInt(cs.fontWeight, 10) >= 600 || cs.fontWeight === 'bold';
    const italica = cs.fontStyle === 'italic' || cs.fontStyle === 'oblique';
    return { fam, cara: `${fam}-${negrita ? (italica ? 'bi' : 'b') : (italica ? 'i' : 'r')}`, negrita, italica };
  }

  function transformar(txt, t) {
    if (t === 'uppercase') return txt.toUpperCase();
    if (t === 'lowercase') return txt.toLowerCase();
    if (t === 'capitalize') return txt.replace(/(^|\s)(\S)/g, (m, a, b) => a + b.toUpperCase());
    return txt;
  }

  /**
   * Maqueta un HTML y devuelve sus elementos dibujables.
   * @param {string} html
   * @param {object} o  { anchoPx, candidatos:[{clave,nombre,contentId,url,bytes,formato,ancho,alto}],
   *                       cargarFuente(cara)->Uint8Array, esTexto }
   */
  async function renderizar(html, o) {
    const diag = [];
    const lim = limpiarHtml(o.esTexto ? textoAHtml(html) : html, o.candidatos || [], diag);
    const marco = await crearMarco(documentoMarco(lim, o.anchoPx), o.anchoPx);
    const win = marco.contentWindow, doc = marco.contentDocument;
    let fuentes;
    try {
      prepararListas(doc, win);
      const familias = mapearFuentes(doc, win);
      fuentes = await cargarFuentes(doc, familias, o.cargarFuente);
      // esperar imágenes
      await Promise.all(Array.from(doc.images).map(im => im.complete ? null : new Promise(r => { im.onload = im.onerror = r; })));
      await cuadro(win); await cuadro(win);

      const anchoReal = Math.max(doc.documentElement.scrollWidth, doc.body.scrollWidth, o.anchoPx);
      const escala = Math.min(1, o.anchoPx / anchoReal);
      const resultado = recolectar(doc, win, o, fuentes.bytes);
      resultado.escala = escala;
      // Alto real del contenido (el marco mide 1 px, así que no infla la medida)
      const finItems = resultado.items.reduce((mx, it) => it.t === 'enlace' ? mx : Math.max(mx, it.inf), 0);
      resultado.altoPx = Math.max(finItems, Math.min(doc.documentElement.scrollHeight, doc.body.scrollHeight + 2));
      resultado.usadas = lim.usadas;
      resultado.externas = lim.externas;
      resultado.fuentes = fuentes.bytes;
      resultado.diagnostico = diag;
      return resultado;
    } finally {
      if (fuentes) fuentes.urls.forEach(u => URL.revokeObjectURL(u));
      marco.remove();
    }
  }

  function recolectar(doc, win, o, bytesFuentes) {
    const items = [];
    const imagenes = new Map();
    const rango = doc.createRange();
    const cacheEstilo = new Map();
    const cacheBase = new Map();
    const cacheClip = new Map();
    const cacheOpacidad = new Map();
    const cacheDeco = new Map();
    const cacheTTF = new Map();
    const glifos = {
      tiene(cara, cp) {
        if (!bytesFuentes[cara]) return false;
        let d = cacheTTF.get(cara);
        if (!d) { d = root.CorreoFuentes.leerTTF(bytesFuentes[cara]); cacheTTF.set(cara, d); }
        return d.cmap.has(cp);
      }
    };
    const est = el => { let c = cacheEstilo.get(el); if (!c) { c = win.getComputedStyle(el); cacheEstilo.set(el, c); } return c; };

    function opacidad(el) {
      if (!el || el === doc.documentElement) return 1;
      if (cacheOpacidad.has(el)) return cacheOpacidad.get(el);
      const v = parseFloat(est(el).opacity) * opacidad(el.parentElement);
      cacheOpacidad.set(el, v);
      return v;
    }

    function recorte(el) {
      if (!el || el === doc.body || el === doc.documentElement) return null;
      if (cacheClip.has(el)) return cacheClip.get(el);
      let clip = recorte(el.parentElement);
      const cs = est(el);
      if (cs.overflowX !== 'visible' || cs.overflowY !== 'visible') {
        const r = el.getBoundingClientRect();
        const propio = { l: r.left, t: r.top, r: r.right, b: r.bottom };
        clip = clip ? { l: Math.max(clip.l, propio.l), t: Math.max(clip.t, propio.t), r: Math.min(clip.r, propio.r), b: Math.min(clip.b, propio.b) } : propio;
      }
      cacheClip.set(el, clip);
      return clip;
    }
    const fuera = (clip, r) => clip && (r.right <= clip.l + 0.5 || r.left >= clip.r - 0.5 || r.bottom <= clip.t + 0.5 || r.top >= clip.b - 0.5);

    function decoracion(el) {
      if (!el || el === doc.documentElement) return { sub: null, tach: null };
      if (cacheDeco.has(el)) return cacheDeco.get(el);
      const padre = (est(el).display === 'inline' || el.tagName === 'A') ? decoracion(el.parentElement) : { sub: null, tach: null };
      const cs = est(el);
      const linea = cs.textDecorationLine || cs.textDecoration || '';
      const c = color(cs.textDecorationColor || cs.color);
      const r = {
        sub: /underline/.test(linea) ? (c ? c.rgb : [0, 0, 0]) : padre.sub,
        tach: /line-through/.test(linea) ? (c ? c.rgb : [0, 0, 0]) : padre.tach
      };
      cacheDeco.set(el, r);
      return r;
    }

    function desplazamientoBase(fam, negrita, italica, tam) {
      const k = `${fam}|${negrita}|${italica}|${tam}`;
      if (cacheBase.has(k)) return cacheBase.get(k);
      const s = doc.createElement('span');
      s.style.cssText = `position:absolute;left:0;top:0;white-space:nowrap;line-height:normal;font-family:"cpdf-${fam}" !important;` +
        `font-size:${tam}px;font-weight:${negrita ? 700 : 400};font-style:${italica ? 'italic' : 'normal'}`;
      s.appendChild(doc.createTextNode('Hg'));
      const m = doc.createElement('span');
      m.style.cssText = 'display:inline-block;width:0;height:0;vertical-align:baseline';
      s.appendChild(m);
      doc.body.appendChild(s);
      const rr = doc.createRange(); rr.selectNodeContents(s.firstChild);
      const v = m.getBoundingClientRect().top - rr.getBoundingClientRect().top;
      s.remove();
      cacheBase.set(k, v);
      return v;
    }

    function caraPara(cp, base) {
      if (glifos.tiene(base.cara, cp)) return base.cara;
      const dj = 'dejavu-' + base.cara.split('-')[1];
      if (glifos.tiene(dj, cp)) return dj;
      return null; // se dibuja como imagen (p. ej. emoji)
    }

    function textoNodo(nodo) {
      const el = nodo.parentElement;
      if (!el) return;
      const cs = est(el);
      if (cs.visibility !== 'visible' || opacidad(el) < 0.05) return;
      const col = color(cs.color);
      if (!col || col.a < 0.05) return;
      const tam = parseFloat(cs.fontSize);
      if (!(tam >= 1)) return;
      const base = caraDe(cs);
      const deco = decoracion(el);
      const clip = recorte(el);
      const texto = nodo.data;
      const re = /\S+/g;
      let m;
      while ((m = re.exec(texto)) !== null) {
        const ini = m.index, fin = ini + m[0].length;
        rango.setStart(nodo, ini); rango.setEnd(nodo, fin);
        const rects = rango.getClientRects();
        if (!rects.length) continue;
        const espacio = /\s/.test(texto.charAt(fin)) ? ' ' : '';
        // Dividir en trozos por línea (si la palabra se partió) y por fuente
        const chars = Array.from(m[0]);
        let pos = ini;
        let trozo = null;
        const trozos = [];
        chars.forEach(ch => {
          const cp = ch.codePointAt(0);
          const cara = caraPara(cp, base);
          const largo = ch.length;
          let lineaTop = null;
          if (rects.length > 1) {
            rango.setStart(nodo, pos); rango.setEnd(nodo, pos + largo);
            const r = rango.getBoundingClientRect();
            lineaTop = Math.round(r.top);
          }
          if (!trozo || trozo.cara !== cara || trozo.linea !== lineaTop || cara === null) {
            trozo = { cara, linea: lineaTop, ini: pos, fin: pos, txt: '' };
            trozos.push(trozo);
          }
          trozo.txt += ch; trozo.fin = pos + largo;
          pos += largo;
        });
        trozos.forEach((t, idx) => {
          rango.setStart(nodo, t.ini); rango.setEnd(nodo, t.fin);
          const r = rango.getBoundingClientRect();
          if (!r.width && !r.height) return;
          if (fuera(clip, r)) return;
          if (t.cara === null) {
            items.push({ t: 'glifo', x: r.left, y: r.top, w: r.width, h: r.height, txt: t.txt, tam, color: col.rgb, sup: r.top, inf: r.bottom });
            return;
          }
          const fam = t.cara.split('-')[0];
          const y = r.top + desplazamientoBase(fam, /b/.test(t.cara.split('-')[1]), /i/.test(t.cara.split('-')[1]), tam);
          const txt = transformar(t.txt, cs.textTransform) + (idx === trozos.length - 1 ? espacio : '');
          items.push({ t: 'texto', x: r.left, y, tam, cara: t.cara, color: col.rgb, txt, ancho: r.width,
            sub: deco.sub, tach: deco.tach, sup: r.top, inf: r.bottom });
        });
      }
    }

    function elemento(el) {
      const cs = est(el);
      if (cs.display === 'none') return;
      const visible = cs.visibility === 'visible' && opacidad(el) >= 0.05;
      const clip = recorte(el.parentElement);
      const tag = el.tagName;

      if (visible && el !== doc.body && el !== doc.documentElement) {
        // Fondo
        const fondo = color(cs.backgroundColor);
        if (fondo && fondo.a > 0.05 && !(fondo.rgb[0] > 0.99 && fondo.rgb[1] > 0.99 && fondo.rgb[2] > 0.99)) {
          const rects = cs.display === 'inline' ? Array.from(el.getClientRects()) : [el.getBoundingClientRect()];
          rects.forEach(r => { if (r.width && r.height && !fuera(clip, r)) items.push({ t: 'rect', x: r.left, y: r.top, w: r.width, h: r.height, color: fondo.rgb, sup: r.top, inf: r.bottom }); });
        }
        // Bordes
        if (cs.display !== 'inline') {
          const r = el.getBoundingClientRect();
          if (r.width || r.height) {
            [['Top', 'h'], ['Bottom', 'h'], ['Left', 'v'], ['Right', 'v']].forEach(([lado]) => {
              const w = parseFloat(cs['border' + lado + 'Width']);
              const estilo = cs['border' + lado + 'Style'];
              const c = color(cs['border' + lado + 'Color']);
              if (!(w > 0) || estilo === 'none' || estilo === 'hidden' || !c || c.a < 0.05) return;
              let x = r.left, y = r.top, ancho = r.width, alto = r.height;
              if (lado === 'Top') alto = w; else if (lado === 'Bottom') { y = r.bottom - w; alto = w; }
              else if (lado === 'Left') ancho = w; else { x = r.right - w; ancho = w; }
              const cr = { left: x, top: y, right: x + ancho, bottom: y + alto };
              if (fuera(clip, cr)) return;
              items.push({ t: 'rect', x, y, w: ancho, h: alto, color: c.rgb, guiones: estilo === 'dashed' ? 3 : estilo === 'dotted' ? 1 : 0, sup: y, inf: y + alto });
            });
          }
        }
        // Imagen
        if (tag === 'IMG' && el.complete && el.naturalWidth > 0) {
          const r = el.getBoundingClientRect();
          const bl = parseFloat(cs.borderLeftWidth) + parseFloat(cs.paddingLeft);
          const bt = parseFloat(cs.borderTopWidth) + parseFloat(cs.paddingTop);
          const br = parseFloat(cs.borderRightWidth) + parseFloat(cs.paddingRight);
          const bb = parseFloat(cs.borderBottomWidth) + parseFloat(cs.paddingBottom);
          const caja = { left: r.left + bl, top: r.top + bt, right: r.right - br, bottom: r.bottom - bb };
          if (caja.right > caja.left && caja.bottom > caja.top && !fuera(clip, caja)) {
            const clave = el.getAttribute('data-cpdf-img') || ('img-' + items.length);
            if (!imagenes.has(clave)) imagenes.set(clave, extraerImagen(el, clave, o.candidatos));
            items.push({ t: 'img', x: caja.left, y: caja.top, w: caja.right - caja.left, h: caja.bottom - caja.top, clave, sup: caja.top, inf: caja.bottom });
          }
        }
      }
      // Enlaces
      if (tag === 'A') {
        const href = el.getAttribute('href') || '';
        if (/^(https?:|mailto:)/i.test(href)) {
          Array.from(el.getClientRects()).forEach(r => {
            if (r.width > 0 && r.height > 0) items.push({ t: 'enlace', x: r.left, y: r.top, w: r.width, h: r.height, url: href, sup: r.top, inf: r.bottom });
          });
        }
      }
      if (tag === 'IMG' || tag === 'SVG' || tag === 'svg') return;
      el.childNodes.forEach(n => {
        if (n.nodeType === 1) elemento(n);
        else if (n.nodeType === 3 && n.data.trim()) textoNodo(n);
      });
    }

    elemento(doc.body);
    // Caracteres sin fuente (emoji, símbolos raros): convertirlos en pequeñas imágenes
    items.forEach(it => {
      if (it.t !== 'glifo') return;
      const clave = 'g:' + it.txt + '|' + it.tam + '|' + it.color.join(',');
      if (!imagenes.has(clave)) imagenes.set(clave, { bytes: glifoComoImagen(it.txt, it.tam, it.color), formato: 'png' });
      it.t = 'img'; it.clave = clave;
    });
    return { items, imagenes };
  }

  function extraerImagen(img, clave, candidatos) {
    const c = (candidatos || []).find(x => x.clave === clave);
    if (c && (c.formato === 'png' || (c.formato === 'jpg' && c.orientacion === 1))) {
      return { bytes: c.bytes, formato: c.formato };
    }
    // Otros formatos (GIF, BMP, WEBP, JPG rotado, data:): pasar por lienzo
    const lienzo = document.createElement('canvas');
    lienzo.width = img.naturalWidth; lienzo.height = img.naturalHeight;
    const ctx = lienzo.getContext('2d');
    ctx.drawImage(img, 0, 0);
    const datos = lienzo.toDataURL('image/png');
    return { bytes: base64Bytes(datos.split(',')[1]), formato: 'png' };
  }

  function base64Bytes(b64) {
    const bin = atob(b64);
    const u = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) u[i] = bin.charCodeAt(i);
    return u;
  }

  /** Dibuja un carácter sin fuente disponible (p. ej. emoji) como imagen PNG. */
  function glifoComoImagen(txt, tamPx, rgb) {
    const escala = 4;
    const lienzo = document.createElement('canvas');
    const ctx = lienzo.getContext('2d');
    const fuente = `${tamPx * escala}px "Segoe UI Emoji","Apple Color Emoji","Noto Color Emoji","Segoe UI Symbol",sans-serif`;
    ctx.font = fuente;
    const w = Math.ceil(ctx.measureText(txt).width) + 2;
    lienzo.width = Math.max(w, 1); lienzo.height = Math.ceil(tamPx * escala * 1.3);
    ctx.font = fuente;
    ctx.textBaseline = 'alphabetic';
    ctx.fillStyle = `rgb(${rgb.map(v => Math.round(v * 255)).join(',')})`;
    ctx.fillText(txt, 1, Math.round(tamPx * escala * 1.02));
    return base64Bytes(lienzo.toDataURL('image/png').split(',')[1]);
  }

  /** Lee ancho y alto de una imagen (para emparejar referencias por proporción). */
  async function dimensiones(bytes, tipo) {
    try {
      const bmp = await createImageBitmap(new Blob([bytes], { type: tipo || 'image/png' }));
      const d = { ancho: bmp.width, alto: bmp.height };
      bmp.close && bmp.close();
      return d;
    } catch (e) { return { ancho: 1, alto: 1 }; }
  }

  root.CorreoRender = { renderizar, glifoComoImagen, dimensiones, escaparHtml, MAPA_FUENTES };
})(typeof self !== 'undefined' ? self : this);
