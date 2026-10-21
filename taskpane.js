/* Correo a PDF — panel del complemento de Outlook (prototipo v0.3.0) */
(function () {
  'use strict';

  const $ = id => document.getElementById(id);
  let urlPdf = null;
  let ultimoDiagnostico = null;
  const cacheFuentes = new Map();

  Office.onReady(info => {
    $('version').textContent = 'v' + CorreoPDF.VERSION;
    if (info.host !== Office.HostType.Outlook) {
      mostrarError('Este complemento solo funciona dentro de Outlook.');
      return;
    }
    $('generar').addEventListener('click', generar);
    $('abrir').addEventListener('click', () => { if (urlPdf) window.open(urlPdf, '_blank'); });
    $('diagnostico').addEventListener('click', descargarDiagnostico);
    if (Office.context.mailbox.addHandlerAsync) {
      Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, cargarCorreo);
    }
    cargarCorreo();
  });

  // ------------------------------------------------------------------ vista previa

  function cargarCorreo() {
    limpiarResultado();
    const item = Office.context.mailbox.item;
    if (!item) {
      $('sin-correo').classList.remove('oculto');
      ['correo', 'anexos-caja'].forEach(id => $(id).classList.add('oculto'));
      $('generar').disabled = true;
      return;
    }
    $('sin-correo').classList.add('oculto');
    $('correo').classList.remove('oculto');
    $('asunto').textContent = item.subject || '(sin asunto)';
    $('remitente').textContent = 'De: ' + persona(item.from) + ' · ' + CorreoPDF.fechaOutlook(item.dateTimeCreated, 'America/Bogota');

    const adjuntos = item.attachments || [];
    const lista = $('lista-anexos');
    lista.innerHTML = '';
    $('anexos-caja').classList.remove('oculto');
    $('conteo').textContent = '(' + adjuntos.length + ')';
    let noConvertibles = 0;
    if (!adjuntos.length) agregarFila(lista, 'info', 'i', 'El correo no trae adjuntos.', 'Se generará el PDF del correo con sus enlaces.');
    adjuntos.forEach(a => {
      const tam = a.size ? ' · ' + CorreoPDF.tamanoLegible(a.size) : '';
      if (a.attachmentType === 'cloud') return agregarFila(lista, 'info', '↗', a.name, 'Archivo en la nube: se incluirá como enlace' + tam);
      if (a.attachmentType === 'item') { noConvertibles++; return agregarFila(lista, 'no', '!', a.name, 'Correo adjunto: no se convierte en esta versión' + tam); }
      const p = CorreoPDF.prediccionPorNombre(a.name);
      if (p.descripcion === 'Imagen') return agregarFila(lista, 'ok', '✓', a.name, 'Imagen: si aparece en el cuerpo se muestra en su lugar; si no, se une como anexo' + tam);
      if (p.soportado) return agregarFila(lista, 'ok', '✓', a.name, p.descripcion + ' · se unirá al PDF' + tam);
      noConvertibles++;
      agregarFila(lista, 'no', '!', a.name, p.descripcion + ' · NO se puede convertir a PDF' + tam);
    });
    const alerta = $('alerta-previa');
    if (noConvertibles) {
      alerta.textContent = `Atención: ${noConvertibles} archivo(s) no son PDF ni imagen y no se pueden convertir. ` +
        'Quedarán registrados como "no convertidos" en la hoja de control. Gestiónelos por separado.';
      alerta.classList.remove('oculto');
    } else alerta.classList.add('oculto');
    $('generar').disabled = false;
  }

  function agregarFila(lista, clase, icono, nombre, detalle) {
    const li = document.createElement('li');
    li.className = clase;
    const i = document.createElement('span'); i.className = 'icono'; i.textContent = icono;
    const t = document.createElement('span');
    const n = document.createElement('span'); n.className = 'nombre'; n.textContent = nombre;
    const d = document.createElement('span'); d.className = 'detalle'; d.textContent = detalle;
    t.appendChild(n); t.appendChild(d); li.appendChild(i); li.appendChild(t);
    lista.appendChild(li);
  }

  // ------------------------------------------------------------------ lectura del correo

  async function obtenerCorreo(item) {
    const diag = { origen: null, errores: [] };
    let html = null, texto = null, partes = [];

    // 1. Mensaje original completo (MIME): HTML original y adjuntos con su Content-ID
    if (typeof item.getAsFileAsync === 'function') {
      try {
        progreso('Leyendo el mensaje original…');
        const b64 = await promesa(cb => item.getAsFileAsync(cb));
        const r = CorreoMime.analizar(base64ABytes(b64));
        html = r.html; texto = r.texto;
        partes = r.partes.map(p => ({ nombre: p.nombre, tipo: p.tipo, bytes: p.bytes, contentId: p.contentId, enLinea: p.enLinea, esCorreo: p.esCorreo }));
        diag.origen = 'eml';
      } catch (e) { diag.errores.push('getAsFileAsync: ' + (e.message || e)); }
    }
    // 2. Respaldo: cuerpo y adjuntos por separado
    if (!diag.origen) {
      diag.origen = 'api';
      html = await cuerpo(item, Office.CoercionType.Html).catch(() => null);
      texto = await cuerpo(item, Office.CoercionType.Text).catch(() => '');
      const adjuntos = (item.attachments || []).filter(a => a.attachmentType !== 'cloud');
      for (let i = 0; i < adjuntos.length; i++) {
        const a = adjuntos[i];
        progreso(`Leyendo adjunto ${i + 1} de ${adjuntos.length}: ${a.name}`);
        const base = { nombre: a.name, tipo: a.contentType || '', contentId: a.contentId || '', enLinea: !!a.isInline, idOutlook: a.id, esCorreo: a.attachmentType === 'item' };
        if (base.esCorreo) { partes.push(Object.assign(base, { bytes: null })); continue; }
        try {
          const c = await promesa(cb => item.getAttachmentContentAsync(a.id, cb));
          partes.push(Object.assign(base, { bytes: c.format === Office.MailboxEnums.AttachmentContentFormat.Base64 ? base64ABytes(c.content) : null }));
        } catch (e) { partes.push(Object.assign(base, { bytes: null, error: e.message || String(e) })); }
      }
    }
    // Archivos en la nube (no vienen en el MIME)
    const nube = [];
    for (const a of (item.attachments || []).filter(x => x.attachmentType === 'cloud')) {
      try {
        const c = await promesa(cb => item.getAttachmentContentAsync(a.id, cb));
        nube.push({ nombre: a.name, url: c.content });
      } catch (e) { nube.push({ nombre: a.name, url: '' }); }
    }
    // Para buscar radicados se usa el texto visible del HTML (la parte de texto plano puede ser un resumen)
    if (html) texto = new DOMParser().parseFromString(html, 'text/html').body.textContent || texto || '';
    return { html, texto: texto || '', partes, nube, diag };
  }

  // ------------------------------------------------------------------ generación

  let ultimaEtapa = '';
  async function generar() {
    const item = Office.context.mailbox.item;
    if (!item) return;
    limpiarResultado();
    $('generar').disabled = true;
    ultimoDiagnostico = { version: CorreoPDF.VERSION, fecha: new Date().toISOString(), etapas: [] };
    let vigilante;
    const tiempoMaximo = new Promise((ok, mal) => {
      vigilante = setTimeout(() => mal(new Error('El proceso tardó más de 2 minutos en la etapa: "' + ultimaEtapa + '".')), 120000);
    });
    try {
      await Promise.race([procesar(item), tiempoMaximo]);
    } catch (e) {
      console.error(e);
      ultimoDiagnostico.error = String(e && e.message || e);
      ultimoDiagnostico.etapaFallida = ultimaEtapa;
      mostrarError('No se pudo generar el PDF: ' + (e.message || e));
    } finally {
      clearTimeout(vigilante);
      $('progreso').classList.add('oculto');
      $('generar').disabled = false;
    }
  }

  async function procesar(item) {
    const urlsTemporales = [];
    try {
      const correo = await obtenerCorreo(item);
      const meta = {
        asunto: item.subject, de: persona(item.from),
        para: (item.to || []).map(persona), cc: (item.cc || []).map(persona),
        recibido: item.dateTimeCreated, idMensaje: item.internetMessageId
      };

      // Imágenes candidatas a estar dentro del cuerpo
      progreso('Preparando imágenes del cuerpo…');
      const candidatos = [];
      for (let i = 0; i < correo.partes.length; i++) {
        const p = correo.partes[i];
        if (!p.bytes || p.esCorreo) continue;
        const tipo = CorreoPDF.detectarTipo(p.bytes, p.nombre);
        if (tipo.categoria !== 'imagen' && tipo.categoria !== 'imagen-convertible') continue;
        const mime = { png: 'image/png', jpg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', webp: 'image/webp' }[tipo.tipo];
        const url = URL.createObjectURL(new Blob([p.bytes], { type: mime }));
        urlsTemporales.push(url);
        const dim = await CorreoRender.dimensiones(p.bytes, mime);
        candidatos.push({
          clave: 'p' + i, nombre: p.nombre, contentId: p.contentId, idOutlook: p.idOutlook, bytes: p.bytes,
          formato: tipo.tipo, orientacion: tipo.tipo === 'jpg' ? CorreoPDF.orientacionExif(p.bytes) : 1,
          url, ancho: dim.ancho, alto: dim.alto
        });
      }

      progreso('Maquetando el correo…');
      const comunes = { anchoPx: CorreoPDF.ANCHO_PX, cargarFuente, alProgresar: progreso };
      const segCuerpo = await CorreoRender.renderizar(correo.html || correo.texto, Object.assign({ candidatos, esTexto: !correo.html }, comunes));
      // El encabezado se arma después del cuerpo: así la línea de adjuntos excluye las imágenes que van dentro del cuerpo
      const adjuntosVisibles = listaAdjuntos(item, correo.partes, segCuerpo.usadas);
      const segEncabezado = await CorreoRender.renderizar(CorreoPDF.htmlEncabezado(meta, 'America/Bogota', adjuntosVisibles), Object.assign({ candidatos: [] }, comunes));

      const anexos = correo.partes.map((p, i) => ({
        nombre: p.nombre, tamano: p.bytes ? p.bytes.length : null, tipo: p.esCorreo ? 'item' : 'archivo',
        bytes: p.bytes, enLinea: p.enLinea, enCuerpo: segCuerpo.usadas.has('p' + i), error: p.error
      })).concat(correo.nube.map(n => ({ nombre: n.nombre, tipo: 'nube', url: n.url })));

      const resultado = await CorreoPDF.generarPdf({
        meta,
        segmentos: [
          Object.assign({ anchoPx: CorreoPDF.ANCHO_PX }, segEncabezado),
          Object.assign({ anchoPx: CorreoPDF.ANCHO_PX }, segCuerpo)
        ],
        imagenes: new Map([...segEncabezado.imagenes, ...segCuerpo.imagenes]),
        fuentes: Object.assign({}, segEncabezado.fuentes, segCuerpo.fuentes),
        anexos,
        enlaces: enlacesDeHtml(correo.html),
        cuerpoTexto: correo.texto,
        externas: segCuerpo.externas,
        origen: correo.diag.origen
      }, { nombreDespacho: nombreBuzon(), nombreCuenta: nombreCuenta(), enlaceCorreo: enlaceCorreo(item), alProgresar: progreso });

      ultimoDiagnostico = {
        version: CorreoPDF.VERSION, fecha: new Date().toISOString(), etapas: ultimoDiagnostico.etapas,
        cliente: Office.context.diagnostics ? { host: Office.context.diagnostics.hostName, version: Office.context.diagnostics.hostVersion, plataforma: Office.context.diagnostics.platform } : null,
        origen: correo.diag.origen, errores: correo.diag.errores,
        htmlCaracteres: correo.html ? correo.html.length : 0,
        imagenesCuerpo: segCuerpo.diagnostico,
        partes: correo.partes.map((p, i) => ({ nombre: p.nombre, tipo: p.tipo, bytes: p.bytes ? p.bytes.length : 0, conContentId: !!p.contentId, enLinea: p.enLinea, usadaEnCuerpo: segCuerpo.usadas.has('p' + i) })),
        paginasCorreo: resultado.paginasCorreo, totalPaginas: resultado.totalPaginas
      };
      mostrarResultado(resultado);
    } finally {
      urlsTemporales.forEach(u => URL.revokeObjectURL(u));
    }
  }

  async function cargarFuente(cara) {
    if (!cacheFuentes.has(cara)) {
      cacheFuentes.set(cara, fetch(`fuentes/${cara}.ttf`).then(r => {
        if (!r.ok) throw new Error(`No se pudo descargar la fuente "${cara}" (respuesta ${r.status}). Verifique que la carpeta fuentes/ esté publicada junto al complemento.`);
        return r.arrayBuffer();
      }).then(b => new Uint8Array(b)));
    }
    return cacheFuentes.get(cara);
  }

  function mostrarResultado(r) {
    const blob = new Blob([r.bytes], { type: 'application/pdf' });
    urlPdf = URL.createObjectURL(blob);
    const enlace = $('descargar');
    enlace.href = urlPdf;
    enlace.download = r.nombreArchivo;

    const incluidos = r.anexos.filter(a => a.estado === 'incluido').length;
    const enCuerpo = r.anexos.filter(a => a.estado === 'en_cuerpo').length;
    const noConv = r.anexos.filter(a => a.estado === 'no_convertido');
    const resumen = $('resumen');
    resumen.innerHTML = '';
    [
      `✓ PDF generado: ${r.totalPaginas} páginas (correo: ${r.paginasCorreo})`,
      `Anexos unidos: ${incluidos}` + (enCuerpo ? ` · Imágenes dentro del cuerpo: ${enCuerpo}` : ''),
      `Enlaces registrados: ${r.enlaces.length}`,
      r.radicados.length ? `Radicado detectado: ${r.radicados[0]}` : 'Sin radicado detectado',
      `Archivo: ${r.nombreArchivo}`
    ].forEach(t => { const p = document.createElement('p'); p.textContent = t; resumen.appendChild(p); });

    const cajaAlertas = $('alertas-finales');
    cajaAlertas.innerHTML = '';
    if (noConv.length) {
      const div = document.createElement('div');
      div.className = 'alerta';
      const titulo = document.createElement('strong');
      titulo.textContent = `${noConv.length} archivo(s) NO incluidos en el PDF:`;
      div.appendChild(titulo);
      const ul = document.createElement('ul');
      noConv.forEach(a => { const li = document.createElement('li'); li.style.display = 'block'; li.textContent = `${a.nombre} — ${a.motivo}`; ul.appendChild(li); });
      div.appendChild(ul);
      cajaAlertas.appendChild(div);
    }
    $('huella').textContent = r.sha256;
    $('resultado').classList.remove('oculto');
    try { enlace.click(); } catch (e) { /* el usuario usa el botón */ }
  }

  function descargarDiagnostico() {
    if (!ultimoDiagnostico) return;
    const blob = new Blob([JSON.stringify(ultimoDiagnostico, null, 2)], { type: 'application/json' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = 'diagnostico-correo-a-pdf.json';
    document.body.appendChild(a); a.click(); a.remove();
  }

  // ------------------------------------------------------------------ utilidades Office

  function promesa(llamar) {
    return new Promise((ok, mal) => llamar(r => r.status === Office.AsyncResultStatus.Succeeded
      ? ok(r.value) : mal(new Error(r.error && r.error.message || 'error desconocido'))));
  }
  function cuerpo(item, tipo) { return promesa(cb => item.body.getAsync(tipo, cb)); }

  function persona(p) {
    if (!p) return '—';
    if (p.displayName && p.emailAddress && p.displayName !== p.emailAddress) return `${p.displayName} <${p.emailAddress}>`;
    return p.emailAddress || p.displayName || '—';
  }

  /** Adjuntos tal como los lista Outlook (nombre y tamaño reales), sin las imágenes mostradas dentro del cuerpo. */
  function listaAdjuntos(item, partes, usadas) {
    const lista = (item.attachments || []).map(a => ({ nombre: a.name, tamano: a.size || 0, enLinea: !!a.isInline, tipo: a.attachmentType }));
    partes.forEach((p, i) => {
      if (!usadas.has('p' + i)) return;
      // quitar una aparición con el mismo nombre (preferir la marcada como incrustada)
      let k = lista.findIndex(a => a.nombre === p.nombre && a.enLinea);
      if (k < 0) k = lista.findIndex(a => a.nombre === p.nombre && a.tipo !== 'item' && a.tipo !== 'cloud');
      if (k >= 0) lista.splice(k, 1);
    });
    return lista;
  }

  function nombreCuenta() {
    try { return Office.context.mailbox.userProfile.displayName || ''; } catch (e) { return ''; }
  }

  /** Enlace permanente al mensaje en Outlook web, con el identificador real del correo. */
  function enlaceCorreo(item) {
    if (!item.itemId) return '';
    let origen = 'https://outlook.office365.com';
    try {
      const ref = document.referrer ? new URL(document.referrer) : null;
      if (ref && /(^|\.)outlook\.(office|office365|live|cloud)\.(com|microsoft)$/i.test(ref.hostname)) origen = ref.origin;
      else if (Office.context.mailbox.restUrl) origen = new URL(Office.context.mailbox.restUrl).origin;
    } catch (e) { /* usar el valor por defecto */ }
    return `${origen}/owa/?ItemID=${encodeURIComponent(item.itemId)}&exvsurl=1&viewmodel=ReadMessageItem`;
  }

  function nombreBuzon() {
    try {
      const u = Office.context.mailbox.userProfile;
      return u.displayName ? `${u.displayName} (${u.emailAddress})` : 'Secretaría';
    } catch (e) { return 'Secretaría'; }
  }

  function enlacesDeHtml(html) {
    if (!html) return [];
    try {
      const doc = new DOMParser().parseFromString(html, 'text/html');
      return Array.from(doc.querySelectorAll('a[href]')).map(a => {
        const url = a.getAttribute('href') || '';
        const texto = (a.textContent || '').trim().replace(/\s+/g, ' ');
        const etiqueta = texto && texto !== url && !/^https?:\/\//i.test(texto) ? texto.slice(0, 120) : '';
        return { url, etiqueta };
      }).filter(e => /^https?:\/\//i.test(e.url));
    } catch (e) { return []; }
  }

  function base64ABytes(b64) {
    const bin = atob(b64);
    const bytes = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
    return bytes;
  }

  // ------------------------------------------------------------------ interfaz

  function progreso(msg) {
    const p = $('progreso'); p.textContent = msg; p.classList.remove('oculto');
    ultimaEtapa = msg;
    if (ultimoDiagnostico && ultimoDiagnostico.etapas) ultimoDiagnostico.etapas.push(Math.round(performance.now()) + ' ms · ' + msg);
  }

  function limpiarResultado() {
    if (urlPdf) { URL.revokeObjectURL(urlPdf); urlPdf = null; }
    $('resultado').classList.add('oculto');
    $('error').classList.add('oculto');
  }

  function mostrarError(msg) {
    const e = $('error');
    e.textContent = msg;
    const b = document.createElement('button');
    b.className = 'enlace-sutil';
    b.textContent = 'Descargar diagnóstico técnico (sin contenido del correo)';
    b.addEventListener('click', descargarDiagnostico);
    e.appendChild(document.createElement('br'));
    e.appendChild(b);
    e.classList.remove('oculto');
  }
})();
