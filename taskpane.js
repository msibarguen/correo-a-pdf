/* Correo a PDF — panel del complemento de Outlook (prototipo v0.1.0) */
(function () {
  'use strict';

  const $ = id => document.getElementById(id);
  let urlPdf = null;

  Office.onReady(info => {
    $('version').textContent = 'v' + CorreoPDF.VERSION;
    if (info.host !== Office.HostType.Outlook) {
      mostrarError('Este complemento solo funciona dentro de Outlook.');
      return;
    }
    $('generar').addEventListener('click', generar);
    $('abrir').addEventListener('click', () => { if (urlPdf) window.open(urlPdf, '_blank'); });
    // Panel anclado: actualizar cuando el usuario cambia de correo
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
    $('remitente').textContent = 'De: ' + persona(item.from) + ' · ' + formatoFecha(item.dateTimeCreated);

    const adjuntos = item.attachments || [];
    const lista = $('lista-anexos');
    lista.innerHTML = '';
    $('anexos-caja').classList.remove('oculto');
    $('conteo').textContent = '(' + adjuntos.length + ')';

    let noConvertibles = 0;
    if (!adjuntos.length) {
      const li = document.createElement('li');
      li.className = 'info';
      li.innerHTML = '<span class="icono">i</span><span>El correo no trae adjuntos. Se generará el PDF del cuerpo y sus enlaces.</span>';
      lista.appendChild(li);
    }
    adjuntos.forEach(a => {
      let clase, icono, detalle;
      if (a.attachmentType === 'cloud') {
        clase = 'info'; icono = '↗'; detalle = 'Archivo en la nube: se incluirá como enlace';
      } else if (a.attachmentType === 'item') {
        clase = 'no'; icono = '!'; detalle = 'Correo adjunto: no se convierte en esta versión';
        noConvertibles++;
      } else {
        const p = CorreoPDF.prediccionPorNombre(a.name);
        if (a.isInline) {
          clase = 'info'; icono = '·'; detalle = 'Imagen incrustada en el cuerpo';
        } else if (p.soportado) {
          clase = 'ok'; icono = '✓'; detalle = p.descripcion + ' · se unirá al PDF';
        } else {
          clase = 'no'; icono = '!'; detalle = p.descripcion + ' · NO se puede convertir a PDF';
          noConvertibles++;
        }
      }
      const li = document.createElement('li');
      li.className = clase;
      const i = document.createElement('span'); i.className = 'icono'; i.textContent = icono;
      const t = document.createElement('span');
      const n = document.createElement('span'); n.className = 'nombre'; n.textContent = a.name;
      const d = document.createElement('span'); d.className = 'detalle';
      d.textContent = detalle + (a.size ? ' · ' + CorreoPDF.tamanoLegible(a.size) : '');
      t.appendChild(n); t.appendChild(d);
      li.appendChild(i); li.appendChild(t);
      lista.appendChild(li);
    });

    const alerta = $('alerta-previa');
    if (noConvertibles) {
      alerta.textContent = `Atención: ${noConvertibles} archivo(s) no son PDF ni imagen y no se pueden convertir. ` +
        'Quedarán registrados como "no convertidos" en la carátula. Gestiónelos por separado.';
      alerta.classList.remove('oculto');
    } else {
      alerta.classList.add('oculto');
    }
    $('generar').disabled = false;
  }

  // ------------------------------------------------------------------ generación

  async function generar() {
    const item = Office.context.mailbox.item;
    if (!item) return;
    limpiarResultado();
    $('generar').disabled = true;
    progreso('Leyendo el correo…');

    try {
      const [texto, html] = await Promise.all([
        cuerpo(item, Office.CoercionType.Text),
        cuerpo(item, Office.CoercionType.Html).catch(() => '')
      ]);

      const anexos = [];
      const adjuntos = item.attachments || [];
      for (let i = 0; i < adjuntos.length; i++) {
        const a = adjuntos[i];
        progreso(`Leyendo adjunto ${i + 1} de ${adjuntos.length}: ${a.name}`);
        const base = { nombre: a.name, tamano: a.size, enLinea: !!a.isInline };
        if (a.attachmentType === 'item') { anexos.push(Object.assign(base, { tipo: 'item' })); continue; }
        try {
          const contenido = await contenidoAdjunto(item, a.id);
          if (contenido.format === Office.MailboxEnums.AttachmentContentFormat.Url) {
            anexos.push(Object.assign(base, { tipo: 'nube', url: contenido.content }));
          } else if (contenido.format === Office.MailboxEnums.AttachmentContentFormat.Base64) {
            anexos.push(Object.assign(base, { tipo: 'archivo', bytes: base64ABytes(contenido.content) }));
          } else {
            anexos.push(Object.assign(base, { tipo: 'item' }));
          }
        } catch (e) {
          anexos.push(Object.assign(base, { tipo: a.attachmentType === 'cloud' ? 'nube' : 'archivo', error: e.message || String(e) }));
        }
      }

      const resultado = await CorreoPDF.generarPdf({
        meta: {
          asunto: item.subject,
          de: persona(item.from),
          para: (item.to || []).map(persona),
          cc: (item.cc || []).map(persona),
          recibido: item.dateTimeCreated,
          idMensaje: item.internetMessageId
        },
        cuerpoTexto: texto,
        enlaces: enlacesDeHtml(html),
        anexos
      }, {
        incluirImagenesEnLinea: $('incluir-en-linea').checked,
        convertirImagen,
        nombreDespacho: nombreBuzon(),
        alProgresar: progreso
      });

      mostrarResultado(resultado);
    } catch (e) {
      mostrarError('No se pudo generar el PDF: ' + (e.message || e));
    } finally {
      $('progreso').classList.add('oculto');
      $('generar').disabled = false;
    }
  }

  function mostrarResultado(r) {
    const blob = new Blob([r.bytes], { type: 'application/pdf' });
    urlPdf = URL.createObjectURL(blob);
    const enlace = $('descargar');
    enlace.href = urlPdf;
    enlace.download = r.nombreArchivo;

    const incluidos = r.anexos.filter(a => a.estado === 'incluido').length;
    const noConv = r.anexos.filter(a => a.estado === 'no_convertido');
    const resumen = $('resumen');
    resumen.innerHTML = '';
    [
      `✓ PDF generado: ${r.totalPaginas} páginas`,
      `Anexos unidos: ${incluidos} de ${r.anexos.length}`,
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
      noConv.forEach(a => {
        const li = document.createElement('li');
        li.style.display = 'block';
        li.textContent = `${a.nombre} — ${a.motivo}`;
        ul.appendChild(li);
      });
      div.appendChild(ul);
      cajaAlertas.appendChild(div);
    }
    $('huella').textContent = r.sha256;
    $('resultado').classList.remove('oculto');

    // Intento de descarga automática (si el entorno la bloquea, queda el botón)
    try { enlace.click(); } catch (e) { /* el usuario usa el botón */ }
  }

  // ------------------------------------------------------------------ utilidades Office

  function cuerpo(item, tipo) {
    return new Promise((ok, mal) => item.body.getAsync(tipo, r =>
      r.status === Office.AsyncResultStatus.Succeeded ? ok(r.value || '') : mal(r.error)));
  }

  function contenidoAdjunto(item, id) {
    return new Promise((ok, mal) => item.getAttachmentContentAsync(id, r =>
      r.status === Office.AsyncResultStatus.Succeeded ? ok(r.value) : mal(new Error(r.error && r.error.message || 'error desconocido'))));
  }

  function persona(p) {
    if (!p) return '—';
    if (p.displayName && p.emailAddress && p.displayName !== p.emailAddress) return `${p.displayName} <${p.emailAddress}>`;
    return p.emailAddress || p.displayName || '—';
  }

  function nombreBuzon() {
    try {
      const u = Office.context.mailbox.userProfile;
      return u.displayName ? `${u.displayName} (${u.emailAddress})` : 'Secretaría';
    } catch (e) { return 'Secretaría'; }
  }

  function formatoFecha(f) {
    if (!f) return '';
    try {
      return new Intl.DateTimeFormat('es-CO', { timeZone: 'America/Bogota', dateStyle: 'medium', timeStyle: 'short' }).format(new Date(f));
    } catch (e) { return String(f); }
  }

  function enlacesDeHtml(html) {
    if (!html) return [];
    try {
      const doc = new DOMParser().parseFromString(html, 'text/html');
      return Array.from(doc.querySelectorAll('a[href]'))
        .map(a => {
          const url = a.getAttribute('href') || '';
          const texto = (a.textContent || '').trim().replace(/\s+/g, ' ');
          const etiqueta = texto && texto !== url && !/^https?:\/\//i.test(texto) ? texto.slice(0, 120) : '';
          return { url, etiqueta };
        })
        .filter(e => /^https?:\/\//i.test(e.url));
    } catch (e) { return []; }
  }

  function base64ABytes(b64) {
    const bin = atob(b64);
    const bytes = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
    return bytes;
  }

  // Convierte GIF/BMP/WEBP a PNG y endereza fotos JPG con orientación EXIF
  async function convertirImagen(bytes, origen, destino) {
    const mime = { jpg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', webp: 'image/webp', png: 'image/png' }[origen];
    const bmp = await createImageBitmap(new Blob([bytes], { type: mime }), { imageOrientation: 'from-image' });
    const lienzo = document.createElement('canvas');
    lienzo.width = bmp.width; lienzo.height = bmp.height;
    const ctx = lienzo.getContext('2d');
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, lienzo.width, lienzo.height);
    ctx.drawImage(bmp, 0, 0);
    const tipo = destino === 'jpg' ? 'image/jpeg' : 'image/png';
    const salida = await new Promise(ok => lienzo.toBlob(ok, tipo, 0.92));
    return new Uint8Array(await salida.arrayBuffer());
  }

  // ------------------------------------------------------------------ interfaz

  function progreso(msg) {
    const p = $('progreso');
    p.textContent = msg;
    p.classList.remove('oculto');
  }

  function limpiarResultado() {
    if (urlPdf) { URL.revokeObjectURL(urlPdf); urlPdf = null; }
    $('resultado').classList.add('oculto');
    $('error').classList.add('oculto');
  }

  function mostrarError(msg) {
    const e = $('error');
    e.textContent = msg;
    e.classList.remove('oculto');
  }
})();
