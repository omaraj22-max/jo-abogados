/**
 * JP ABOGADOS · Receptor de leads → Google Sheets
 * ---------------------------------------------------------------
 * CÓMO INSTALARLO (5 minutos, sin n8n):
 * 1. Abre tu Google Sheet (o crea una nueva).
 * 2. Menú: Extensiones → Apps Script.
 * 3. Borra lo que haya y pega TODO este archivo.
 * 4. Cambia SHEET_NAME si tu pestaña no se llama "Leads".
 * 5. Menú: Implementar → Nueva implementación → tipo "Aplicación web".
 *    - Ejecutar como: Yo
 *    - Quién tiene acceso: Cualquiera
 *    - Copia la URL que termina en /exec → esa va en la landing (WEBHOOK_URL).
 * 6. (Opcional) Pon tu correo en NOTIFY_EMAIL para recibir aviso de leads calientes.
 * ---------------------------------------------------------------
 */

const SHEET_NAME = 'Leads';
const NOTIFY_EMAIL = ''; // ej. 'contacto@jp-abogados.com.mx' — vacío = sin correo

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);

    // Crea la pestaña y encabezados si no existen
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow([
        'Fecha', 'Nombre', 'Teléfono', 'Tipo de caso', 'Urgencia',
        '¿Buscas contratar?', 'Qué pasó', 'Score', 'Banda',
        'utm_source', 'utm_campaign', 'utm_content', 'Página destino'
      ]);
      sheet.getRange(1, 1, 1, 13).setFontWeight('bold');
    }

    sheet.appendRow([
      new Date(),
      data.nombre || '',
      data.tel || '',
      data.caso || '',
      data.urgencia || '',
      data.intencion || '',
      data.detalle || '',
      data.score || '',
      data.banda || '',
      data.utm_source || '',
      data.utm_campaign || '',
      data.utm_content || '',
      data.destino || ''
    ]);

    // Aviso por correo SOLO para leads calientes (banda 80+)
    if (NOTIFY_EMAIL && data.banda === 'CALIENTE') {
      MailApp.sendEmail({
        to: NOTIFY_EMAIL,
        subject: '🔥 LEAD CALIENTE JP — ' + (data.caso || 'caso') + ' — contactar YA',
        body: 'Lead prioritario (score ' + data.score + ')\n\n' +
              'Nombre: ' + (data.nombre || '') + '\n' +
              'Teléfono: ' + (data.tel || '') + '\n' +
              'Caso: ' + (data.caso || '') + '\n' +
              'Urgencia: ' + (data.urgencia || '') + '\n' +
              'Intención: ' + (data.intencion || '') + '\n' +
              'Qué pasó: ' + (data.detalle || '') + '\n\n' +
              'Responde en menos de 5 minutos.'
      });
    }

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Permite probar que la implementación está viva (abre la URL en el navegador)
function doGet() {
  return ContentService.createTextOutput('JP Abogados lead receiver activo.');
}
