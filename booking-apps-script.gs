/**
 * ═══════════════════════════════════════════════════════════════════════
 *  SAN JOSÉ HOTEL — Google Apps Script para Sistema de Reservas
 *  Versión: 1.0 | Mayo 2026
 * ═══════════════════════════════════════════════════════════════════════
 *
 *  INSTRUCCIONES DE CONFIGURACIÓN:
 *
 *  1. Abrí Google Sheets en tu cuenta de Google:
 *     https://sheets.google.com → Crear nueva hoja
 *
 *  2. Renombrá la primera hoja como "Reservas"
 *
 *  3. En el menú: Extensiones → Apps Script → pegá este código
 *
 *  4. Cambiá SPREADSHEET_ID (lo encontrás en la URL de tu Google Sheet:
 *     https://docs.google.com/spreadsheets/d/[ESTE_ES_EL_ID]/edit)
 *
 *  5. Cambiá HOTEL_EMAIL por el email del hotel
 *
 *  6. Click en "Implementar" → "Nueva implementación"
 *     - Tipo: Aplicación web
 *     - Ejecutar como: Yo
 *     - Quién tiene acceso: Cualquier persona (anónima)
 *     → Copiá la URL generada
 *
 *  7. Pegá esa URL en index.html en la variable APPS_SCRIPT_URL
 *
 * ═══════════════════════════════════════════════════════════════════════
 */

const SPREADSHEET_ID = '1HA5t9FQghgJN3zcgpUjRD9Hzhk-S-zbYoe80_tFtMEY';
const HOTEL_EMAIL    = 'sanjosehotelvcp@gmail.com';
const HOTEL_NAME     = 'San José Hotel';
const HOTEL_ADDRESS  = 'Alvear 195, Villa Carlos Paz, Córdoba';
const HOTEL_WHATSAPP = '3541 372428';
const HOTEL_INSTAGRAM= '@sanjosehotelvcp';

/* ── Manejo de GET (verificar disponibilidad) ── */
function doGet(e) {
  const action = e.parameter.action;

  if (action === 'checkAvailability') {
    return checkAvailability(e.parameter);
  }

  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', message: 'San José Hotel Booking API' }))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ── Manejo de POST (guardar reserva) ── */
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    if (data.action === 'saveBooking') {
      return saveBooking(data);
    }

    return jsonResponse({ success: false, error: 'Acción no reconocida.' });
  } catch (err) {
    return jsonResponse({ success: false, error: err.toString() });
  }
}

/* ═══════════════════════════════════════════════
   VERIFICAR DISPONIBILIDAD
═══════════════════════════════════════════════ */
function checkAvailability(params) {
  const { hab, checkin, checkout } = params;
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Reservas');
  const rows  = sheet.getDataRange().getValues();

  const newIn  = new Date(checkin);
  const newOut = new Date(checkout);

  /* Verificar si hay conflicto con reservas existentes (estado: confirmada/pendiente) */
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    const existHab    = row[6];  // columna G: habitación
    const existIn     = new Date(row[9]);  // columna J: checkin
    const existOut    = new Date(row[10]); // columna K: checkout
    const estado      = row[11]; // columna L: estado

    if (existHab !== hab) continue;
    if (estado === 'cancelada') continue;

    /* Hay conflicto si los rangos de fecha se superponen */
    if (newIn < existOut && newOut > existIn) {
      return jsonResponse({ available: false });
    }
  }

  return jsonResponse({ available: true });
}

/* ═══════════════════════════════════════════════
   GUARDAR RESERVA
═══════════════════════════════════════════════ */
function saveBooking(data) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Reservas');

  /* Inicializar encabezados si la hoja está vacía */
  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      'ID Reserva', 'Timestamp', 'Nombre', 'DNI', 'Teléfono', 'Email',
      'Habitación', 'Noches', 'Dirección', 'Check-in', 'Check-out',
      'Estado', 'Notas', 'Fecha Creación'
    ]);
    /* Formato de encabezados */
    const headerRange = sheet.getRange(1, 1, 1, 14);
    headerRange.setBackground('#4A5830');
    headerRange.setFontColor('#FFFFFF');
    headerRange.setFontWeight('bold');
  }

  /* Generar ID único */
  const reservaId = 'SJ-' + new Date().getFullYear() + '-' + String(sheet.getLastRow()).padStart(4, '0');
  const fechaCreacion = new Date().toLocaleString('es-AR', { timeZone: 'America/Argentina/Buenos_Aires' });

  /* Agregar fila */
  sheet.appendRow([
    reservaId,
    data.timestamp,
    data.nombre,
    data.dni,
    data.telefono,
    data.email,
    data.hab,
    data.noches,
    data.direccion || '',
    data.checkin,
    data.checkout,
    'pendiente_seña',
    data.notas || '',
    fechaCreacion
  ]);

  /* Dar formato a la nueva fila */
  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow, 12).setBackground('#FFF9C4'); /* Estado: amarillo */

  /* Enviar email de confirmación al huésped */
  try {
    sendConfirmationEmail(data, reservaId);
  } catch(e) {
    Logger.log('Error enviando email: ' + e.toString());
  }

  /* Enviar notificación interna al hotel */
  try {
    sendHotelNotification(data, reservaId);
  } catch(e) {
    Logger.log('Error enviando notificación hotel: ' + e.toString());
  }

  return jsonResponse({
    success:   true,
    reservaId: reservaId,
    message:   'Reserva registrada correctamente.'
  });
}

/* ═══════════════════════════════════════════════
   EMAIL DE CONFIRMACIÓN AL HUÉSPED
═══════════════════════════════════════════════ */
function sendConfirmationEmail(data, reservaId) {
  const checkinFmt  = formatDateES(data.checkin);
  const checkoutFmt = formatDateES(data.checkout);

  const subject = `✅ Confirmación de solicitud de reserva — ${HOTEL_NAME} [${reservaId}]`;

  const htmlBody = `
    <!DOCTYPE html>
    <html lang="es">
    <head>
      <meta charset="UTF-8">
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
    </head>
    <body style="margin:0;padding:0;background:#F5F0E8;font-family:'Helvetica Neue',Helvetica,Arial,sans-serif;">
      <table width="100%" cellpadding="0" cellspacing="0" style="background:#F5F0E8;padding:40px 20px;">
        <tr>
          <td align="center">
            <table width="600" cellpadding="0" cellspacing="0" style="background:#FDFAF5;border-radius:12px;overflow:hidden;box-shadow:0 4px 24px rgba(0,0,0,0.08);">

              <!-- Header -->
              <tr>
                <td style="background:#4A5830;padding:40px 48px;text-align:center;">
                  <p style="color:#B5C28A;font-size:11px;letter-spacing:4px;text-transform:uppercase;margin:0 0 8px;">Hotel</p>
                  <h1 style="color:#FDFAF5;font-size:28px;font-weight:300;margin:0;letter-spacing:2px;font-style:italic;">San José</h1>
                  <p style="color:#B5C28A;font-size:10px;letter-spacing:3px;text-transform:uppercase;margin:8px 0 0;">Villa Carlos Paz · Córdoba</p>
                </td>
              </tr>

              <!-- Body -->
              <tr>
                <td style="padding:40px 48px;">
                  <p style="color:#C4704A;font-size:11px;letter-spacing:3px;text-transform:uppercase;margin:0 0 8px;">Solicitud recibida</p>
                  <h2 style="color:#2A2826;font-size:22px;font-weight:400;margin:0 0 16px;">Hola, ${data.nombre.split(' ')[0]} 👋</h2>
                  <p style="color:#8A8880;font-size:14px;line-height:1.8;margin:0 0 32px;">
                    Recibimos tu solicitud de reserva. Para confirmarla, vamos a contactarte por WhatsApp para coordinar el pago de la seña.
                    ¡Estamos con vos!
                  </p>

                  <!-- Resumen reserva -->
                  <table width="100%" cellpadding="0" cellspacing="0" style="background:#F5F0E8;border-radius:8px;margin-bottom:32px;">
                    <tr><td style="padding:24px 28px;">
                      <p style="color:#6B7A4A;font-size:10px;letter-spacing:3px;text-transform:uppercase;margin:0 0 20px;">Detalle de tu solicitud</p>
                      <table width="100%" cellpadding="0" cellspacing="0">
                        <tr>
                          <td style="padding:8px 0;border-bottom:1px solid #E8D9C8;color:#8A8880;font-size:13px;">ID de reserva</td>
                          <td style="padding:8px 0;border-bottom:1px solid #E8D9C8;color:#2A2826;font-size:13px;font-weight:600;text-align:right;">${reservaId}</td>
                        </tr>
                        <tr>
                          <td style="padding:8px 0;border-bottom:1px solid #E8D9C8;color:#8A8880;font-size:13px;">Habitación</td>
                          <td style="padding:8px 0;border-bottom:1px solid #E8D9C8;color:#2A2826;font-size:13px;text-align:right;">${data.hab}</td>
                        </tr>
                        <tr>
                          <td style="padding:8px 0;border-bottom:1px solid #E8D9C8;color:#8A8880;font-size:13px;">Check-in</td>
                          <td style="padding:8px 0;border-bottom:1px solid #E8D9C8;color:#2A2826;font-size:13px;text-align:right;">${checkinFmt}</td>
                        </tr>
                        <tr>
                          <td style="padding:8px 0;border-bottom:1px solid #E8D9C8;color:#8A8880;font-size:13px;">Check-out</td>
                          <td style="padding:8px 0;border-bottom:1px solid #E8D9C8;color:#2A2826;font-size:13px;text-align:right;">${checkoutFmt}</td>
                        </tr>
                        <tr>
                          <td style="padding:10px 0 0;color:#2A2826;font-size:14px;font-weight:600;">Total de noches</td>
                          <td style="padding:10px 0 0;color:#6B7A4A;font-size:16px;font-weight:700;text-align:right;">${data.noches} noches</td>
                        </tr>
                      </table>
                    </td></tr>
                  </table>

                  <!-- Próximos pasos -->
                  <table width="100%" cellpadding="0" cellspacing="0" style="background:#F0F4E8;border-left:3px solid #6B7A4A;border-radius:0 8px 8px 0;margin-bottom:32px;">
                    <tr><td style="padding:20px 24px;">
                      <p style="color:#4A5830;font-size:13px;font-weight:600;margin:0 0 10px;">¿Qué sigue?</p>
                      <p style="color:#6B7A4A;font-size:13px;line-height:1.7;margin:0;">
                        📱 <strong>1.</strong> Nos ponemos en contacto por WhatsApp al número que nos dejaste.<br>
                        💳 <strong>2.</strong> Coordinamos el pago de la seña para confirmar la reserva.<br>
                        ✅ <strong>3.</strong> Te enviamos la confirmación final con todos los detalles.
                      </p>
                    </td></tr>
                  </table>

                  ${data.notas ? `
                  <table width="100%" cellpadding="0" cellspacing="0" style="background:#FFF8F0;border:1px solid #E8D9C8;border-radius:8px;margin-bottom:32px;">
                    <tr><td style="padding:16px 20px;">
                      <p style="color:#C4704A;font-size:11px;letter-spacing:2px;text-transform:uppercase;margin:0 0 6px;">Tu nota</p>
                      <p style="color:#8A8880;font-size:13px;margin:0;font-style:italic;">"${data.notas}"</p>
                    </td></tr>
                  </table>
                  ` : ''}

                  <p style="color:#8A8880;font-size:13px;line-height:1.8;margin:0;">
                    ¿Tenés alguna duda? Escribinos directamente:<br>
                    📱 WhatsApp: <a href="https://wa.me/54${HOTEL_WHATSAPP.replace(/\s/g,'')}" style="color:#6B7A4A;">${HOTEL_WHATSAPP}</a><br>
                    📸 Instagram: <a href="https://instagram.com/sanjosehotelvcp" style="color:#6B7A4A;">${HOTEL_INSTAGRAM}</a>
                  </p>
                </td>
              </tr>

              <!-- Footer -->
              <tr>
                <td style="background:#2A2826;padding:28px 48px;text-align:center;">
                  <p style="color:rgba(253,250,245,0.5);font-size:11px;margin:0 0 4px;letter-spacing:0.5px;">
                    ${HOTEL_NAME} · ${HOTEL_ADDRESS}
                  </p>
                  <p style="color:rgba(253,250,245,0.25);font-size:10px;margin:0;font-style:italic;">
                    "Alma de sierra y amor de familia."
                  </p>
                </td>
              </tr>

            </table>
          </td>
        </tr>
      </table>
    </body>
    </html>
  `;

  GmailApp.sendEmail(
    data.email,
    subject,
    `Hola ${data.nombre.split(' ')[0]}, recibimos tu solicitud de reserva en ${HOTEL_NAME}. ID: ${reservaId}. Check-in: ${checkinFmt}. Check-out: ${checkoutFmt}. Habitación: ${data.hab}. Nos contactamos por WhatsApp para coordinar la seña. ¡Gracias!`,
    {
      htmlBody:    htmlBody,
      name:        HOTEL_NAME,
      replyTo:     HOTEL_EMAIL
    }
  );
}

/* ═══════════════════════════════════════════════
   NOTIFICACIÓN INTERNA AL HOTEL
═══════════════════════════════════════════════ */
function sendHotelNotification(data, reservaId) {
  const checkinFmt  = formatDateES(data.checkin);
  const checkoutFmt = formatDateES(data.checkout);

  GmailApp.sendEmail(
    HOTEL_EMAIL,
    `🛎️ Nueva solicitud de reserva — ${data.nombre} [${reservaId}]`,
    '',
    {
      htmlBody: `
        <h2 style="font-family:sans-serif;color:#4A5830;">Nueva reserva recibida</h2>
        <table style="font-family:sans-serif;font-size:14px;border-collapse:collapse;">
          <tr><td style="padding:6px 16px 6px 0;color:#888;"><b>ID:</b></td><td>${reservaId}</td></tr>
          <tr><td style="padding:6px 16px 6px 0;color:#888;"><b>Nombre:</b></td><td>${data.nombre}</td></tr>
          <tr><td style="padding:6px 16px 6px 0;color:#888;"><b>DNI:</b></td><td>${data.dni}</td></tr>
          <tr><td style="padding:6px 16px 6px 0;color:#888;"><b>Teléfono:</b></td><td>${data.telefono}</td></tr>
          <tr><td style="padding:6px 16px 6px 0;color:#888;"><b>Email:</b></td><td>${data.email}</td></tr>
          <tr><td style="padding:6px 16px 6px 0;color:#888;"><b>Domicilio:</b></td><td>${data.direccion || '—'}</td></tr>
          <tr><td style="padding:6px 16px 6px 0;color:#888;"><b>Habitación:</b></td><td><b>${data.hab}</b></td></tr>
          <tr><td style="padding:6px 16px 6px 0;color:#888;"><b>Check-in:</b></td><td>${checkinFmt}</td></tr>
          <tr><td style="padding:6px 16px 6px 0;color:#888;"><b>Check-out:</b></td><td>${checkoutFmt}</td></tr>
          <tr><td style="padding:6px 16px 6px 0;color:#888;"><b>Noches:</b></td><td>${data.noches}</td></tr>
          ${data.notas ? `<tr><td style="padding:6px 16px 6px 0;color:#888;"><b>Notas:</b></td><td>${data.notas}</td></tr>` : ''}
        </table>
        <p style="font-family:sans-serif;margin-top:20px;">
          <a href="https://wa.me/54${data.telefono.replace(/\D/g,'')}"
             style="background:#25D366;color:white;padding:10px 20px;border-radius:6px;text-decoration:none;font-family:sans-serif;">
            Contactar por WhatsApp
          </a>
        </p>
      `,
      name: 'San José Hotel — Sistema de Reservas'
    }
  );
}

/* ═══════════════════════════════════════════════
   UTILS
═══════════════════════════════════════════════ */
function formatDateES(dateStr) {
  const months = ['enero','febrero','marzo','abril','mayo','junio','julio','agosto','septiembre','octubre','noviembre','diciembre'];
  const d = new Date(dateStr + 'T12:00:00');
  return `${d.getDate()} de ${months[d.getMonth()]} de ${d.getFullYear()}`;
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ═══════════════════════════════════════════════
   FUNCIÓN DE PRUEBA (ejecutar manualmente)
═══════════════════════════════════════════════ */
function testBooking() {
  const testData = {
    action:    'saveBooking',
    nombre:    'Juan Pérez',
    dni:       '28.000.000',
    telefono:  '351 123 4567',
    email:     'juan@test.com',
    direccion: 'Córdoba Capital',
    notas:     'Llegada tardía aprox 22hs',
    hab:       'Doble',
    checkin:   '2026-07-15',
    checkout:  '2026-07-18',
    noches:    3,
    timestamp: new Date().toISOString()
  };
  const result = saveBooking(testData);
  Logger.log(result.getContent());
}
