// ===== Preksha Store — Google Sheets Backend =====
// Deploy this as a Web App (Deploy > New deployment > type: Web app,
// Execute as: Me, Who has access: Anyone).
//
// This version adds OTP-verified signup. Alongside your existing sheets
// (Orders, ProductOverrides, CustomProducts, Users, Settings), add a new
// sheet named exactly "OTPs" with this header row in row 1:
//   email | otp | expiresAt | createdAt
// Nothing else needs to be created — sendOtp/verifyOtp manage rows in it
// automatically; a row for an email is overwritten on resend.

function getSheet(name) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
}

function rowsToObjects(sheet) {
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  return values.slice(1)
    .filter(row => row.some(cell => cell !== ''))
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => obj[h] = row[i]);
      return obj;
    });
}

function getOrCreateReceiptsFolder() {
  const folderName = 'Preksha Store Receipts';
  const folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(folderName);
}

function jsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ===== "Plan activated" email =====
// Sent for free via Gmail's built-in MailApp service (the Google account
// this script is deployed under) — no EmailJS template needed, so it
// doesn't count against EmailJS's free-tier 2-template limit. Fires once,
// right when an order's status flips to "fulfilled" (see updateOrderStatus
// below). Consumer Gmail accounts get ~100 emails/day from MailApp, which
// is far more than order volume needs; Workspace accounts get more.
// ===== One-click test — run this manually from the Apps Script editor =====
// Select "test_sendPlanActivatedEmail" from the function dropdown at the
// top of the editor, click Run, and grant any permissions Google asks for
// (specifically "Send email as you"). If it finishes without an error, a
// real test email has landed in the inbox you set below — check your inbox
// (and Spam folder) to confirm formatting looks right.
// This is the fastest way to find out whether the "nothing happens" issue
// is a missing authorization vs. something else, because running it here
// shows you the FULL error message and stack trace directly in the editor,
// instead of it being swallowed or shown only as "Failed" in the browser.
function test_sendPlanActivatedEmail() {
  sendPlanActivatedEmail({
    id: 'TEST-0001',
    customerName: 'Test Customer',
    customerEmail: Session.getActiveUser().getEmail() || 'you@example.com', // sends to yourself
    items: [{ id: 'youtube-premium', name: 'YouTube Premium', price: 890, quantity: 1 }]
  }, 'test.member@example.com');
  Logger.log('If you see this with no error above, the email was sent successfully.');
}

function sendPlanActivatedEmail(order, memberEmail) {
  const serviceName = (order.items || []).map(it => it.name).join(', ') || 'your plan';
  const activationDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Colombo', 'd MMMM yyyy');
  const html = buildPlanActivatedEmailHtml({
    toName: order.customerName,
    orderId: order.id,
    serviceName: serviceName,
    memberEmail: memberEmail || '',
    activationDate: activationDate
  });
  MailApp.sendEmail({
    to: order.customerEmail,
    subject: 'Your ' + serviceName + ' plan is now active — Preksha Store',
    htmlBody: html,
    name: 'Preksha Store' // sets the "From" display name shown in the inbox
  });
}

function buildPlanActivatedEmailHtml(data) {
  const memberEmailRow = data.memberEmail
    ? data.memberEmail
    : '<span style="color:#b5b5c0;">—</span>';
  return `
<div style="margin:0; padding:0; background-color:#f0f0f3; width:100%;">
  <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background-color:#f0f0f3; padding:40px 12px;">
    <tr>
      <td align="center">
        <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="max-width:520px; background-color:#ffffff; border-radius:14px; overflow:hidden; border:1px solid #e6e6ec; font-family:'Segoe UI', Roboto, Helvetica, Arial, sans-serif;">
          <tr>
            <td style="height:4px; background-color:#ff3d6e; background-image:linear-gradient(90deg,#ff3d6e,#ffb020); font-size:0; line-height:0;">&nbsp;</td>
          </tr>
          <tr>
            <td style="padding:32px 32px 20px 32px; text-align:center;">
              <img src="https://github.com/Akila-Widuruwan/Preksha/blob/main/preksha-logo.png?raw=true" width="44" height="44" alt="Preksha Store" style="display:block; margin:0 auto 12px auto; border-radius:10px;">
              <div style="font-size:17px; font-weight:700; color:#0a0b10; letter-spacing:0.2px;">Preksha<span style="font-weight:400; color:#8a8a99;">.store</span></div>
            </td>
          </tr>
          <tr>
            <td style="padding:4px 32px 6px 32px; text-align:center;">
              <span style="display:inline-block; padding:6px 14px; background-color:#e8f7ee; color:#1a8a4d; font-size:12px; font-weight:700; letter-spacing:0.4px; text-transform:uppercase; border-radius:999px;">Plan Active</span>
              <h1 style="margin:18px 0 8px 0; font-size:21px; color:#0a0b10; font-weight:700;">Your ${data.serviceName} slot is ready</h1>
              <p style="margin:0; font-size:14.5px; color:#5a5a66; line-height:1.65;">
                Hi ${data.toName}, your subscription has been activated and you can start using it right away. Details for your records are below.
              </p>
            </td>
          </tr>
          <tr>
            <td style="padding:24px 32px 8px 32px;">
              <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background-color:#fafafc; border-radius:10px; border:1px solid #ececf2;">
                <tr>
                  <td style="padding:15px 20px; font-size:13px; color:#8a8a99; border-bottom:1px solid #ececf2;">Order ID</td>
                  <td style="padding:15px 20px; font-size:13px; color:#0a0b10; font-weight:600; text-align:right; border-bottom:1px solid #ececf2;">${data.orderId}</td>
                </tr>
                <tr>
                  <td style="padding:15px 20px; font-size:13px; color:#8a8a99; border-bottom:1px solid #ececf2;">Plan</td>
                  <td style="padding:15px 20px; font-size:13px; color:#0a0b10; font-weight:600; text-align:right; border-bottom:1px solid #ececf2;">${data.serviceName}</td>
                </tr>
                <tr>
                  <td style="padding:15px 20px; font-size:13px; color:#8a8a99; border-bottom:1px solid #ececf2;">Activated on</td>
                  <td style="padding:15px 20px; font-size:13px; color:#0a0b10; font-weight:600; text-align:right; border-bottom:1px solid #ececf2;">${data.activationDate}</td>
                </tr>
                <tr>
                  <td style="padding:15px 20px; font-size:13px; color:#8a8a99;">Slot login email</td>
                  <td style="padding:15px 20px; font-size:13px; color:#0a0b10; font-weight:600; text-align:right;">${memberEmailRow}</td>
                </tr>
              </table>
            </td>
          </tr>
          <tr>
            <td style="padding:20px 32px 4px 32px;">
              <p style="margin:0; font-size:12.5px; color:#9a9aa5; line-height:1.6; text-align:center;">
                If the slot login email above is blank or looks incorrect, simply reply to this email and we'll correct it right away.
              </p>
            </td>
          </tr>
          <tr>
            <td style="padding:26px 32px 30px 32px; text-align:center; border-top:1px solid #ececf2; margin-top:8px;">
              <p style="margin:0 0 4px 0; font-size:13px; color:#5a5a66; font-weight:600;">Thank you for choosing Preksha.store</p>
              <p style="margin:0; font-size:11.5px; color:#b5b5c0;">This is an automated message regarding your order — no action is needed unless something looks incorrect.</p>
            </td>
          </tr>
        </table>
      </td>
    </tr>
  </table>
</div>`;
}

function doGet(e) {
  const action = e.parameter.action;
  try {
    if (action === 'getOrders') {
      const orders = rowsToObjects(getSheet('Orders')).map(o => ({
        id: o.id, date: o.date, payment: o.payment, total: o.total,
        status: o.status, memberEmail: o.memberEmail,
        referenceNumber: o.referenceNumber || '', receiptUrl: o.receiptUrl || '',
        customer: { name: o.customerName, email: o.customerEmail },
        items: JSON.parse(o.items || '[]')
      })).reverse(); // newest first
      return jsonResponse({ ok: true, data: orders });
    }
    if (action === 'getProducts') {
      const overrides = {};
      rowsToObjects(getSheet('ProductOverrides')).forEach(r => {
        overrides[r.id] = JSON.parse(r.data || '{}');
      });
      return jsonResponse({ ok: true, data: overrides });
    }
    if (action === 'getCustomProducts') {
      const list = rowsToObjects(getSheet('CustomProducts')).map(r => JSON.parse(r.data));
      return jsonResponse({ ok: true, data: list });
    }
    if (action === 'getUsers') {
      return jsonResponse({ ok: true, data: rowsToObjects(getSheet('Users')) });
    }
    if (action === 'getSetting') {
      const rows = rowsToObjects(getSheet('Settings'));
      const row = rows.find(r => r.key === e.parameter.key);
      return jsonResponse({ ok: true, data: row ? row.value : null });
    }
    return jsonResponse({ ok: false, error: 'Unknown action' });
  } catch (err) {
    return jsonResponse({ ok: false, error: err.message });
  }
}

function doPost(e) {
  const action = e.parameter.action;
  const body = JSON.parse(e.postData.contents || '{}');
  try {
    if (action === 'addOrder') {
      const s = getSheet('Orders');
      s.appendRow([
        body.id, body.date, body.customer.name, body.customer.email,
        body.payment, JSON.stringify(body.items), body.total,
        body.status, body.memberEmail || '',
        body.referenceNumber || '', body.receiptUrl || ''
      ]);
      return jsonResponse({ ok: true });
    }
    if (action === 'uploadReceipt') {
      // Saves a bank-slip / bank-transfer receipt image (sent as base64 from the
      // storefront) into a Drive folder and returns a viewable URL for the
      // admin dashboard to show. The file is shared as "anyone with the link
      // can view" so the <img> tag in the admin dashboard can load it directly.
      const folder = getOrCreateReceiptsFolder();
      const bytes = Utilities.base64Decode(body.base64);
      const mimeType = body.mimeType || 'image/jpeg';
      const fileName = body.fileName || ('receipt-' + Date.now() + '.jpg');
      const blob = Utilities.newBlob(bytes, mimeType, fileName);
      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      const fileId = file.getId();
      return jsonResponse({
        ok: true,
        data: {
          fileId: fileId,
          url: 'https://drive.google.com/uc?export=view&id=' + fileId,
          viewUrl: 'https://drive.google.com/file/d/' + fileId + '/view'
        }
      });
    }
    if (action === 'clearAllOrders') {
      const s = getSheet('Orders');
      const lastRow = s.getLastRow();
      if (lastRow > 1) s.getRange(2, 1, lastRow - 1, s.getLastColumn()).clearContent();
      return jsonResponse({ ok: true });
    }
    if (action === 'deleteOrder') {
      const s = getSheet('Orders');
      const values = s.getDataRange().getValues();
      for (let i = 1; i < values.length; i++) {
        if (values[i][0] === body.id) {
          s.deleteRow(i + 1);
          break;
        }
      }
      return jsonResponse({ ok: true });
    }
    if (action === 'updateOrderStatus') {
      const s = getSheet('Orders');
      const values = s.getDataRange().getValues();
      for (let i = 1; i < values.length; i++) {
        if (values[i][0] === body.id) {
          const prevStatus = values[i][7]; // status column, before this update
          s.getRange(i + 1, 8).setValue(body.status); // status column
          if (body.memberEmail !== undefined) s.getRange(i + 1, 9).setValue(body.memberEmail);
          // Fire the "plan activated" email exactly once, the moment this
          // order newly becomes fulfilled — not on every re-save of an
          // already-fulfilled order.
          if (body.status === 'fulfilled' && prevStatus !== 'fulfilled') {
            const row = values[i];
            const order = {
              id: row[0], customerName: row[2], customerEmail: row[3],
              items: JSON.parse(row[5] || '[]')
            };
            const memberEmail = body.memberEmail !== undefined ? body.memberEmail : row[8];
            try {
              sendPlanActivatedEmail(order, memberEmail);
            } catch (mailErr) {
              // Don't fail the status update if the email fails to send —
              // just log it so it shows up in Apps Script's execution log.
              console.error('sendPlanActivatedEmail failed: ' + mailErr);
            }
          }
          break;
        }
      }
      return jsonResponse({ ok: true });
    }
    if (action === 'resendPlanActivatedEmail') {
      // Manually (re)sends the plan-activated email regardless of the
      // order's current/previous status — useful for testing the setup, or
      // for resending if a customer says they never got it. Unlike the
      // automatic send above, errors here are NOT swallowed — they're
      // returned to the admin dashboard so you can actually see what went
      // wrong (e.g. missing Gmail send authorization).
      const s = getSheet('Orders');
      const values = s.getDataRange().getValues();
      for (let i = 1; i < values.length; i++) {
        if (values[i][0] === body.id) {
          const row = values[i];
          const order = {
            id: row[0], customerName: row[2], customerEmail: row[3],
            items: JSON.parse(row[5] || '[]')
          };
          sendPlanActivatedEmail(order, row[8]);
          return jsonResponse({ ok: true });
        }
      }
      return jsonResponse({ ok: false, error: 'Order not found' });
    }
    if (action === 'saveProductOverride') {
      const s = getSheet('ProductOverrides');
      const values = s.getDataRange().getValues();
      let found = false;
      for (let i = 1; i < values.length; i++) {
        if (values[i][0] === body.id) {
          s.getRange(i + 1, 2).setValue(JSON.stringify(body.data));
          found = true;
          break;
        }
      }
      if (!found) s.appendRow([body.id, JSON.stringify(body.data)]);
      return jsonResponse({ ok: true });
    }
    if (action === 'saveCustomProduct') {
      const s = getSheet('CustomProducts');
      const values = s.getDataRange().getValues();
      let found = false;
      for (let i = 1; i < values.length; i++) {
        if (values[i][0] === body.id) {
          s.getRange(i + 1, 2).setValue(JSON.stringify(body));
          found = true;
          break;
        }
      }
      if (!found) s.appendRow([body.id, JSON.stringify(body)]);
      return jsonResponse({ ok: true });
    }
    if (action === 'deleteCustomProduct') {
      const s = getSheet('CustomProducts');
      const values = s.getDataRange().getValues();
      for (let i = 1; i < values.length; i++) {
        if (values[i][0] === body.id) {
          s.deleteRow(i + 1);
          break;
        }
      }
      return jsonResponse({ ok: true });
    }
    if (action === 'sendOtp') {
      const email = String(body.email || '').toLowerCase().trim();
      if (!email) return jsonResponse({ ok: false, error: 'Email is required.' });
      const otp = String(Math.floor(100000 + Math.random() * 900000)); // 6 digits
      const expiresAt = new Date(Date.now() + 10 * 60 * 1000).toISOString(); // 10 min
      const createdAt = new Date().toISOString();
      const s = getSheet('OTPs');
      const values = s.getDataRange().getValues();
      let found = false;
      for (let i = 1; i < values.length; i++) {
        if (String(values[i][0]).toLowerCase() === email) {
          s.getRange(i + 1, 1, 1, 4).setValues([[email, otp, expiresAt, createdAt]]);
          found = true;
          break;
        }
      }
      if (!found) s.appendRow([email, otp, expiresAt, createdAt]);
      // otp is returned so the client can hand it to EmailJS for delivery —
      // EmailJS only works client-side, so the code must pass through the
      // client either way. verifyOtp below is still the source of truth.
      return jsonResponse({ ok: true, data: { otp: otp, expiresAt: expiresAt } });
    }
    if (action === 'verifyOtp') {
      const email = String(body.email || '').toLowerCase().trim();
      const code = String(body.otp || '').trim();
      const s = getSheet('OTPs');
      const values = s.getDataRange().getValues();
      for (let i = 1; i < values.length; i++) {
        if (String(values[i][0]).toLowerCase() === email) {
          const storedOtp = String(values[i][1]);
          const expiresAt = new Date(values[i][2]);
          if (new Date() > expiresAt) {
            s.deleteRow(i + 1);
            return jsonResponse({ ok: false, error: 'expired' });
          }
          if (storedOtp !== code) {
            return jsonResponse({ ok: false, error: 'invalid' });
          }
          s.deleteRow(i + 1); // consume — one-time use
          return jsonResponse({ ok: true });
        }
      }
      return jsonResponse({ ok: false, error: 'not_found' });
    }
    if (action === 'saveUser') {
      const s = getSheet('Users');
      const values = s.getDataRange().getValues();
      let found = false;
      for (let i = 1; i < values.length; i++) {
        if (values[i][0] === body.email) {
          s.getRange(i + 1, 1, 1, 5).setValues([[
            body.email, body.passwordHash || '', body.name || '',
            body.provider || '', body.photoUrl || ''
          ]]);
          found = true;
          break;
        }
      }
      if (!found) s.appendRow([body.email, body.passwordHash || '', body.name || '', body.provider || '', body.photoUrl || '']);
      return jsonResponse({ ok: true });
    }
    if (action === 'setSetting') {
      const s = getSheet('Settings');
      const values = s.getDataRange().getValues();
      let found = false;
      for (let i = 1; i < values.length; i++) {
        if (values[i][0] === body.key) {
          s.getRange(i + 1, 2).setValue(body.value);
          found = true;
          break;
        }
      }
      if (!found) s.appendRow([body.key, body.value]);
      return jsonResponse({ ok: true });
    }
    return jsonResponse({ ok: false, error: 'Unknown action' });
  } catch (err) {
    return jsonResponse({ ok: false, error: err.message });
  }
}
