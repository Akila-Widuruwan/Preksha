// ===== Preksha Store — Google Sheets Backend =====
// Deploy this as a Web App (Deploy > New deployment > type: Web app,
// Execute as: Me, Who has access: Anyone).

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
          s.getRange(i + 1, 8).setValue(body.status); // status column
          if (body.memberEmail !== undefined) s.getRange(i + 1, 9).setValue(body.memberEmail);
          break;
        }
      }
      return jsonResponse({ ok: true });
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
