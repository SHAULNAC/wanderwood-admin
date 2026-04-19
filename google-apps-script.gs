/**
 * WanderWood Admin backend for Google Apps Script.
 *
 * Expected sheet columns (row 1 headers):
 * A: id
 * B: name
 * C: price
 * D: description
 * E: image   (comma-separated URLs)
 */

function doGet(e) {
  try {
    var action = (e && e.parameter && e.parameter.action) ? e.parameter.action : 'list';

    if (action !== 'list') {
      return jsonOutput({
        ok: false,
        error: 'Unsupported action for GET',
        action: action
      });
    }

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var values = sheet.getDataRange().getValues();

    if (values.length <= 1) {
      return jsonOutput({ ok: true, products: [] });
    }

    var headers = values[0].map(function (h) { return String(h).trim(); });
    var products = [];

    for (var r = 1; r < values.length; r++) {
      var row = values[r];
      if (row.join('') === '') continue;

      var product = {};
      for (var c = 0; c < headers.length; c++) {
        var key = headers[c] || ('col_' + c);
        product[key] = row[c];
      }

      if (!product.id) {
        product.id = String(r + 1);
      }
      if (!product.image) {
        product.image = '';
      }

      products.push(product);
    }

    return jsonOutput({ ok: true, products: products });
  } catch (err) {
    return jsonOutput({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  try {
    var payload = parsePayload_(e);
    var action = payload.action || 'create';

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    ensureHeaders_(sheet);

    if (action === 'create') {
      var newId = payload.id || String(new Date().getTime());
      sheet.appendRow([
        newId,
        payload.name || '',
        payload.price || '',
        payload.description || '',
        payload.image || ''
      ]);

      return jsonOutput({ ok: true, action: 'create', id: newId });
    }

    if (action === 'update') {
      var targetId = payload.originalId || payload.id;
      if (!targetId) {
        return jsonOutput({ ok: false, error: 'Missing id/originalId for update' });
      }

      var values = sheet.getDataRange().getValues();
      var rowIndex = findRowById_(values, String(targetId));
      if (rowIndex === -1) {
        return jsonOutput({ ok: false, error: 'Product not found', id: targetId });
      }

      sheet.getRange(rowIndex, 1, 1, 5).setValues([[
        payload.id || targetId,
        payload.name || '',
        payload.price || '',
        payload.description || '',
        payload.image || ''
      ]]);

      return jsonOutput({ ok: true, action: 'update', id: payload.id || targetId });
    }

    return jsonOutput({ ok: false, error: 'Unsupported action', action: action });
  } catch (err) {
    return jsonOutput({ ok: false, error: String(err) });
  }
}

function parsePayload_(e) {
  var body = (e && e.postData && e.postData.contents) ? e.postData.contents : '{}';
  return JSON.parse(body);
}

function ensureHeaders_(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['id', 'name', 'price', 'description', 'image']);
    return;
  }

  var header = sheet.getRange(1, 1, 1, 5).getValues()[0];
  var expected = ['id', 'name', 'price', 'description', 'image'];
  var mismatch = false;

  for (var i = 0; i < expected.length; i++) {
    if (String(header[i] || '').trim() !== expected[i]) {
      mismatch = true;
      break;
    }
  }

  if (mismatch) {
    sheet.getRange(1, 1, 1, 5).setValues([expected]);
  }
}

function findRowById_(values, targetId) {
  // values includes header row; spreadsheet rows are 1-indexed
  for (var r = 1; r < values.length; r++) {
    if (String(values[r][0]) === targetId) {
      return r + 1;
    }
  }
  return -1;
}

function jsonOutput(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
