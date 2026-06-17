// ============================================================
// SIGV — Apps Script para Google Sheets (Backend da Frota)
// ============================================================
// INSTRUÇÕES DE IMPLANTAÇÃO:
// 1. Abra a planilha: https://docs.google.com/spreadsheets/d/1kPxHyUi7CbEDdJaTabiqFs5P5cjYfjQtnoTVBKND61c
// 2. Menu: Extensões > Apps Script
// 3. Apague todo o conteúdo de Code.gs e cole este código
// 4. Salve (Ctrl+S)
// 5. Clique em "Implantar" > "Nova implantação"
// 6. Tipo: "App da Web"
// 7. Executar como: "Eu" (sua conta)
// 8. Quem tem acesso: "Qualquer pessoa"
// 9. Clique "Implantar" e autorize quando solicitado
// 10. Copie a URL gerada e cole no SIGV (ícone ⚙ no menu lateral)
// ============================================================

var SHEET_NAME = 'Frota';
var HEADERS = ['Prefixo','Modalidade','Modelo','Placa','Ano','Km','CustoMan','Status'];

function doGet(e) {
  var sheet = getOrCreateSheet();
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return jsonResp([]);

  var rows = [];
  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    if (!r[0]) continue;
    rows.push({
      pref: String(r[0]),
      mod: String(r[1]),
      modelo: String(r[2]),
      placa: String(r[3]),
      ano: Number(r[4]) || 0,
      km: Number(r[5]) || 0,
      custoMan: Number(r[6]) || 0,
      status: String(r[7]) || 'Operacional'
    });
  }
  return jsonResp(rows);
}

function doPost(e) {
  var body = JSON.parse(e.postData.contents);
  var sheet = getOrCreateSheet();

  if (body.action === 'save')    return salvar(sheet, body.vtr);
  if (body.action === 'delete')  return excluir(sheet, body.pref);
  if (body.action === 'saveAll') return salvarTodos(sheet, body.vtrs);

  return jsonResp({error: 'Ação inválida'});
}

function getOrCreateSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
  }
  return sheet;
}

function salvar(sheet, vtr) {
  var data = sheet.getDataRange().getValues();
  var rowIdx = -1;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(vtr.pref)) { rowIdx = i + 1; break; }
  }
  var row = [vtr.pref, vtr.mod, vtr.modelo, vtr.placa, vtr.ano, vtr.km, vtr.custoMan, vtr.status];
  if (rowIdx > 0) {
    sheet.getRange(rowIdx, 1, 1, HEADERS.length).setValues([row]);
  } else {
    sheet.appendRow(row);
  }
  return jsonResp({ok: true});
}

function excluir(sheet, pref) {
  var data = sheet.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]) === String(pref)) {
      sheet.deleteRow(i + 1);
      break;
    }
  }
  return jsonResp({ok: true});
}

function salvarTodos(sheet, vtrs) {
  sheet.clear();
  sheet.appendRow(HEADERS);
  sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
  if (vtrs && vtrs.length > 0) {
    var rows = vtrs.map(function(v) {
      return [v.pref, v.mod, v.modelo, v.placa, v.ano, v.km, v.custoMan, v.status];
    });
    sheet.getRange(2, 1, rows.length, HEADERS.length).setValues(rows);
  }
  return jsonResp({ok: true});
}

function jsonResp(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
