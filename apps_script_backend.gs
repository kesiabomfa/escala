/**
 * Backend simples usando Google Sheets como banco de dados
 * para os plantões trocados (overrides) da Escala de Anestesiologia 2026.
 *
 * COMO INSTALAR:
 * 1. Na planilha, vá em Extensões > Apps Script.
 * 2. Apague o conteúdo padrão e cole este arquivo inteiro.
 * 3. Clique em "Implantar" (Deploy) > "Nova implantação".
 * 4. Tipo: "App da Web" (Web app).
 * 5. Executar como: "Eu" (sua conta).
 * 6. Quem tem acesso: "Qualquer pessoa" (Anyone).
 * 7. Clique em Implantar, autorize as permissões pedidas.
 * 8. Copie a URL que termina em /exec — essa é a SHEET_API_URL
 *    que você vai colar no index.html.
 *
 * A aba da planilha precisa se chamar exatamente "Overrides"
 * com cabeçalhos na linha 1: Date | Key | Name | UpdatedAt
 */

const SHEET_NAME = "Overrides";

function doGet(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const result = {};
  for (let i = 1; i < data.length; i++) {
    const date = data[i][0];
    const key = data[i][1];
    const name = data[i][2];
    if (!date || !key) continue;
    if (!result[date]) result[date] = {};
    result[date][key] = name;
  }
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const body = JSON.parse(e.postData.contents);
  const date = body.date;
  const key = body.key;
  const name = body.name;
  const action = body.action; // "set" ou "delete"

  const data = sheet.getDataRange().getValues();
  let rowIndex = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === date && data[i][1] === key) {
      rowIndex = i + 1; // linhas do Sheets começam em 1
      break;
    }
  }

  if (action === "delete") {
    if (rowIndex > -1) sheet.deleteRow(rowIndex);
  } else {
    if (rowIndex > -1) {
      sheet.getRange(rowIndex, 3).setValue(name);
      sheet.getRange(rowIndex, 4).setValue(new Date());
    } else {
      sheet.appendRow([date, key, name, new Date()]);
    }
  }

  return ContentService.createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}
