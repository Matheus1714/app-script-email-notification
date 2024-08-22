const SHEET_ID_WITH_NAME_AND_EMAIL =
  "1J_BdrH6S5xT3I929T2uhsTAt2j5IxckhFWiX2HdMDMY";

/**
 * Função para obter nomes e emails de uma planilha específica.
 *
 * @param {string} sheetId
 * @return {Array<{ name: string, email: string }>}
 */
function getNameAndEmail(sheetId) {
  const spreadsheet = SpreadsheetApp.openById(sheetId);
  const sheet = spreadsheet.getSheets()[0];
  const range = sheet.getDataRange().getValues();

  const data = [];
  for (let i = 1; i < range.length; i++) {
    const user = {};
    user.name = range[i][0];
    user.email = range[i][1];
    data.push(user);
  }
  return data;
}

/**
 * Função para obter o template do email.
 *
 * @return {string}
 */
function getTemplate() {
  return HtmlService.createHtmlOutputFromFile("Template").getContent();
}

/**
 * Função para enviar emails a uma lista de usuários.
 *
 */
function sendEmail() {
  const users = getNameAndEmail(SHEET_ID_WITH_NAME_AND_EMAIL);
  const template = getTemplate();

  users.forEach((user) => {
    const subject = "Email de Notificação Automática";
    const body = template?.replace("{name}", user.name);

    if (user.email?.length) {
      MailApp.sendEmail({
        to: user.email,
        subject,
        htmlBody: body,
      });
    }
  });
}
