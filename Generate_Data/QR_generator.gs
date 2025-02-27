function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Custom Menu")
    .addItem("Show vCard", "showVCardPopup")
    .addItem("Generate QR", "openQrDialog")
    .addItem("Generate GUID", "showGUIDPopup")
    .addToUi();
}

/** Function to generate GUID */
function showGUIDPopup() {
  const guid = Utilities.getUuid();
  const htmlOutput = HtmlService.createHtmlOutput(
    `
    <div style="text-align: center;">
      <p>Your GUID:</p>
      <input type="text" value="${guid}" id="guid" readonly style="width: 90%; padding: 10px;"/>
      <button onclick="copyGUID()" style="margin-top: 10px;">Copy GUID</button>
    </div>
    <script>
      function copyGUID() {
        const guidField = document.getElementById('guid');
        guidField.select();
        document.execCommand('copy');
        alert('GUID copied to clipboard');
      }
    </script>
  `
  )
    .setWidth(300)
    .setHeight(150);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, "Generated GUID");
}

/** Function to create vCard QR */
function showVCardPopup() {
  const htmlOutput = HtmlService.createHtmlOutputFromFile("vCardPopup")
    .setWidth(600)
    .setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, "vCard QR Code");
}

function getSelectedRowData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const range = sheet.getActiveRange();
  const rowData = sheet
    .getRange(range.getRow(), 1, 1, sheet.getLastColumn())
    .getValues()[0];

  const data = {
    fname: rowData[0],
    lname: rowData[1],
    personalphone: rowData[2],
    workphone: rowData[3],
    personalemail: rowData[4],
    workemail: rowData[5],
    street: rowData[6],
    city: rowData[7],
    state: rowData[8],
    postalcode: rowData[9],
    country: rowData[10],
    position: rowData[11],
    company: rowData[12],
    website: rowData[13],
    ezkwebsite: rowData[14],
    workphone2: rowData[15],
  };

  return data;
}

/** Function to create QR from the URL */
function openQrDialog() {
  const htmlOutput = HtmlService.createHtmlOutputFromFile("QRGenerator")
    .setWidth(600)
    .setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, "QR Code Generator");
}

function getSelectedRowDataForURL() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const range = sheet.getActiveRange();
  const row = range.getRow();
  const lastColumn = sheet.getLastColumn();

  const url = sheet.getRange(row, lastColumn).getValue(); // Get URL from the last column
  const firstName = sheet.getRange(row, 1).getValue(); // Get value from the first column
  const lastName = sheet.getRange(row, 2).getValue(); // Get value from the second column

  const fileName = `${firstName}_${lastName}`;
  console.log("URL:", url);
  console.log("File Name:", fileName);
  return { url, fileName };
}
