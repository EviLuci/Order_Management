/** Global variables and constants */
// Variables
var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
var destinationsheet = SpreadsheetApp.openById(
  "ID of the destination spreadsheet"
);
var lock = LockService.getScriptLock();

// Consolidate sheet references into an object
var sheets = {
  database: spreadsheet.getSheetByName("Orders"),
  profile: destinationsheet.getSheetByName("Customer Details"),
  // cards
  ezeekard: destinationsheet.getSheetByName("Ezeekard"),
  company: destinationsheet.getSheetByName("Company Card"),
  google: destinationsheet.getSheetByName("Google Review"),
  trip: destinationsheet.getSheetByName("Trip Advisor"),
  insta: destinationsheet.getSheetByName("Instagram Card"),
  vCard: destinationsheet.getSheetByName("vCard"),
};

// Save data from the form in an array
var databaseSheet = sheets.database;
var appendedRow = databaseSheet.getLastRow();
var appendedRowLength = databaseSheet.getLastColumn();
var appendedData = databaseSheet
  .getRange(appendedRow, 1, 1, appendedRowLength)
  .getValues()[0];
// Get the current date
var currentDate = new Date();
var formattedDate = Utilities.formatDate(
  currentDate,
  Session.getScriptTimeZone(),
  "yyyy-MM-dd"
);

// Get Guid for the web url
var generated_guid = Utilities.getUuid();
var web_url = "Web URL";
var profile_url_with_guid = [web_url, generated_guid, "/"].join("");
var whatsapp = "wa.me/+977-";
var viber = "viber://chat?number=%2B977-";

/** End of global variables and constants */

/** Utility Functions */
/** Function to get sheet by name */
function getSheet(name) {
  var sheet = sheets[name];

  if (!sheet) {
    throw new Error("Sheet '" + name + "' not found.");
  }

  return sheet;
}

/** Function to get data from a target sheet and range. */
function getData(sheetName, ranges) {
  var sheet = getSheet(sheetName);

  if (!Array.isArray(ranges)) {
    ranges = [ranges]; // Convert single range to array for consistency
  }

  // Fetch data from each range and create an array of arrays
  var data = ranges.map(function (range) {
    var rangeValues = sheet.getRange(range).getValues();
    return rangeValues.flat(); // Flatten each range's values into a single array
  });

  return data;
}

/** Function to set data in a target sheet and range. */
function setData(targetSheetName, data, startRow, startColumn) {
  var targetSheet = getSheet(targetSheetName);

  // Ensure that 'data' is always a 2D array
  if (!Array.isArray(data[0])) {
    data = [data]; // Wrap in an array if it's a 1D array
  }

  // Determine the dimensions of the data
  var numRows = data.length;
  var numColumns = data[0].length;

  // Get the target range
  var targetRange = targetSheet.getRange(
    startRow,
    startColumn,
    numRows,
    numColumns
  );

  // Set values to the target range
  targetRange.setValues(data);

  // Return cell reference where data was populated
  return targetRange;
}

/** Function to generate link based on the range and set it on destination */
function generateLink(targetRange, displayText, destinationSheet, row, column) {
  // Get the spreadsheet URL of the target range's sheet
  var targetSpreadsheet = targetRange.getSheet().getParent();
  var sheetUrl = targetSpreadsheet.getUrl();

  // Construct the link URL pointing to the target range
  var linkUrl =
    sheetUrl +
    "#gid=" +
    targetRange.getSheet().getSheetId() +
    "&range=" +
    targetRange.getA1Notation();

  // Create the rich text link with the display text
  var richText = SpreadsheetApp.newRichTextValue()
    .setText(displayText)
    .setLinkUrl(linkUrl)
    .build();

  // Insert the rich text link into the specified cell
  var destinationCell = destinationSheet.getRange(row, column);
  destinationCell.setRichTextValue(richText);
}

/** Function to generate google map link using the fields given */
function generateGoogleMapsLink(street, city, state, zip, country) {
  // Ensure all fields are strings and filter out empty fields
  const addressParts = [street, city, state, zip, country]
    .map((part) => (part ? String(part).trim() : "")) // Convert to string and trim
    .filter((part) => part); // Remove empty strings

  // Join non-empty fields with commas
  const address = addressParts.join(", ");

  // URL-encode the address
  const encodedAddress = encodeURIComponent(address);

  return `https://www.google.com/maps/search/?api=1&query=${encodedAddress}`;
}

/** End of Utility Functions */

/** Function to save customer  */
function saveCustomerDetails() {
  // Get customer contact details
  var customerDetails = [
    formattedDate, // Date
    [appendedData[1], appendedData[2]].join(" "), // First and Last Name Joined
    appendedData[3], // Personal Number
    [appendedData[7], appendedData[8]].join(", "), // Address
    appendedData[15], // Company Name
    appendedData[0], // Products Selected
  ];
  // console.log(customerDetails);

  setData("profile", [customerDetails], sheets.profile.getLastRow() + 1, 1);
}

function test() {
  var ezeekardData = [].concat(
    appendedData.slice(1, 17), // Customer Data
    appendedData[19], // 1st social link
    appendedData.slice(31, 40), // Rest Links
    appendedData.slice(29, 31), // Communication Platforms
    [whatsapp, appendedData[29]].join(""), // whatsapp
    [viber, appendedData[30]].join(""), // viber
    appendedData[27], // Ezeekard Type
    appendedData[28], // Card Design
    // googleMapLink, // Google Map Link
    generated_guid, // Generated Guid
    profile_url_with_guid // Ezeekard URL
  );

  console.log(ezeekardData);
}

/** Function to save vCard Data */
function saveVCardData() {
  var selectedProducts = appendedData[0];
  if (
    selectedProducts.includes("Ezeekard") ||
    selectedProducts.includes("Company-Card")
  ) {
    var vCardData = [].concat(
      appendedData.slice(1, 12), // Customer Data
      appendedData[13], // Occupation
      appendedData[15], // Company Name
      appendedData[16], // Website
      profile_url_with_guid // Ezeekard URL
    );
    var targetRange = setData(
      "vCard",
      [vCardData],
      sheets.vCard.getLastRow() + 1,
      1
    );

    generateLink(
      targetRange,
      "vCard",
      sheets.profile,
      sheets.profile.getLastRow(),
      7
    );
  }
}

/** Function to manage order (saves data acording to product selected to their respective sheets) */
function manageOrderData() {
  // Parse the array like string of selected products passed by form into an actual JavaScript array
  var selectedProducts = JSON.parse(appendedData[0]);
  // console.log(selectedProducts);

  var googleMapLink = generateGoogleMapsLink(
    appendedData[7],
    appendedData[8],
    appendedData[9],
    appendedData[10],
    appendedData[11]
  );

  // Loop through each card type in the array
  for (var i = 0; i < selectedProducts.length; i++) {
    var cardType = selectedProducts[i];

    // Check the selected products and populate the respective cells with specified data
    if (cardType === "Ezeekard") {
      var ezeekardData = [].concat(
        appendedData.slice(1, 17), // Customer Data
        appendedData[19], // 1st social link
        appendedData.slice(31, 40), // Rest Links
        [whatsapp, appendedData[29]].join(""), // whatsapp
        [viber, appendedData[30]].join(""), // viber
        appendedData[27], // Ezeekard Type
        appendedData[28], // Card Design
        googleMapLink, // Google Map Link
        generated_guid, // Generated Guid
        profile_url_with_guid // Ezeekard URL
      );
      var targetRange = setData(
        "ezeekard",
        [ezeekardData],
        sheets.ezeekard.getLastRow() + 1,
        1
      );

      // Create a link in the Customer Details sheet
      generateLink(
        targetRange,
        cardType,
        sheets.profile,
        sheets.profile.getLastRow(),
        8
      );
    } else if (cardType === "Company-Card") {
      var companyCardData = [].concat(
        appendedData.slice(1, 19), // Customer Data
        appendedData[19], // 1st social link
        appendedData.slice(31, 40), // Rest links
        [whatsapp, appendedData[29]].join(""), // whatsapp
        [viber, appendedData[30]].join(""), // viber
        googleMapLink, // Google Map Link
        generated_guid, // Generated Guid
        profile_url_with_guid // Ezeekard URL
      );
      var targetRange = setData(
        "company",
        [companyCardData],
        sheets.company.getLastRow() + 1,
        1
      );

      // Create a link in the Customer Details sheet
      generateLink(
        targetRange,
        cardType,
        sheets.profile,
        sheets.profile.getLastRow(),
        9
      );
    } else if (cardType === "Google-Review-Card") {
      var googleReviewData = [].concat(
        appendedData.slice(1, 12), // Customer Data
        appendedData.slice(20, 22), // Company name and logo
        appendedData[23] // Google Review Link
      );
      var targetRange = setData(
        "google",
        [googleReviewData],
        sheets.google.getLastRow() + 1,
        1
      );

      // Create a link in the Customer Details sheet
      generateLink(
        targetRange,
        cardType,
        sheets.profile,
        sheets.profile.getLastRow(),
        10
      );
    } else if (cardType === "Trip-Advisor-Review-Card") {
      var tripAdvisorData = [].concat(
        appendedData.slice(1, 12), // Customer Data
        appendedData.slice(20, 22), // Company name and logo
        appendedData[22] // Trip Advisor Link
      );
      var targetRange = setData(
        "trip",
        [tripAdvisorData],
        sheets.trip.getLastRow() + 1,
        1
      );

      // Create a link in the Customer Details sheet
      generateLink(
        targetRange,
        cardType,
        sheets.profile,
        sheets.profile.getLastRow(),
        11
      );
    } else if (cardType === "Instagram-Card") {
      var instagramCardData = [].concat(
        appendedData.slice(1, 12), // Customer Data
        appendedData[24], // name
        appendedData[26], // logo
        appendedData[25] // link
      );
      var targetRange = setData(
        "insta",
        [instagramCardData],
        sheets.insta.getLastRow() + 1,
        1
      );

      // Create a link in the Customer Details sheet
      generateLink(
        targetRange,
        cardType,
        sheets.profile,
        sheets.profile.getLastRow(),
        12
      );
    }
  }
}

/** Function to handle new data and move it to respective sheets */
function handleNewData() {
  // Lock the script
  lock.waitLock(10000); // wait 10 seconds for others' use of the code section and lock to stop and then proceed

  try {
    /** Forminator form triggers the scripts two time for every form submission. so we need to delete a form submission */
    // Comparision data
    var newProfileDataCompare = [
      [appendedData[1], appendedData[2]].join(" "), // First and Last Name Joined
      appendedData[3], // Personal Number
      [appendedData[7], appendedData[8]].join(", "), // Address
      appendedData[15], // Company Name
      appendedData[0], // Product Selected
    ];

    // Check for duplication
    var profileSheet = sheets.profile;
    var profileLastRow = profileSheet.getLastRow();
    var isDuplicate = false;

    var profileLastData = profileSheet
      .getRange(profileLastRow, 2, 1, newProfileDataCompare.length)
      .getValues()[0];
    isDuplicate = newProfileDataCompare.every(function (value, index) {
      return value === profileLastData[index];
    });

    if (isDuplicate) {
      // console.log('Duplicate data found. Exiting script.');
      return;
    } else {
      saveCustomerDetails();
      saveVCardData();
      manageOrderData();
    }
  } catch (error) {
    Logger.log("Error: " + error.toString());
  } finally {
    // Release the lock
    lock.releaseLock();
  }
}
