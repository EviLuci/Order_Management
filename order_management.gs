/** Global variables and constants */
// Variables
var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
var destinationsheet = SpreadsheetApp.openById(
    "destinationsheet id from the url"
);
var lock = LockService.getScriptLock();

// Consolidate sheet references into an object
var sheets = {
    database: spreadsheet.getSheetByName("Sheet1"),
    profile: destinationsheet.getSheetByName("Profile"),
    digital: destinationsheet.getSheetByName("Digital Profile Card"),
    hybrid: destinationsheet.getSheetByName("Hybrid Card"),
    both: destinationsheet.getSheetByName("Both"),
    google: destinationsheet.getSheetByName("Google Review"),
    custom: destinationsheet.getSheetByName("Customized Card"),
};

// No of column of data from database to be managed
var ORDER_COLUMN_NUM = 25;
var ORDER_COLUMN_NUM_CUSTOM = 27;

/** End of global variables and constants */

/** Utility Functions */
// Function to get sheet by name
function getSheet(name) {
    var sheet = sheets[name];

    if (!sheet) {
        throw new Error("Sheet '" + name + "' not found.");
    }

    return sheet;
}

/**
 * Function to get data from a target sheet and range.
 * @param {string} sheetName - Name of the target sheet.
 * @param {string | string[]} ranges - Range(s) to retrieve data from.
 * @returns {any[][]} - Data fetched from the target sheet and range.
 */
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

/**
 * Function to set data in a target sheet and range.
 * @param {string} targetSheetName - Name of the target sheet.
 * @param {string[][] | any[][]} data - Data to be set in the target sheet.
 * @param {number} startRow - Starting row index in the target sheet.
 * @param {number} startColumn - Starting column index in the target sheet.
 */
function setData(targetSheetName, data, startRow, startColumn) {
    var targetSheet = getSheet(targetSheetName);

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
}

/** Function to handle new data and move it to respective sheets */
function handleNewData() {
    // Lock the script
    lock.waitLock(10000); // wait 10 seconds for others' use of the code section and lock to stop and then proceed

    try {
        var databaseSheet = sheets.database;
        var lastRow = databaseSheet.getLastRow();
        var lastRowData = databaseSheet
            .getRange(lastRow, 1, 1, ORDER_COLUMN_NUM)
            .getValues()[0];
        var lastRowDataCustom = databaseSheet
            .getRange(lastRow, 1, 1, ORDER_COLUMN_NUM_CUSTOM)
            .getValues()[0];

        // Columns for condition checking
        var conditionColumns = databaseSheet
            .getRange(lastRow, ORDER_COLUMN_NUM + 1, 1, 2)
            .getValues()[0];

        // Retrieve the additional columns for screenshots
        var additionalColumnsRange = databaseSheet.getRange(
            lastRow,
            ORDER_COLUMN_NUM_CUSTOM + 1,
            1,
            databaseSheet.getLastColumn() - ORDER_COLUMN_NUM_CUSTOM
        );
        var additionalColumnsData = additionalColumnsRange.getValues()[0];

        // Find the single non-empty value in the additional columns
        var paymentData = additionalColumnsData.find(function (cell) {
            return cell !== "";
        });

        // Get the current date
        var currentDate = new Date();
        var formattedDate = Utilities.formatDate(
            currentDate,
            Session.getScriptTimeZone(),
            "yyyy-MM-dd"
        );

        // Create a new data array for the profile sheet
        var newProfileData = [
            formattedDate, // Add the current date
            [lastRowData[0], lastRowData[1], lastRowData[2]].join(" "), // Concatenate the name
            lastRowDataCustom[4], // Phone No.
            lastRowDataCustom[25], // Card Design
            lastRowDataCustom[26], // Card Type
            paymentData || "", // The payment data from the additional columns
        ];

        // Comparision data
        var newProfileDataCompare = [
            [lastRowData[0], lastRowData[1], lastRowData[2]].join(" "), // Concatenate the name
            lastRowDataCustom[4], // Phone No.
            lastRowDataCustom[25], // Card Design
            lastRowDataCustom[26], // Card Type
            paymentData || "", // The payment data from the additional columns
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
            // Maintain Profile sheet without exception
            setData(
                "profile",
                [newProfileData],
                sheets.profile.getLastRow() + 1,
                1
            );

            // Conditions
            if (conditionColumns[0] == "Customized Design") {
                setData(
                    "custom",
                    [lastRowDataCustom],
                    sheets.custom.getLastRow() + 1,
                    1
                );
            } else {
                switch (conditionColumns[1]) {
                    case "Digital Profile Card":
                        setData(
                            "digital",
                            [lastRowData],
                            sheets.digital.getLastRow() + 1,
                            1
                        );
                        break;
                    case "Hybrid Card":
                        setData(
                            "hybrid",
                            [lastRowData],
                            sheets.hybrid.getLastRow() + 1,
                            1
                        );
                        break;
                    case "Both":
                        setData(
                            "both",
                            [lastRowData],
                            sheets.both.getLastRow() + 1,
                            1
                        );
                        break;
                    case "Google Review":
                        setData(
                            "google",
                            [lastRowData],
                            sheets.google.getLastRow() + 1,
                            1
                        );
                        break;
                    default:
                        break;
                }
            }
        }
    } finally {
        // Release the lock
        lock.releaseLock();
    }
}
