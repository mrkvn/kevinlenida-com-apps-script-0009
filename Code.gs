function doGet() {
    return HtmlService.createTemplateFromFile("index")
        .evaluate()
        .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

/**
 * Returns a UUID.
 * @return UUID.
 * @customfunction
 */
function UUID() {
    return Utilities.getUuid();
}

function add(sheetName, data) {
    if (!data.id) {
        data["id"] = UUID();
    }
    
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);

    // Get all data in the sheet
    const sheetData = sheet.getDataRange().getValues();

    // Get the header row
    const headers = sheetData[0];

    // Find the index of the 'id' column
    const idColumnIndex = headers.indexOf('id');

    // Create a new row array with the same length as headers
    const newRow = new Array(headers.length).fill("");

    // Populate the new row with data from the object
    for (let key in data) {
        const columnIndex = headers.indexOf(key);
        if (columnIndex !== -1) {
            newRow[columnIndex] = data[key];
        }
    }

    // Append a blank row first to get the last row index
    sheet.appendRow(new Array(headers.length).fill(""));

    // Get the last row index
    const newRowIndex = sheet.getLastRow() + 1;

    // Format the 'id' column cell of the new row as plain text
    if (idColumnIndex !== -1) {
        const idCell = sheet.getRange(newRowIndex, idColumnIndex + 1);
        idCell.setNumberFormat('@STRING@'); // Set cell to plain text format

        // Set the value of the ID cell
        idCell.setValue(data['id']);
    }

    // Populate the rest of the row with data from the object
    for (let key in data) {
        const columnIndex = headers.indexOf(key);
        if (columnIndex !== -1 && key !== 'id') {
            sheet.getRange(newRowIndex, columnIndex + 1).setValue(data[key]);
        }
    }

    return data["id"];
}

function fetchAll(sheetName) {
    const sheet =
        SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    // const range = sheet.getRange(1, 1, 10, sheet.getLastColumn()); // first 10 rows
    const sheetData = sheet.getDataRange().getValues();
    // Get the headers (first row)
    let headers = sheetData[0];

    // Convert rows to objects
    let data = [];
    for (let i = 1; i < sheetData.length; i++) {
        const row = sheetData[i];
        let rowObject = {};
        for (let j = 0; j < row.length; j++) {
            rowObject[headers[j]] = row[j];
        }
        data.push(rowObject);
    }
    return data; // Return the data for further use if needed
}

function deleteById(sheetName, id) {
    // Open the active spreadsheet and get the sheet
    const sheet =
        SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);

    // Get all data in the sheet
    const sheetData = sheet.getDataRange().getValues();

    // Loop through all the rows
    for (let i = 0; i < sheetData.length; i++) {
        // Check if the ID in the current row is equal to the given ID
        if (sheetData[i][0] == id) {
            // Assuming the ID is in the first column (index 0)
            // Delete the row (i+1 because rows are 1-indexed)
            sheet.deleteRow(i + 1);
            // Exit the loop once the row is deleted
            break;
        }
    }
}

function update(sheetName, data) {
    const sheet =
        SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);

    // Get all data in the sheet
    const sheetData = sheet.getDataRange().getValues();

    // Get the header row
    const headers = sheetData[0];

    // Find the index of the row with the matching ID
    let rowIndex = -1;
    for (let i = 1; i < sheetData.length; i++) {
        if (sheetData[i][headers.indexOf("id")] === data.id) {
            rowIndex = i;
            break;
        }
    }

    if (rowIndex === -1) {
        throw new Error("ID not found");
    }

    // Update the row with the new data
    for (let key in data) {
        let columnIndex = headers.indexOf(key);
        if (columnIndex !== -1) {
            sheet.getRange(rowIndex + 1, columnIndex + 1).setValue(data[key]);
        }
    }
}
