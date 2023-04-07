"use strict";

export async function getOfficeContextAsync() {
    return window.OfficeContext;
}

export async function getWeatherDataAsync(recordset) {
    console.log("Enter: getWeatherData()");
    for (const header of recordset.headers) {
        console.log(`header=${header}`);
    }
    for (const row of recordset.records) {
        console.log(JSON.stringify(row));
    }
}

export async function insertWeatherDataAsync(recordset) {
    // Get the current context
    const officeContext = await getOfficeContextAsync();
    console.log(`officeContext=${officeContext}`);
    // make sure we are in Excel
    if (officeContext !== Office.HostType.Excel) { return; }

    await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const selectedRange = context.workbook.getSelectedRange();
        // Get the first cell of the range (row 0, column 0)
        const firstCell = selectedRange.getCell(0, 0);
        // Load the address of the first cell
        firstCell.load("address");
        await context.sync();
        // get the very first cell in the selected range
        const firstCellAddress = firstCell.address.split("!")[1]; // Remove the sheet name from the address
        console.log("First cell address:", firstCellAddress);

        const headers = recordset.headers;
        const records = recordset.records;

        const rowCount = 1; // Change this to the desired number of rows
        const columnCount = headers.length;
        console.log(`headers.length=${headers.length}`);

        const hrange = sheet.getRange(firstCellAddress).getResizedRange(rowCount - 1, columnCount - 1);
        //range.select(); // This line selects the range, remove this line if you don't want to select it
        hrange.values = [headers];
        hrange.format.font.bold = true;
        hrange.format.fill.color = "#D9E1F2"; // light blue
        hrange.format.borders.color = "#FF0000"; // TODO

        // loop on the forecast records and add them to the sheet
        console.log(`records.length=${records.length}`);
        const rows = [];
        const propertyNames = Object.getOwnPropertyNames(records[0]);
        for (const record of records) {
            const row = [];
            for (const propertyName of propertyNames) {
                var value = record[propertyName];
                row.push(value);
            }
            rows.push(row);
        }

        const drange = sheet.getRange(firstCellAddress).getCell(1, 0).getResizedRange(records.length - 1, columnCount - 1);
        //drange.select(); // This line selects the range, remove this line if you don't want to select it
        drange.values = rows;
        drange.format.font.bold = false;
        drange.format.autofitColumns();

        await context.sync();
    });
}
