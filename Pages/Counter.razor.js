"use strict";

export async function getOfficeContextAsync() {
    return window.OfficeContext;
}

export async function incrementCountAsync(count) {
	console.log("Enter: incrementCountAsync()");
    // Get the current context
    const officeContext = await getOfficeContextAsync();
    console.log(`officeContext=${ officeContext}`);
    // make sure we are in Excel
    if (officeContext !== Office.HostType.Excel) { return; }

    await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const selectedRange = context.workbook.getSelectedRange();
        // Get the first cell of the range (row 0, column 0)
        const firstCell = selectedRange.getCell(0, 0);
        // Load the address of the first cell
        firstCell.load("address");
        console.log("First cell:", firstCell);
        await context.sync();

        // get the very first cell in the selected range
        const firstCellAddress = firstCell.address.split("!")[1]; // Remove the sheet name from the address
        console.log("First cell address:", firstCellAddress);

        // intsert value
        const range = sheet.getRange(firstCellAddress);
        range.values = count;

		// select the cell below
		const cellBelow = range.getCell(1, 0);
		cellBelow.select();
        //range.format.autofitColumns();
        await context.sync();
    });
}

export async function registerHandlersAsync()
{
	console.log("Enter: registerHandlersAsync()");
	// Get the current context
	const officeContext = await getOfficeContextAsync();
	console.log(`officeContext=${ officeContext}`);
	// make sure we are in Excel
	if (officeContext !== Office.HostType.Excel) { return; }

	await Excel.run(async (context) => {
		const sheet = context.workbook.worksheets.getActiveWorksheet();
		sheet.onChanged.add(onChangedHandler);
		sheet.onSelectionChanged.add(onSelectionChangedHandler);
	});
}

const onChangedHandler = async (event) => {
	console.log("Enter: onChangedHandler()");
	await Excel.run(async (context) => {
		await context.sync();
		console.log("Change type of event: " + event.changeType);
		console.log("Address of event: " + event.address);
		console.log("Source of event: " + event.source);
		// get value of cell
		const sheet = context.workbook.worksheets.getActiveWorksheet();
		const range = sheet.getRange(event.address);
		range.load("values");
		await context.sync();
		console.log("Value of cell: " + range.values);
		// invoke the C# method
		window.dotNetHelper?.invokeMethodAsync('OnChanged', event.address, `${range.values}`);
	});
	await context.sync();
	console.log("Event handler successfully registered for onChanged event in the worksheet.");
};

const onSelectionChangedHandler = async (event) => {
	console.log("Enter: onSelectionChangedHandler()");
	await Excel.run(async (context) => {
		await context.sync();
		console.log("Change type of event: " + event.changeType);
		console.log("Address of event: " + event.address);
		console.log("Source of event: " + event.source);
		console.assert(window.dotNetHelper);
		// get value of cell
		const sheet = context.workbook.worksheets.getActiveWorksheet();
		const range = sheet.getRange(event.address);
		range.load("values");
		await context.sync();
		console.log("Value of cell: " + range.values);
		// invoke the C# method
		window.dotNetHelper?.invokeMethodAsync('OnSelectionChanged', event.address, `${range.values}`);
	});
	await context.sync();
	console.log("Event handler successfully registered for onSelectionChanged event in the worksheet.");
};
