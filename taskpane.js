let editedRows = new Set();

Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getActiveWorksheet();

            // Track changes in the worksheet
            sheet.onChanged.add(handleChange);

            await context.sync();
        });
    }
});

// Handle changes
async function handleChange(event) {
    await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();

        // event.address is the changed range like "B2:D4"
        const rangeAddress = event.address;
        const rows = parseRowsFromAddress(rangeAddress);

        rows.forEach(r => editedRows.add(r));

        // Update display cell in sheet
        const displayCell = sheet.getRange("aa1"); // You can pick any cell
        displayCell.values = [[editedRows.size]];

        // Also update task pane UI
        document.getElementById("rowCount").textContent = editedRows.size;

        await context.sync();
    });
}

// Helper: parse row numbers from range address
function parseRowsFromAddress(address) {
    // Example: "B2:D4" → rows 2,3,4
    const parts = address.split(":");
    const startRow = parseInt(parts[0].match(/\d+/)[0]);
    const endRow = parts[1] ? parseInt(parts[1].match(/\d+/)[0]) : startRow;

    let rowNums = [];
    for (let i = startRow; i <= endRow; i++) {
        rowNums.push(i);
    }
    return rowNums;
}
