Office.onReady(() => {
    console.log("Datanize Add-in Ready");
});

// Example function for ribbon button
function syncgaugepart2Matching(event) {
    Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        sheet.getRange("A1").values = [["Syncgauge Matching ran"]];
        await context.sync();
    });
    event.completed();
}

// Add more functions as needed, e.g.:
// BuildTemplateAndRunChecks(event), CategorizeEmployeesAndRevenue(event), etc.
