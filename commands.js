

// This file will replace your VBA macro for Excel Online

Office.onReady(() => {
  console.log("Datanize Add-in is ready.");
});

// Sample function for the Syncgauge Matching button
async function syncgaugepart2Matching() {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem("Sheet1");
      const lastRow = sheet.getRange("C:C").getLastRow(); // last row with data in column C
      const range = sheet.getRange(`A1:BM${lastRow}`);
      range.load("values");
      await context.sync();

      const data = range.values;

      // Insert headers starting at column BA (53)
      const headers = [
        "Final_Set","Company","CompanyLinkedin","Position",
        "Duration","Profile Location","ExperienceLocation","Salesnav ID",
        "CompanySubstringOfProfileHeader","PositionLengthExceedsLimit",
        "Contains nonASCII chars","Status_Offshore","CountofPresent","Unverified"
      ];
      for (let i = 0; i < headers.length; i++) {
        data[0][52 + i] = headers[i];
      }

      // Loop through rows
      for (let row = 1; row < data.length; row++) {
        let targetCompany = data[row][3]; // Column D
        let durationSets = [];
        let companySets = [];
        let companyLinkedinSets = [];
        let positionSets = [];
        let locationSets = [];

        // Read sets
        for (let i = 0; i < 7; i++) {
          companySets[i] = data[row][17 + i * 5];       // Column R
          companyLinkedinSets[i] = data[row][18 + i * 5]; // Column S
          positionSets[i] = data[row][19 + i * 5];      // Column T
          durationSets[i] = data[row][20 + i * 5];      // Column U
          locationSets[i] = data[row][21 + i * 5];      // Column V
        }

        // TODO: replicate VBA logic for matching "Present" / most recent duration
        // TODO: clean company names, compute SalesNav ID, check non-ASCII, etc.

        // Example: write target company to BA
        data[row][53] = "Set1";  // Just placeholder
        data[row][54] = targetCompany;
      }

      range.values = data;
      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}

// Export functions for manifest
if (typeof module !== "undefined") {
  module.exports = {
    syncgaugepart2Matching
  };
}
