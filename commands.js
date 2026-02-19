Office.onReady(() => {});

function syncgaugepart2Matching(event) { runExample("Syncgauge Matching", event); }
function BuildTemplateAndRunChecks(event) { runExample("Syncgauge Templating", event); }
function CategorizeEmployeesAndRevenue(event) { runExample("Categorize Employees & Revenue", event); }
function ContactsOffshoreStatus_SalesNavPull(event) { runExample("SalesNav Pull", event); }
function FindDateMatchAndFormatToTMUpdateSheet(event) { runExample("Format To TM Update Sheet", event); }
function CleanWebsiteAndDomain(event) { runExample("Clean Website & Domain", event); }
function ExportEmailsToTxt(event) { runExample("Export Emails to TXT", event); }
function GetEmailsFromWebsites(event) { runExample("Get Emails From Websites", event); }
function AppendEducationWideWithHeaders(event) { runExample("Append Education Wide", event); }
function ConcatenateByContactID_AllValues(event) { runExample("Concatenate By ContactID", event); }
function CreateConcatenatedJobHistory(event) { runExample("Create Concatenated Job History", event); }
function HighlightInvalidData_Dynamic(event) { runExample("QA New Companies", event); }

function runExample(text, event) {
    Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        sheet.getRange("A1").values = [[text + " executed"]];
        await context.sync();
    });
    event.completed();
}
