using ClosedXML.Excel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

if(false){
// Step 1: Create a new workbook
  using (var workbook = new XLWorkbook())
  {
    // Step 2: Add a worksheet to the workbook
    var worksheet = workbook.AddWorksheet("Sheet1");

    // Step 3: Populate the worksheet with data
    worksheet.Cell(1, 1).Value = "Name";
    worksheet.Cell(1, 2).Value = "Age";
    worksheet.Cell(1, 3).Value = "City";

    worksheet.Cell(2, 1).Value = "Alice";
    worksheet.Cell(2, 2).Value = 25;
    worksheet.Cell(2, 2).Style.NumberFormat.SetNumberFormatId((int) XLPredefinedFormat.Number.Integer);
    worksheet.Cell(2, 3).Value = "New York";

    worksheet.Cell(3, 1).Value = "Bob";
    worksheet.Cell(3, 2).Value = 30;
    worksheet.Cell(3, 2).Style.NumberFormat.SetNumberFormatId((int) XLPredefinedFormat.Number.Integer);
    worksheet.Cell(3, 3).Value = "Los Angeles";

    worksheet.Cell(4, 1).Value = "Charlie";
    worksheet.Cell(4, 2).Value = 35;
    worksheet.Cell(4, 2).Style.NumberFormat.SetNumberFormatId((int) XLPredefinedFormat.Number.Integer);
    worksheet.Cell(4, 3).Value = "Chicago";

    // Step 4: Save the workbook to a file
    workbook.SaveAs("output.xlsx");
  }

  Console.WriteLine("Excel file created successfully!");
}

Console.WriteLine("Hello world");

if(true){
  using(SpreadsheetDocument document = SpreadsheetDocument.Create("output-plain.xlsx", SpreadsheetDocumentType.Workbook))
  {
    // Step 2: Add a WorkbookPart to the document
    WorkbookPart workbookPart = document.AddWorkbookPart();
    workbookPart.Workbook = new Workbook();

    // Step 3: Add a WorksheetPart to the WorkbookPart
    WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
    worksheetPart.Worksheet = new Worksheet(new SheetData());

    // Step 4: Add Sheets to the Workbook
    Sheets sheets = document.WorkbookPart.Workbook.AppendChild(new Sheets());

    // Step 5: Append a new worksheet and associate it with the workbook
    Sheet sheet = new Sheet()
    {
      Id = document.WorkbookPart.GetIdOfPart(worksheetPart),
         SheetId = 1,
         Name = "Sheet1"
    };
    sheets.Append(sheet);
            
    // Save the workbook
    workbookPart.Workbook.Save();

  }
}
