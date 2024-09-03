using System;
using ClosedXML.Excel;

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
