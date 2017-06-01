using System;
using System.Linq;
using System.Web.Mvc;
using ClosedXML.Excel;

namespace ClosedXMLSample.Controllers
{
    public class RangesController : BaseController
    {
        /// <summary>
        /// https://github.com/ClosedXML/ClosedXML/wiki/Defining-Ranges
        /// </summary>
        /// <returns></returns>
        public ActionResult Defining()
        {
            GetInstance("Defining a Range", out XLWorkbook wb, out IXLWorksheet ws);

            // With a string
            var range1 = ws.Range("A1:B1");
            range1.Cell(1, 1).Value = "ws.Range(\"A1:B1\").Merge()";
            range1.Merge();

            // With two XLAddresses
            var range2 = ws.Range(ws.Cell(2, 1).Address, ws.Cell(2, 2).Address);
            range2.Cell(1, 1).Value = "ws.Range(ws.Cell(2, 1).Address, ws.Cell(2, 2).Address).Merge()";
            range2.Merge();

            // With two strings
            var range4 = ws.Range("A3", "B3");
            range4.Cell(1, 1).Value = "ws.Range(\"A3\", \"B3\").Merge()";
            range4.Merge();

            // With 4 points
            var range5 = ws.Range(4, 1, 4, 2);
            range5.Cell(1, 1).Value = "ws.Range(4, 1, 4, 2).Merge()";
            range5.Merge();

            ws.Column("A").AdjustToContents();

            //wb.SaveAs("DefiningRanges.xlsx");

            return ExportExcel(wb, "DefiningRanges");
        }

        /// <summary>
        /// https://github.com/ClosedXML/ClosedXML/wiki/Merging-Cells
        /// </summary>
        /// <returns></returns>
        public ActionResult MergingCells()
        {
            GetInstance("Merge Cells", out XLWorkbook wb, out IXLWorksheet ws);

            // Merge a row
            ws.Cell("B2").Value = "Merged Row(1) of Range (B2:D3)";
            ws.Range("B2:D3").Row(1).Merge();

            // Merge a column
            ws.Cell("F2").Value = "Merged Column(1) of Range (F2:G8)";
            ws.Cell("F2").Style.Alignment.WrapText = true;
            ws.Range("F2:G8").Column(1).Merge();

            // Merge a range
            ws.Cell("B4").Value = "Merged Range (B4:D6)";
            ws.Cell("B4").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            ws.Cell("B4").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            ws.Range("B4:D6").Merge();

            // Unmerging a range...
            ws.Cell("B8").Value = "Unmerged";
            ws.Range("B8:D8").Merge();
            ws.Range("B8:D8").Unmerge();

            //---
            //var mergedRange = ws.MergedRanges.First(r => r.Contains("B4"));
            //mergedRange.Style.Fill.BackgroundColor = XLColor.Red;
            //---

            //wb.SaveAs("MergeCells.xlsx");

            return ExportExcel(wb, "MergeCells");
        }

        /// <summary>
        /// https://github.com/ClosedXML/ClosedXML/wiki/Clearing-Ranges
        /// </summary>
        /// <returns></returns>
        public ActionResult Clearing()
        {
            GetInstance("Clearing Ranges", out XLWorkbook wb, out IXLWorksheet ws);

            foreach (var ro in Enumerable.Range(1, 10))
            {
                foreach (var co in Enumerable.Range(1, 10))
                {
                    var cell = ws.Cell(ro, co);
                    cell.Value = cell.Address.ToString();
                    cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    cell.Style.Fill.BackgroundColor = XLColor.Turquoise;
                    cell.Style.Font.Bold = true;
                }
            }

            // Clearing a range
            ws.Range("B1:C2").Clear();

            // Clearing a row in a range
            ws.Range("B4:C5").Row(1).Clear();

            // Clearing a column in a range
            ws.Range("E1:F4").Column(2).Clear();

            // Clear an entire row
            ws.Row(7).Clear();

            // Clear an entire column
            ws.Column("H").Clear();

            //wb.SaveAs("ClearingRanges.xlsx");

            return ExportExcel(wb, "ClearingRanges");
        }

        /// <summary>
        /// https://github.com/ClosedXML/ClosedXML/wiki/Deleting-Ranges
        /// </summary>
        /// <returns></returns>
        public ActionResult Deleting()
        {
            GetInstance("Deleting Ranges", out XLWorkbook wb, out IXLWorksheet ws);

            foreach (var ro in Enumerable.Range(1, 10))
            foreach (var co in Enumerable.Range(1, 10))
                ws.Cell(ro, co).Value = ws.Cell(ro, co).Address.ToString();

            // Delete range and shift cells up
            ws.Range("B4:C5").Delete(XLShiftDeletedCells.ShiftCellsUp);

            // Delete range and shift cells left
            ws.Range("D1:E3").Delete(XLShiftDeletedCells.ShiftCellsLeft);

            // Delete an entire row
            ws.Row(5).Delete();

            // Delete a row in a range, shift cells up
            ws.Range("A1:C4").Row(2).Delete(XLShiftDeletedCells.ShiftCellsUp);

            // Delete an entire column
            ws.Column(5).Delete();

            // Delete a column in a range, shift cells up
            ws.Range("A1:C4").Column(2).Delete(XLShiftDeletedCells.ShiftCellsLeft);

            //wb.SaveAs("DeletingRanges.xlsx");

            return ExportExcel(wb, "DeletingRanges");
        }

        /// <summary>
        /// https://github.com/ClosedXML/ClosedXML/wiki/Multiple-Ranges
        /// </summary>
        /// <returns></returns>
        public ActionResult Multiple()
        {
            GetInstance("Multiple Ranges", out XLWorkbook wb, out IXLWorksheet ws);

            // using multiple string range definitions
            ws.Ranges("A1:B2,C3:D4,E5:F6").Style.Fill.BackgroundColor = XLColor.Red;

            // using a single string separated by commas
            ws.Ranges("A5:B6,E1:F2").Style.Fill.BackgroundColor = XLColor.Orange;

            //wb.SaveAs("MultipleRanges.xlsx");

            return ExportExcel(wb, "MultipleRanges");
        }

        /// <summary>
        /// https://github.com/ClosedXML/ClosedXML/wiki/Shifting-Ranges
        /// </summary>
        /// <returns></returns>
        public ActionResult Shifting()
        {
            // TODO: Path
            string fileName = @"C:\Users\BruceChen\Downloads\BasicTable.xlsx";
            var wb = new XLWorkbook(fileName);
            var ws = wb.Worksheet(1);

            // Get a range object
            var rngHeaders = ws.Range("B3:F3");

            // Insert some rows/columns before the range
            ws.Row(1).InsertRowsAbove(2);
            ws.Column(1).InsertColumnsBefore(2);

            // Change the background color of the headers
            // Notice that rngHeaders point to the right place
            rngHeaders.Style.Fill.BackgroundColor = XLColor.LightSalmon;

            ws.Columns().AdjustToContents();

            //wb.SaveAs("ShiftingRanges.xlsx");

            return ExportExcel(wb, "ShiftingRanges");
        }

        /// <summary>
        /// https://github.com/ClosedXML/ClosedXML/wiki/Transpose-Ranges
        /// </summary>
        /// <returns></returns>
        public ActionResult Transpose()
        {
            // TODO: Path
            string fileName = @"C:\Users\BruceChen\Downloads\BasicTable.xlsx";

            var wb = new XLWorkbook(fileName);
            var ws = wb.Worksheet(1);
            var rngTable = ws.Range("B2:F6");

            // 行/列互換
            rngTable.Transpose(XLTransposeOptions.MoveCells);

            ws.Columns().AdjustToContents();

            //wb.SaveAs("TransposeRanges.xlsx");

            return ExportExcel(wb, "TransposeRanges");
        }

        /// <summary>
        /// https://github.com/ClosedXML/ClosedXML/wiki/Named-Ranges
        /// </summary>
        /// <returns></returns>
        public ActionResult Named()
        {
            GetInstance("Presentation", out XLWorkbook wb, out IXLWorksheet wsPresentation);

            //var wsPresentation = wb.Worksheets.Add("Presentation");
            var wsData = wb.Worksheets.Add("Data");

            // Fill up some data
            wsData.Cell(1, 1).Value = "Name";
            wsData.Cell(1, 2).Value = "Age";
            wsData.Cell(2, 1).Value = "Tom";
            wsData.Cell(2, 2).Value = 30;
            wsData.Cell(3, 1).Value = "Dick";
            wsData.Cell(3, 2).Value = 25;
            wsData.Cell(4, 1).Value = "Harry";
            wsData.Cell(4, 2).Value = 29;

            // Create a named range with the data:
            wsData.Range("A2:B4").AddToNamed("PeopleData"); // Default named range scope is Workbook

            // Let's use the named range in a formula:
            wsPresentation.Cell(1, 1).Value = "People Count:";
            wsPresentation.Cell(1, 2).FormulaA1 = "COUNT(PeopleData)";

            // Create a named range with worksheet scope:
            wsPresentation.Range("B1").AddToNamed("PeopleCount", XLScope.Worksheet);

            // Let's use the named range:
            wsPresentation.Cell(2, 1).Value = "Total:";
            wsPresentation.Cell(2, 2).FormulaA1 = "PeopleCount";

            // Copy the data in a named range:
            wsPresentation.Cell(4, 1).Value = "People Data:";
            wsPresentation.Cell(5, 1).Value = wb.Range("PeopleData");

            //===
            // For the Excel geeks out there who actually know about
            // named ranges with relative addresses, you can
            // create such a thing with the following methods:

            // The following creates a relative named range pointing to the same row
            // and one column to the right. For example if the current cell is B4
            // relativeRange1 will point to C4.
            wsPresentation.NamedRanges.Add("relativeRange1", "Presentation!B1");

            // The following creates a ralative named range pointing to the same row
            // and one column to the left. For example if the current cell is D2
            // relativeRange2 will point to C2.
            wb.NamedRanges.Add("relativeRange2", "Presentation!XFD1");

            // Explanation: The address of a relative range always starts at A1
            // and moves from then on. To get the desired relative range just
            // add or subtract the required rows and/or columns from A1.
            // Column -1 = XFD, Column -2 = XFC, etc.
            // Row -1 = 1048576, Row -2 = 1048575, etc.
            //===

            wsData.Columns().AdjustToContents();
            wsPresentation.Columns().AdjustToContents();

            //wb.SaveAs("NamedRanges.xlsx");

            return ExportExcel(wb, "NamedRanges");
        }

        /// <summary>
        /// https://github.com/ClosedXML/ClosedXML/wiki/Accessing-Named-Ranges
        /// </summary>
        /// <returns></returns>
        public ActionResult AccessingNamed()
        {
            GetInstance("SeeWiki", out XLWorkbook wb, out IXLWorksheet ws);
            return ExportExcel(wb, "SeeWiki");
        }

        /// <summary>
        /// https://github.com/ClosedXML/ClosedXML/wiki/Copying-Ranges
        /// </summary>
        /// <returns></returns>
        public ActionResult Copying()
        {
            // TODO: Path
            string fileName = @"C:\Users\BruceChen\Downloads\BasicTable.xlsx";
            var wb = new XLWorkbook(fileName);
            var ws = wb.Worksheet(1);

            // Define a range with the data
            var firstTableCell = ws.FirstCellUsed();
            var lastTableCell = ws.LastCellUsed();
            var rngData = ws.Range(firstTableCell.Address, lastTableCell.Address);

            // Copy the table to another worksheet
            var wsCopy = wb.Worksheets.Add("Contacts Copy");
            wsCopy.Cell(1, 1).Value = rngData;

            //wb.SaveAs("CopyingRanges.xlsx");

            return ExportExcel(wb, "CopyingRanges");
        }

        public ActionResult Tables()
        {
            // TODO: Path
            string fileName = @"C:\Users\BruceChen\Downloads\BasicTable.xlsx";
            var wb = new XLWorkbook(fileName);
            var ws = wb.Worksheet(1);

            //---
            // Table 1


            var firstCell = ws.FirstCellUsed();
            var lastCell = ws.LastCellUsed();
            var range = ws.Range(firstCell.Address, lastCell.Address);
            // Deleting the "Contacts" header (we don't need it for our purposes)
            range.Row(1).Delete();

            // We want to use a theme for table, not the hard coded format of the BasicTable
            range.Clear(XLClearOptions.Formats);
            // Put back the date and number formats
            range.Column(4).Style.NumberFormat.NumberFormatId = 15;
            range.Column(5).Style.NumberFormat.Format = "$ #,##0";

            // You can also use range.AsTable() if you want to
            // manipulate the range as a table but don't want
            // to create the table in the worksheet.
            var table = range.CreateTable();

            // Let's activate the Totals row and add the sum of Income
            table.ShowTotalsRow = true;
            table.Field("Income").TotalsRowFunction = XLTotalsRowFunction.Sum;
            // Just for fun let's add the text "Sum Of Income" to the totals row
            table.Field(0).TotalsRowLabel = "Sum Of Income";

            //---
            // Table 2

            // Copy all the headers
            int columnWithHeaders = lastCell.Address.ColumnNumber + 2; // H
            int currentRow = table.RangeAddress.FirstAddress.RowNumber;
            ws.Cell(currentRow, columnWithHeaders).Value = "Table Headers";
            foreach (var cell in table.HeadersRow().Cells())
            {
                currentRow++;
                ws.Cell(currentRow, columnWithHeaders).Value = cell.Value;
            }

            // Format the headers as a table with a different style and no autofilters
            var htFirstCell = ws.Cell(table.RangeAddress.FirstAddress.RowNumber, columnWithHeaders);
            var htLastCell = ws.Cell(currentRow, columnWithHeaders);
            var headersTable = ws.Range(htFirstCell, htLastCell).CreateTable("Headers");
            headersTable.Theme = XLTableTheme.TableStyleLight10;
            headersTable.ShowAutoFilter = false;

            // Add a custom formula to the headersTable
            headersTable.ShowTotalsRow = true;
            headersTable.Field(0).TotalsRowFormulaA1 = "CONCATENATE(\"Count: \", CountA(Headers[Table Headers]))";

            //---
            // Table 3

            // Copy the names
            int columnWithNames = columnWithHeaders + 2;
            currentRow = table.RangeAddress.FirstAddress.RowNumber; // reset the currentRow
            ws.Cell(currentRow, columnWithNames).Value = "Names";
            foreach (var row in table.DataRange.Rows())
            {
                currentRow++;
                var fName = row.Field("FName").GetString(); // Notice how we're calling the cell by field name
                var lName = row.Field("LName").GetString(); // Notice how we're calling the cell by field name
                var name = $"{fName} {lName}";
                ws.Cell(currentRow, columnWithNames).Value = name;
            }

            // Format the names as a table with a different style and no autofilters
            var ntFirstCell = ws.Cell(table.RangeAddress.FirstAddress.RowNumber, columnWithNames);
            var ntLastCell = ws.Cell(currentRow, columnWithNames);
            var namesTable = ws.Range(ntFirstCell, ntLastCell).CreateTable();
            namesTable.Theme = XLTableTheme.TableStyleLight12;
            namesTable.ShowAutoFilter = false;

            ws.Columns().AdjustToContents();
            ws.Columns("A,G,I").Width = 3;

            //wb.SaveAs("UsingTables.xlsx");

            return ExportExcel(wb, "UsingTables");
        }

        [NonAction]
        public ActionResult CodeBase()
        {
            GetInstance("Collections", out XLWorkbook wb, out IXLWorksheet ws);
            return ExportExcel(wb, "Collections");
        }
    }
}