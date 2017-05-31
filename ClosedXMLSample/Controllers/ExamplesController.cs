using System;
using System.IO;
using System.Web.Mvc;
using ClosedXML.Excel;

namespace ClosedXMLSample.Controllers
{
    public class ExamplesController : BaseController
    {
        /// <summary>
        /// https://github.com/ClosedXML/ClosedXML/wiki/Hello-World
        /// </summary>
        /// <returns></returns>
        public ActionResult HelloWorld()
        {
            GetInstance("Sample Sheet", out XLWorkbook wb, out IXLWorksheet ws);
            ws.Cell("A1").SetValue(1).CellBelow().SetValue(1);
            ws.Cell("A1").Value = "Hello World!";
            //wb.SaveAs("HelloWorld.xlsx");

            return ExportExcel(wb, "HelloWorld");
        }

        /// <summary>
        /// https://github.com/ClosedXML/ClosedXML/wiki/Basic-Table
        /// </summary>
        /// <returns></returns>
        public ActionResult BasicTable()
        {
            GetInstance("Contacts", out XLWorkbook wb, out IXLWorksheet ws);
            ws.Cell("A1").SetValue(1).CellBelow().SetValue(1);

            // Title
            ws.Cell("B2").Value = "Contacts";

            // First Names
            ws.Cell("B3").Value = "FName";
            ws.Cell("B4").Value = "John";
            ws.Cell("B5").Value = "Hank";
            ws.Cell("B6").Value = "Dagny";

            // Last Names
            ws.Cell("C3").Value = "LName";
            ws.Cell("C4").Value = "Galt";
            ws.Cell("C5").Value = "Rearden";
            ws.Cell("C6").Value = "Taggart";

            // Boolean
            ws.Cell("D3").Value = "Outcast";
            ws.Cell("D4").Value = true;
            ws.Cell("D5").Value = false;
            ws.Cell("D6").Value = false;

            // DateTime
            ws.Cell("E3").Value = "DOB";
            ws.Cell("E4").Value = new DateTime(1919, 1, 21);
            ws.Cell("E5").Value = new DateTime(1907, 3, 4);
            ws.Cell("E6").Value = new DateTime(1921, 12, 15);

            // Numeric
            ws.Cell("F3").Value = "Income";
            ws.Cell("F4").Value = 2000;
            ws.Cell("F5").Value = 40000;
            ws.Cell("F6").Value = 10000;

            // From worksheet
            var rngTable = ws.Range("B2:F6");

            // From another range
            var rngDates = rngTable.Range("D3:D5");
            var rngNumbers = rngTable.Range("E3:E5");

            // Using OpenXML's predefined formats
            rngDates.Style.NumberFormat.NumberFormatId = 15;

            // Using a custom format
            rngNumbers.Style.NumberFormat.Format = "$ #,##0";


            // The address is relative to rngTable (NOT the worksheet)
            var rngHeaders = rngTable.Range("A2:E2");
            rngHeaders.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            rngHeaders.Style.Font.Bold = true;
            rngHeaders.Style.Fill.BackgroundColor = XLColor.Aqua;

            rngTable.Style.Border.BottomBorder = XLBorderStyleValues.Thin;

            rngTable.Cell(1, 1).Style.Font.Bold = true;
            rngTable.Cell(1, 1).Style.Fill.BackgroundColor = XLColor.CornflowerBlue;
            rngTable.Cell(1, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

            // We could've also used: rngTable.Range("A1:E1").Merge()
            rngTable.Row(1).Merge();

            //Add a thick outside border
            rngTable.Style.Border.OutsideBorder = XLBorderStyleValues.Thick;

            // You can also specify the border for each side with:
            // rngTable.FirstColumn().Style.Border.LeftBorder = XLBorderStyleValues.Thick;
            // rngTable.LastColumn().Style.Border.RightBorder = XLBorderStyleValues.Thick;
            // rngTable.FirstRow().Style.Border.TopBorder = XLBorderStyleValues.Thick;
            // rngTable.LastRow().Style.Border.BottomBorder = XLBorderStyleValues.Thick;

            ws.Columns(2, 6).AdjustToContents();

            //wb.SaveAs("BasicTable.xlsx");

            return ExportExcel(wb, "BasicTable");

        }

        /// <summary>
        /// https://github.com/ClosedXML/ClosedXML/wiki/Showcase
        /// </summary>
        /// <returns></returns>
        public ActionResult Showcase()
        {
            GetInstance("Contacts", out XLWorkbook wb, out IXLWorksheet ws);
            ws.Cell("A1").SetValue(1).CellBelow().SetValue(1);

            #region Adding data text

            // Title
            ws.Cell("B2").Value = "Contacts";

            // First Name
            ws.Cell("B3").Value = "FName";
            ws.Cell("B4").Value = "John";
            ws.Cell("B5").Value = "Hank";
            // Another way to set the value
            ws.Cell("B6").SetValue("Dagny");

            // Last Names
            ws.Cell("C3").Value = "LName";
            ws.Cell("C4").Value = "Galt";
            ws.Cell("C5").Value = "Rearden";
            // Another way to set the value
            ws.Cell("C6").SetValue("Taggart");

            // Boolean
            ws.Cell("D3").Value = "Outcast";
            ws.Cell("D4").Value = true;
            ws.Cell("D5").Value = false;
            // Another way to set the value
            ws.Cell("D6").SetValue(false);

            // DateTime
            ws.Cell("E3").Value = "DOB";
            ws.Cell("E4").Value = new DateTime(1919, 1, 21);
            ws.Cell("E5").Value = new DateTime(1907, 3, 4);
            // Another way to set the value
            ws.Cell("E6").SetValue(new DateTime(1921, 12, 15));

            // Numeric
            ws.Cell("F3").Value = "Income";
            ws.Cell("F4").Value = 2000;
            ws.Cell("F5").Value = 40000;
            // Another way to set the value
            ws.Cell("F6").SetValue(10000);

            #endregion

            #region Defining ranges

            // From worksheet
            var rngTable = ws.Range("B2:F6");

            // From another range
            // The address is relative to rngTable (NOT the worksheet)
            // 這裡的定位是以 rngTable 來計算，非 worksheet 上的 A?, B? 來計算
            var rngDates = rngTable.Range("D3:D5");
            // The address is relative to rngTable (NOT the worksheet)
            var rngNumbers = rngTable.Range("E3:E5");

            // Formatting dates and numbers
            // Using a OpenXML's predefined formats
            rngDates.Style.NumberFormat.NumberFormatId = 15;
            // Using a custom format
            rngNumbers.Style.NumberFormat.Format = "$ #,##0";

            // Format title cell in one shot
            rngTable.FirstCell().Style
                .Font.SetBold()
                .Fill.SetBackgroundColor(XLColor.CornflowerBlue)
                .Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);

            // We could've also used:
            // rngTable.Range("A1:E1").Merge() or rngTable.Row(1).Merge()
            rngTable.FirstRow().Merge();
            #endregion

            // Formatting headers
            // The address is relative to rngTable (NOT the worksheet)
            var rngHeaders = rngTable.Range("A2:E2");
            rngHeaders.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            rngHeaders.Style.Font.Bold = true;
            rngHeaders.Style.Font.FontColor = XLColor.DarkBlue;
            rngHeaders.Style.Fill.BackgroundColor = XLColor.Aqua;

            // Create an Excel table with the data portion
            // 由 ws 定義資料範圍
            var rngData = ws.Range("B3:F6");
            var excelTable = rngData.CreateTable();

            // Add the totals row
            excelTable.ShowTotalsRow = true;
            // Put the average on the field "Income"
            // Notice how we're calling the cell by the column name
            excelTable.Field("Income").TotalsRowFunction = XLTotalsRowFunction.Sum;
            // Put a label on the totals cell of the field "DOB"
            excelTable.Field("DOB").TotalsRowLabel = "Sum:";


            // Add thick borders
            // Add thick borders to the contents of our spreadsheet
            ws.RangeUsed().Style.Border.OutsideBorder = XLBorderStyleValues.Thick;

            // You can also specify the border for each side:
            // contents.FirstColumn().Style.Border.LeftBorder = XLBorderStyleValues.Thick;
            // contents.LastColumn().Style.Border.RightBorder = XLBorderStyleValues.Thick;
            // contents.FirstRow().Style.Border.TopBorder = XLBorderStyleValues.Thick;
            // contents.LastRow().Style.Border.BottomBorder = XLBorderStyleValues.Thick;

            // Adjust column widths to their content
            ws.Columns().AdjustToContents();
            // You can also specify the range of columns to adjust, e.g.
            // ws.Columns(2, 6).AdjustToContents(); or ws.Columns("2-6").AdjustToContents();

            // Saving the workbook
            //wb.SaveAs("Showcase.xlsx");

            // use MVC output file
            return ExportExcel(wb, "Showcase");
        }
    }
}