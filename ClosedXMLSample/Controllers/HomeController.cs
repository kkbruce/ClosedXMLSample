using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using ClosedXML.Excel;

namespace ClosedXMLSample.Controllers
{
    public class HomeController : Controller
    {
        // GET: Home
        public ActionResult Index()
        {
            return View();
        }

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

        /// <summary>
        /// https://github.com/ClosedXML/ClosedXML/wiki/Data-Types
        /// </summary>
        /// <returns></returns>
        public ActionResult DataTypes()
        {
            GetInstance("Data Types", out XLWorkbook wb, out IXLWorksheet ws);
            ws.Cell("A1").SetValue(1).CellBelow().SetValue(1);

            var co = 2;
            var ro = 1;

            ws.Cell(++ro, co).Value = "Plain Text:";
            ws.Cell(ro, co + 1).Value = "Hello World.";

            ws.Cell(++ro, co).Value = "Plain Date:";
            ws.Cell(ro, co + 1).Value = new DateTime(2010, 9, 2);

            ws.Cell(++ro, co).Value = "Plain DateTime:";
            ws.Cell(ro, co + 1).Value = new DateTime(2010, 9, 2, 13, 45, 22);

            ws.Cell(++ro, co).Value = "Plain Boolean:";
            ws.Cell(ro, co + 1).Value = true;

            ws.Cell(++ro, co).Value = "Plain Number:";
            ws.Cell(ro, co + 1).Value = 123.45;

            ws.Cell(++ro, co).Value = "TimeSpan:";
            ws.Cell(ro, co + 1).Value = new TimeSpan(33, 45, 22);

            ro++;

            // 明確指定為「Text」
            ws.Cell(++ro, co).Value = "Explicit Text:";
            ws.Cell(ro, co + 1).Value = "'Hello World.";

            ws.Cell(++ro, co).Value = "Date as Text:";
            ws.Cell(ro, co + 1).Value = "'" + new DateTime(2010, 9, 2).ToString();

            ws.Cell(++ro, co).Value = "DateTime as Text:";
            ws.Cell(ro, co + 1).Value = "'" + new DateTime(2010, 9, 2, 13, 45, 22).ToString();

            ws.Cell(++ro, co).Value = "Boolean as Text:";
            ws.Cell(ro, co + 1).Value = "'" + true.ToString();

            ws.Cell(++ro, co).Value = "Number as Text:";
            ws.Cell(ro, co + 1).Value = "'123.45";

            ws.Cell(++ro, co).Value = "TimeSpan as Text:";
            ws.Cell(ro, co + 1).Value = "'" + new TimeSpan(33, 45, 22).ToString();

            ro++;

            ws.Cell(++ro, co).Value = "Changing Data Types:";

            ro++;

            ws.Cell(++ro, co).Value = "Date to Text:";
            ws.Cell(ro, co + 1).Value = new DateTime(2010, 9, 2);
            ws.Cell(ro, co + 1).DataType = XLCellValues.Text;

            ws.Cell(++ro, co).Value = "DateTime to Text:";
            ws.Cell(ro, co + 1).Value = new DateTime(2010, 9, 2, 13, 45, 22);
            ws.Cell(ro, co + 1).DataType = XLCellValues.Text;

            ws.Cell(++ro, co).Value = "Boolean to Text:";
            ws.Cell(ro, co + 1).Value = true;
            ws.Cell(ro, co + 1).DataType = XLCellValues.Text;

            ws.Cell(++ro, co).Value = "Number to Text:";
            ws.Cell(ro, co + 1).Value = 123.45;
            ws.Cell(ro, co + 1).DataType = XLCellValues.Text;

            ws.Cell(++ro, co).Value = "TimeSpan to Text:";
            ws.Cell(ro, co + 1).Value = new TimeSpan(33, 45, 22);
            ws.Cell(ro, co + 1).DataType = XLCellValues.Text;

            ws.Cell(++ro, co).Value = "Text to Date:";
            ws.Cell(ro, co + 1).Value = "'" + new DateTime(2010, 9, 2).ToString();
            ws.Cell(ro, co + 1).DataType = XLCellValues.DateTime;

            ws.Cell(++ro, co).Value = "Text to DateTime:";
            ws.Cell(ro, co + 1).Value = "'" + new DateTime(2010, 9, 2, 13, 45, 22).ToString();
            ws.Cell(ro, co + 1).DataType = XLCellValues.DateTime;

            ws.Cell(++ro, co).Value = "Text to Boolean:";
            ws.Cell(ro, co + 1).Value = "'" + true.ToString();
            ws.Cell(ro, co + 1).DataType = XLCellValues.Boolean;

            ws.Cell(++ro, co).Value = "Text to Number:";
            ws.Cell(ro, co + 1).Value = "'123.45";
            ws.Cell(ro, co + 1).DataType = XLCellValues.Number;

            ws.Cell(++ro, co).Value = "Text to TimeSpan:";
            ws.Cell(ro, co + 1).Value = "'" + new TimeSpan(33, 45, 22).ToString();
            ws.Cell(ro, co + 1).DataType = XLCellValues.TimeSpan;

            ro++;

            ws.Cell(++ro, co).Value = "Formatted Date to Text:";
            ws.Cell(ro, co + 1).Value = new DateTime(2010, 9, 2);
            ws.Cell(ro, co + 1).Style.DateFormat.Format = "yyyy-MM-dd";
            ws.Cell(ro, co + 1).DataType = XLCellValues.Text;

            ws.Cell(++ro, co).Value = "Formatted Number to Text:";
            ws.Cell(ro, co + 1).Value = 12345.6789;
            ws.Cell(ro, co + 1).Style.NumberFormat.Format = "#,##0.00";
            ws.Cell(ro, co + 1).DataType = XLCellValues.Text;

            ro++;

            ws.Cell(++ro, co).Value = "Blank Text:";
            ws.Cell(ro, co + 1).Value = 12345.6789;
            ws.Cell(ro, co + 1).Style.NumberFormat.Format = "#,##0.00";
            ws.Cell(ro, co + 1).DataType = XLCellValues.Text;
            ws.Cell(ro, co + 1).Value = "";

            ro++;

            // Using inline strings (few users will ever need to use this feature)
            //
            // By default all strings are stored as shared so one block of text
            // can be reference by multiple cells.
            // You can override this by setting the .ShareString property to false
            ws.Cell(++ro, co).Value = "Inline String:";
            var cell = ws.Cell(ro, co + 1);
            cell.Value = "Not Shared";
            cell.ShareString = false;

            // To view all shared strings (all texts in the workbook actually), use the following:
            // workbook.GetSharedStrings()

            ws.Columns(2, 3).AdjustToContents();

            //wb.SaveAs("DataTypes.xlsx");

            return ExportExcel(wb, "DataTypes");
        }

        /// <summary>
        /// https://github.com/ClosedXML/ClosedXML/wiki/Creating-Multiple-Worksheets
        /// https://github.com/ClosedXML/ClosedXML/wiki/Organizing-Sheets
        /// </summary>
        /// <returns></returns>
        public ActionResult MultipleWs()
        {
            var wb = new XLWorkbook();
            foreach (var wsNum in Enumerable.Range(1, 5))
            {
                wb.Worksheets.Add("Original Pos. is " + wsNum.ToString());
            }

            // 移動至最後一個
            wb.Worksheet(1).Position = wb.Worksheets.Count() + 1;
            var ws1 = wb.Worksheet(1);
            ws1.Cell(1, 1).Value = "Hello";

            wb.Worksheet(4).Delete();

            wb.Worksheet(2).Position = 1;
            // 下面的 1 是上面換位置的 2
            var ws2 = wb.Worksheet(1);
            ws2.Cell(2, 2).Value = "World!";

            return ExportExcel(wb, "MultipleProcess");
        }

        /// <summary>
        /// https://github.com/ClosedXML/ClosedXML/wiki/Loading-and-Modifying-Files
        /// </summary>
        /// <returns></returns>
        public ActionResult ModifyingFiles()
        {
            // TODO: Path
            string fileName = @"C:\Users\BruceChen\Downloads\BasicTable.xlsx";
            var wb = new XLWorkbook(fileName);
            var ws = wb.Worksheet(1);

            // Change the background color of the headers
            var rngHeaders = ws.Range("B3:F3");
            rngHeaders.Style.Fill.BackgroundColor = XLColor.LightSalmon;

            // Change the date formats
            var rngDates = ws.Range("E4:E6");
            rngDates.Style.DateFormat.Format = "MM/dd/yyyy";

            // Change the income values to text
            var rngNumbers = ws.Range("F4:F6");
            foreach (var cell in rngNumbers.Cells())
            {
                cell.DataType = XLCellValues.Text;
                cell.Value += " Dollars";
            }

            return ExportExcel(wb, "BasicTableModified");
        }

        /// <summary>
        /// https://github.com/ClosedXML/ClosedXML/wiki/Using-Lambda-Expressions
        /// </summary>
        /// <returns></returns>
        public ActionResult Lambda()
        {
            string fileName = @"C:\Users\BruceChen\Downloads\BasicTable.xlsx";
            var wb = new XLWorkbook(fileName);
            var ws = wb.Worksheet(1);

            // Define a range with the data
            var firstDataCell = ws.Cell("B4");
            var lastDataCell = ws.LastCellUsed();
            // 定義資料範圍
            var rngData = ws.Range(firstDataCell.Address, lastDataCell.Address);

            // Delete all rows where Outcast = false (the 3rd column)
            // 注意 !，所以找 false
            rngData.Rows(r => !r.Cell(3).GetBoolean()) // where the 3rd cell of each row is false
                .ForEach(r => r.Delete()); // delete the row and shift the cells up (the default for rows in a range)

            // Put a light gray background to all text cells
            rngData.Cells(c => c.DataType == XLCellValues.Text) // where the data type is Text
                .ForEach(c => c.Style.Fill.BackgroundColor = XLColor.LightGray); // Fill with a light gray

            // Put a thick border to the bottom of the table (we may have deleted the bottom cells with the border)
            rngData.LastRow().Style.Border.BottomBorder = XLBorderStyleValues.Thick;

            return ExportExcel(wb, "LambdaExpressions");
        }

        /// <summary>
        /// https://github.com/ClosedXML/ClosedXML/wiki/Cell-Values
        /// </summary>
        /// <returns></returns>
        public ActionResult CellValues()
        {
            GetInstance("Cell Values", out XLWorkbook wb, out IXLWorksheet ws);
            ws.Cell("A1").SetValue(1).CellBelow().SetValue(1);

            // Set the titles
            ws.Cell(2, 2).Value = "Initial Value";
            ws.Cell(2, 3).Value = "Casting";
            ws.Cell(2, 4).Value = "Using Get...()";
            ws.Cell(2, 5).Value = "Using GetValue<T>()";
            ws.Cell(2, 6).Value = "GetString()";
            ws.Cell(2, 7).Value = "GetFormattedString()";

            //------------------------------
            // DateTime

            // Fill a cell with a date
            var cellDateTime = ws.Cell(3, 2);
            cellDateTime.Value = new DateTime(2010, 9, 2);
            cellDateTime.Style.DateFormat.Format = "yyyy-MMM-dd";

            // Extract the date in different ways
            DateTime dateTime1 = (DateTime)cellDateTime.Value;
            DateTime dateTime2 = cellDateTime.GetDateTime();
            DateTime dateTime3 = cellDateTime.GetValue<DateTime>();
            String dateTimeString = cellDateTime.GetString();
            String dateTimeFormattedString = cellDateTime.GetFormattedString();

            // Set the values back to cells
            // The apostrophe is to force ClosedXML to treat the date as a string
            ws.Cell(3, 3).Value = dateTime1;
            ws.Cell(3, 4).Value = dateTime2;
            ws.Cell(3, 5).Value = dateTime3;
            ws.Cell(3, 6).Value = "'" + dateTimeString;
            ws.Cell(3, 7).Value = "'" + dateTimeFormattedString;

            //------------------------------
            // Boolean

            // Fill a cell with a boolean
            var cellBoolean = ws.Cell(4, 2);
            cellBoolean.Value = true;

            // Extract the boolean in different ways
            Boolean boolean1 = (Boolean)cellBoolean.Value;
            Boolean boolean2 = cellBoolean.GetBoolean();
            Boolean boolean3 = cellBoolean.GetValue<Boolean>();
            String booleanString = cellBoolean.GetString();
            String booleanFormattedString = cellBoolean.GetFormattedString();

            // Set the values back to cells
            // The apostrophe is to force ClosedXML to treat the boolean as a string
            ws.Cell(4, 3).Value = boolean1;
            ws.Cell(4, 4).Value = boolean2;
            ws.Cell(4, 5).Value = boolean3;
            ws.Cell(4, 6).Value = "'" + booleanString;
            ws.Cell(4, 7).Value = "'" + booleanFormattedString;

            //------------------------------
            // Double

            // Fill a cell with a double
            var cellDouble = ws.Cell(5, 2);
            cellDouble.Value = 1234.567;
            cellDouble.Style.NumberFormat.Format = "#,##0.00";

            // Extract the double in different ways
            Double double1 = (Double)cellDouble.Value;
            Double double2 = cellDouble.GetDouble();
            Double double3 = cellDouble.GetValue<Double>();
            String doubleString = cellDouble.GetString();
            String doubleFormattedString = cellDouble.GetFormattedString();

            // Set the values back to cells
            // The apostrophe is to force ClosedXML to treat the double as a string
            ws.Cell(5, 3).Value = double1;
            ws.Cell(5, 4).Value = double2;
            ws.Cell(5, 5).Value = double3;
            ws.Cell(5, 6).Value = "'" + doubleString;
            ws.Cell(5, 7).Value = "'" + doubleFormattedString;

            //------------------------------
            // String

            // Fill a cell with a string
            var cellString = ws.Cell(6, 2);
            cellString.Value = "Test Case";

            // Extract the string in different ways
            String string1 = (String)cellString.Value;
            String string2 = cellString.GetString();
            String string3 = cellString.GetValue<String>();
            String stringString = cellString.GetString();
            String stringFormattedString = cellString.GetFormattedString();

            // Set the values back to cells
            ws.Cell(6, 3).Value = string1;
            ws.Cell(6, 4).Value = string2;
            ws.Cell(6, 5).Value = string3;
            ws.Cell(6, 6).Value = stringString;
            ws.Cell(6, 7).Value = stringFormattedString;

            //------------------------------
            // TimeSpan

            // Fill a cell with a timeSpan
            var cellTimeSpan = ws.Cell(7, 2);
            cellTimeSpan.Value = new TimeSpan(1, 2, 31, 45);

            // Extract the timeSpan in different ways
            TimeSpan timeSpan1 = (TimeSpan)cellTimeSpan.Value;
            TimeSpan timeSpan2 = cellTimeSpan.GetTimeSpan();
            TimeSpan timeSpan3 = cellTimeSpan.GetValue<TimeSpan>();
            String timeSpanString = cellTimeSpan.GetString();
            String timeSpanFormattedString = cellTimeSpan.GetFormattedString();

            // Set the values back to cells
            ws.Cell(7, 3).Value = timeSpan1;
            ws.Cell(7, 4).Value = timeSpan2;
            ws.Cell(7, 5).Value = timeSpan3;
            ws.Cell(7, 6).Value = "'" + timeSpanString;
            ws.Cell(7, 7).Value = "'" + timeSpanFormattedString;

            //------------------------------
            // Do some formatting
            ws.Columns("B:G").Width = 20;
            var rngTitle = ws.Range("B2:G2");
            rngTitle.Style.Font.Bold = true;
            rngTitle.Style.Fill.BackgroundColor = XLColor.Cyan;

            ws.Columns().AdjustToContents();

            //workbook.SaveAs("CellValues.xlsx");

            return ExportExcel(wb, "CellValues");
        }

        /// <summary>
        /// https://github.com/ClosedXML/ClosedXML/wiki/Workbook-Properties
        /// </summary>
        /// <returns></returns>
        public ActionResult WorkbookProperties()
        {
            GetInstance("Workbook Properties", out XLWorkbook wb, out IXLWorksheet ws);
            ws.Cell("A1").SetValue(1).CellBelow().SetValue(1);

            wb.Properties.Author = "theAuthor";
            wb.Properties.Title = "theTitle";
            wb.Properties.Subject = "theSubject";
            wb.Properties.Category = "theCategory";
            wb.Properties.Keywords = "theKeywords";
            wb.Properties.Comments = "theComments";
            wb.Properties.Status = "theStatus";
            wb.Properties.LastModifiedBy = "theLastModifiedBy";
            wb.Properties.Company = "theCompany";
            wb.Properties.Manager = "theManager";

            wb.CustomProperties.Add("theText", "XXX");
            wb.CustomProperties.Add("theDate", new DateTime(2011, 1, 1));
            wb.CustomProperties.Add("theNumber", 123.456);
            wb.CustomProperties.Add("theBoolean", true);

            //wb.SaveAs("WorkbookProperties.xlsx");

            return ExportExcel(wb, "WorkbookProperties");
        }

        /// <summary>
        /// https://github.com/ClosedXML/ClosedXML/wiki/Using-Formulas
        /// </summary>
        /// <returns></returns>
        public ActionResult Formulas()
        {
            GetInstance("Formulas", out XLWorkbook wb, out IXLWorksheet ws);
            ws.Cell("A1").SetValue(1).CellBelow().SetValue(1);

            ws.Cell(1, 1).Value = "Num1";
            ws.Cell(1, 2).Value = "Num2";
            ws.Cell(1, 3).Value = "Total";
            ws.Cell(1, 4).Value = "cell.FormulaA1";
            ws.Cell(1, 5).Value = "cell.FormulaR1C1";
            ws.Cell(1, 6).Value = "cell.Value";
            ws.Cell(1, 7).Value = "Are Equal?";

            ws.Cell(2, 1).Value = 1;
            ws.Cell(2, 2).Value = 2;
            var cellWithFormulaA1 = ws.Cell(2, 3);

            // ---------------------------------------------
            // cell.Value 測試結果與 wiki 上有差異
            // cell.Value 會取得實際值，而且 Formulas 算式
            // 原因可以看下一篇 wiki
            // ---------------------------------------------

            // Use A1 notation
            cellWithFormulaA1.FormulaA1 = "=A2+$B$2"; // The equal sign (=) in a formula is optional
            ws.Cell(2, 4).Value = cellWithFormulaA1.FormulaA1;
            ws.Cell(2, 5).Value = cellWithFormulaA1.FormulaR1C1;
            ws.Cell(2, 6).Value = cellWithFormulaA1.Value;

            ws.Cell(3, 1).Value = 1;
            ws.Cell(3, 2).Value = 2;
            var cellWithFormulaR1C1 = ws.Cell(3, 3);
            // Use R1C1 notation
            cellWithFormulaR1C1.FormulaR1C1 = "RC[-2]+R3C2"; // The equal sign (=) in a formula is optional
            ws.Cell(3, 4).Value = cellWithFormulaR1C1.FormulaA1;
            ws.Cell(3, 5).Value = cellWithFormulaR1C1.FormulaR1C1;
            ws.Cell(3, 6).Value = cellWithFormulaR1C1.Value;

            ws.Cell(4, 1).Value = "A";
            ws.Cell(4, 2).Value = "B";
            var cellWithStringFormula = ws.Cell(4, 3);

            // Use R1C1 notation
            cellWithStringFormula.FormulaR1C1 = "=\"Test\" & RC[-2] & \"R3C2\"";
            ws.Cell(4, 4).Value = cellWithStringFormula.FormulaA1;
            ws.Cell(4, 5).Value = cellWithStringFormula.FormulaR1C1;
            ws.Cell(4, 6).Value = cellWithStringFormula.Value;

            // Setting the formula of a range
            var rngData = ws.Range(2, 1, 4, 7);
            rngData.LastColumn().FormulaR1C1 = "=IF(RC[-3]=RC[-1],\"Yes\", \"No\")";

            // Using an array formula:
            // Just put the formula between curly braces
            ws.Cell("A6").Value = "Array Formula: ";
            ws.Cell("B6").FormulaA1 = "{A2+A3}";

            ws.Range(1, 1, 1, 7).Style.Fill.BackgroundColor = XLColor.Cyan;
            ws.Range(1, 1, 1, 7).Style.Font.Bold = true;
            ws.Columns().AdjustToContents();

            // You can also change the reference notation:
            wb.ReferenceStyle = XLReferenceStyle.R1C1;

            // And the workbook calculation mode:
            wb.CalculateMode = XLCalculateMode.Auto;

            //wb.SaveAs("Formulas.xlsx");

            return ExportExcel(wb, "Formulas");
        }

        /// <summary>
        /// https://github.com/ClosedXML/ClosedXML/wiki/Evaluating-Formulas
        /// </summary>
        /// <remarks>
        /// 注意：並不是所有公式都支援或有錯誤時，可能會產生重大錯誤。支援公式請查詢 wiki 文件。
        /// </remarks>
        /// <returns></returns>
        public ActionResult EvaluatingFormulas()
        {
            GetInstance("Sheet1", out XLWorkbook wb, out IXLWorksheet ws);
            // CellBelow 是下面一個 cell
            ws.Cell("A1").SetValue(1).CellBelow().SetValue(1);
            ws.Cell("B1").SetValue(1).CellBelow().SetValue(1);
            //ws.Cell("A1").SetValue(1);
            //ws.Cell("B1").SetValue(1);

            ws.Cell("C1").FormulaA1 = "\"The total value is: \" & SUM(A1:B2)";

            // 這裡取得實際值
            var r = ws.Cell("C1").Value;

            // ----------------------------------------------
            // sum = 6, 使用 XLWorkbook.EvaluateExpr 指定運算式
            // 注意，XLWorkbook 是直接運算，它不知道 worksheet
            var sum = XLWorkbook.EvaluateExpr("SUM(1,2,3)");
            ws.Cell(1, 4).Value = "XLWorkbook - Sum:" + sum;

            // 知道 wb, 但不使用原 Sheet 的 Range 計算
            var sum2 = wb.Evaluate("SUM(Sheet1!A1:B1)");
            ws.Cell(1, 5).Value = "wb - Sum:" + sum2;

            // 知道 ws 來計算
            var sum3 = ws.Evaluate("SUM(A1:B2)");
            ws.Cell(1, 6).Value = "ws - Sum:" + sum3;



            return ExportExcel(wb, "EvaluatingFormulas");
        }

        /// <summary>
        /// https://github.com/ClosedXML/ClosedXML/wiki/Creating-Rows-And-Columns-Outlines
        /// </summary>
        /// <returns></returns>
        public ActionResult RowsColumnsOutlines()
        {
            GetInstance("Outline", out XLWorkbook wb, out IXLWorksheet ws);

            ws.Outline.SummaryHLocation = XLOutlineSummaryHLocation.Right;
            ws.Columns(2, 6).Group(); // Create an outline (level 1) for columns 2-6
            ws.Columns(2, 4).Group(); // Create an outline (level 2) for columns 2-4
            ws.Column(2).Ungroup(true); // Remove column 2 from all outlines

            ws.Outline.SummaryVLocation = XLOutlineSummaryVLocation.Bottom;
            ws.Rows(1, 5).Group(); // Create an outline (level 1) for rows 1-5
            ws.Rows(1, 4).Group(); // Create an outline (level 2) for rows 1-4
            ws.Rows(1, 4).Collapse(); // Collapse rows 1-4
            ws.Rows(1, 2).Group(); // Create an outline (level 3) for rows 1-2
            ws.Rows(1, 2).Ungroup(); // Ungroup rows 1-2 from their last outline

            return ExportExcel(wb, "RowsColumnsOutlines");
        }

        /// <summary>
        /// https://github.com/ClosedXML/ClosedXML/wiki/Hide-Unhide-Row(s)-And-Column(s)
        /// </summary>
        /// <returns></returns>
        public ActionResult HideUnhideRowsColumns()
        {
            GetInstance("HideUnhide", out XLWorkbook wb, out IXLWorksheet ws);

            ws.Columns(1, 3).Hide();
            ws.Rows(1, 3).Hide();

            ws.Column(2).Unhide();
            ws.Row(2).Unhide();

            return ExportExcel(wb, "HideUnhide");
        }

        /// <summary>
        /// https://github.com/ClosedXML/ClosedXML/wiki/Freeze-Panes
        /// </summary>
        /// <returns></returns>
        public ActionResult FreezePanes()
        {
            GetInstance("Freeze View", out XLWorkbook wb, out IXLWorksheet ws);

            ws.SheetView.Freeze(3, 3);

            // 一一指定
            // ws.SheetView.FreezeRows(3);
            // ws.SheetView.FreezeColumns(3);

            // 調整分割, 指定為 0 是移除
            ws.SheetView.SplitRow = 2;
            ws.SheetView.SplitColumn = 0;

            return ExportExcel(wb, "FreezePanes");
        }

        /// <summary>
        /// https://github.com/ClosedXML/ClosedXML/wiki/Copying-Worksheets
        /// </summary>
        /// <returns></returns>
        public ActionResult CopyingWs()
        {
            string fileName = @"C:\Users\BruceChen\Downloads\BasicTable.xlsx";
            var wb = new XLWorkbook(fileName);
            var ws = wb.Worksheet(1);
            // copy 至另一 ws
            ws.CopyTo("Copy 1");

            // 新 instance 再做一次 copy 至另一 ws
            var wbNew = new XLWorkbook(fileName);
            wbNew.Worksheet(1).CopyTo(wb, "Copy 2");

            // 會含另 2 個 ws
            return ExportExcel(wb, "CopyingWorksheets");

        }

        /// <summary>
        /// Exports the excel.
        /// </summary>
        /// <param name="wb">The XLWorkbook object.</param>
        /// <param name="fileName">Name of the file.</param>
        /// <returns></returns>
        [NonAction]
        private ActionResult ExportExcel(XLWorkbook wb, string fileName)
        {
            using (var ms = new MemoryStream())
            {
                wb.SaveAs(ms);
                ms.Seek(0, SeekOrigin.Begin);
                return File(ms.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", $"{fileName}.xlsx");
                //return File(ms.ToArray(), "application/vnd.ms-excel", $"{fileName}.xls");
            }
        }


        /// <summary>
        /// XLWorkbook, IXLWorksheet Factory
        /// </summary>
        /// <param name="sheetName">The Worksheet Name</param>
        /// <param name="wb">The XLWorkbook object.</param>
        /// <param name="ws">The IXLWorksheet interface object</param>
        private static void GetInstance(string sheetName, out XLWorkbook wb, out IXLWorksheet ws)
        {
            if (string.IsNullOrWhiteSpace(sheetName))
            {
                throw new ArgumentNullException(nameof(sheetName));
            }

            wb = new XLWorkbook();
            ws = wb.AddWorksheet(sheetName);
        }
    }
}