using System;
using System.Linq;
using System.Web.Mvc;
using ClosedXML.Excel;

namespace ClosedXMLSample.Controllers
{
    public class MiscController : BaseController
    {
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
        /// https://github.com/ClosedXML/ClosedXML/wiki/Using-Hyperlinks
        /// </summary>
        /// <returns></returns>
        public ActionResult Hyperlinks()
        {
            GetInstance("Hyperlinks", out XLWorkbook wb, out IXLWorksheet ws);

            wb.Worksheets.Add("Second Sheet");

            Int32 ro = 0;

            // You can create a link with pretty much anything you can put on a
            // browser: http, ftp, mailto, gopher, news, nntp, etc.

            ws.Cell(++ro, 1).Value = "Link to a web page, no tooltip - Yahoo!";
            ws.Cell(ro, 1).Hyperlink = new XLHyperlink(@"http://www.yahoo.com");

            ws.Cell(++ro, 1).Value = "Link to a web page, with a tooltip - Yahoo!";
            ws.Cell(ro, 1).Hyperlink = new XLHyperlink(@"http://www.yahoo.com", "Click to go to Yahoo!");

            ws.Cell(++ro, 1).Value = "Link to a file - same folder";
            ws.Cell(ro, 1).Hyperlink = new XLHyperlink("Test.xlsx");

            ws.Cell(++ro, 1).Value = "Link to a file - relative address";
            ws.Cell(ro, 1).Hyperlink = new XLHyperlink(@"../Test.xlsx");

            ws.Cell(++ro, 1).Value = "Link to an address in this worksheet";
            ws.Cell(ro, 1).Hyperlink = new XLHyperlink("B1");

            ws.Cell(++ro, 1).Value = "Link to an address in another worksheet";
            ws.Cell(ro, 1).Hyperlink = new XLHyperlink("'Second Sheet'!A1");

            // You can also set the properties of a hyperlink directly:

            ws.Cell(++ro, 1).Value = "Link to a range in this worksheet";
            ws.Cell(ro, 1).Hyperlink.InternalAddress = "B1:C2";
            ws.Cell(ro, 1).Hyperlink.Tooltip = "SquareBox";

            ws.Cell(++ro, 1).Value = "Link to an email message";
            ws.Cell(ro, 1).Hyperlink.ExternalAddress = new Uri(@"mailto:SantaClaus@NorthPole.com?subject=Presents");

            // Deleting a hyperlink
            ws.Cell(++ro, 1).Value = "This is no longer a link";
            ws.Cell(ro, 1).Hyperlink.InternalAddress = "A1";
            ws.Cell(ro, 1).Hyperlink.Delete();

            // Setting a hyperlink preserves previous formatting:
            ws.Cell(++ro, 1).Value = "Odd looking link";
            ws.Cell(ro, 1).Style.Font.FontColor = XLColor.Red;
            ws.Cell(ro, 1).Style.Font.Underline = XLFontUnderlineValues.Double;
            ws.Cell(ro, 1).Hyperlink = new XLHyperlink(ws.Range("B1:C2"));

            // List all hyperlinks in a worksheet:
            var hyperlinksInWorksheet = ws.Hyperlinks;

            // List all hyperlinks in a range:
            var hyperlinksInRange = ws.Range("A1:A3").Hyperlinks;

            ws.Columns().AdjustToContents();

            //wb.SaveAs("Hyperlinks.xlsx");

            return ExportExcel(wb, "Hyperlinks");
        }

        /// <summary>
        /// https://github.com/ClosedXML/ClosedXML/wiki/Data-Validation
        /// </summary>
        /// <returns></returns>
        public ActionResult DataValidation()
        {
            // wiki 範例有問題，依 source code 修正
            // ----------------------------------------------------
            // 1. Change DataValidation to SetDataValidation()
            // 2. 後面的規則覆寫前面的
            // ----------------------------------------------------

            GetInstance("Data Validation", out XLWorkbook wb, out IXLWorksheet ws);

            // Decimal between 1 and 5
            ws.Cell(1, 1).SetDataValidation().Decimal.Between(1, 5);

            // Whole number equals 2
            var dv1 = ws.Range("A2:A3").SetDataValidation();
            dv1.WholeNumber.EqualTo(2);
            // Change the error message
            dv1.ErrorStyle = XLErrorStyle.Warning;
            dv1.ErrorTitle = "Number out of range";
            dv1.ErrorMessage = "This cell only allows the number 2.";

            // Date after the millenium
            var dv2 = ws.Cell("A4").SetDataValidation();
            dv2.Date.EqualOrGreaterThan(new DateTime(2000, 1, 1));
            // Change the input message
            dv2.InputTitle = "Can't party like it's 1999.";
            dv2.InputMessage = "Please enter a date in this century.";

            // From a list
            ws.Cell("C1").Value = "Yes";
            ws.Cell("C2").Value = "No";
            ws.Cell("A5").SetDataValidation().List(ws.Range("C1:C2"));

            ws.Range("C1:C2").AddToNamed("YesNo");
            ws.Cell("A6").SetDataValidation().List("=YesNo");

            // Intersecting dataValidations
            ws.Range("B1:B4").SetDataValidation().WholeNumber.EqualTo(1);
            ws.Range("B3:B4").SetDataValidation().WholeNumber.EqualTo(2);


            // Validate with multiple ranges
            var ws2 = wb.Worksheets.Add("Validate Ranges");
            var rng1 = ws2.Ranges("A1:B2,B4:D7,F4:G5");
            rng1.Style.Fill.SetBackgroundColor(XLColor.YellowGreen);
            var rng1Validation = rng1.SetDataValidation();
            rng1Validation.Decimal.EqualTo(1);
            rng1Validation.IgnoreBlanks = false;

            var rng2 = ws2.Range("A11:E14");
            rng2.Style.Fill.SetBackgroundColor(XLColor.YellowGreen);
            var rng2Validation = rng2.SetDataValidation();
            rng2Validation.Decimal.EqualTo(2);
            rng2Validation.IgnoreBlanks = false;

            var rng3 = ws2.Range("B2:B12");
            //rng3.Style.Fill.SetBackgroundColor(XLColor.YellowGreen);
            var rng3Validation = rng3.SetDataValidation();
            rng3Validation.Decimal.EqualTo(3);
            rng3Validation.IgnoreBlanks = true;

            var rng4 = ws2.Range("D5:D6");
            //rng4.Style.Fill.SetBackgroundColor(XLColor.YellowGreen);
            var rng4Validation = rng4.SetDataValidation();
            rng4Validation.Decimal.EqualTo(4);
            rng4Validation.IgnoreBlanks = true;

            var rng5 = ws2.Range("C13:C14");
            //rng5.Style.Fill.SetBackgroundColor(XLColor.YellowGreen);
            var rng5Validation = rng5.SetDataValidation();
            rng5Validation.Decimal.EqualTo(5);
            rng5Validation.IgnoreBlanks = true;

            var rng6 = ws2.Range("D11:D12");
            //rng6.Style.Fill.SetBackgroundColor(XLColor.YellowGreen);
            var rng6Validation = rng6.SetDataValidation();
            rng6Validation.Decimal.EqualTo(5);
            rng6Validation.IgnoreBlanks = true;

            var rng7 = ws2.Range("G4:G5");
            //rng7.Style.Fill.SetBackgroundColor(XLColor.YellowGreen);
            var rng7Validation = rng7.SetDataValidation();
            rng7Validation.Decimal.EqualTo(5);
            rng7Validation.IgnoreBlanks = true;

            ws.CopyTo(ws.Name + " - Copy");
            ws2.CopyTo(ws2.Name + " - Copy");

            wb.AddWorksheet("Copy From Range 1").FirstCell().Value = ws.RangeUsed(true);
            wb.AddWorksheet("Copy From Range 2").FirstCell().Value = ws2.RangeUsed(true);

            return ExportExcel(wb, "DataValidation");
        }

        /// <summary>
        /// https://github.com/ClosedXML/ClosedXML/wiki/Hide-Worksheets
        /// </summary>
        /// <returns></returns>
        public ActionResult HideWs()
        {
            GetInstance("Visible", out XLWorkbook wb, out IXLWorksheet ws);
            wb.Worksheets.Add("Hidden").Hide();
            wb.Worksheets.Add("Unhidden").Hide().Unhide();
            wb.Worksheets.Add("VeryHidden").Visibility = XLWorksheetVisibility.VeryHidden;

            return ExportExcel(wb, "HideWorksheets");
        }

        /// <summary>
        /// https://github.com/ClosedXML/ClosedXML/wiki/Sheet-Protection
        /// </summary>
        /// <returns></returns>
        public ActionResult SheetProtection()
        {
            GetInstance("Protected No-Password", out XLWorkbook wb, out IXLWorksheet ws);

            ws.Protect()            // On this sheet we will only allow:
              .SetFormatCells()   // Cell Formatting
              .SetInsertColumns() // Inserting Columns
              .SetDeleteColumns() // Deleting Columns
              .SetDeleteRows();   // Deleting Rows

            ws.Cell("A1").SetValue("Locked, No Hidden (Default):")
              .Style.Font.SetBold().Fill.SetBackgroundColor(XLColor.Cyan);
            ws.Cell("B1").Style.Border.SetOutsideBorder(XLBorderStyleValues.Medium);

            ws.Cell("A2").SetValue("Locked, Hidden:")
              .Style.Font.SetBold().Fill.SetBackgroundColor(XLColor.Cyan);
            ws.Cell("B2").Style
              .Protection.SetHidden()
              .Border.SetOutsideBorder(XLBorderStyleValues.Medium);

            ws.Cell("A3").SetValue("Not Locked, Hidden:")
              .Style.Font.SetBold().Fill.SetBackgroundColor(XLColor.Cyan);
            ws.Cell("B3").Style
              .Protection.SetLocked(false)
              .Protection.SetHidden()
              .Border.SetOutsideBorder(XLBorderStyleValues.Medium);

            ws.Cell("A4").SetValue("Not Locked, Not Hidden:")
              .Style.Font.SetBold().Fill.SetBackgroundColor(XLColor.Cyan);
            ws.Cell("B4").Style
              .Protection.SetLocked(false)
              .Border.SetOutsideBorder(XLBorderStyleValues.Medium);

            ws.Columns().AdjustToContents();

            // Protect a sheet with a password
            var protectedSheet = wb.Worksheets.Add("Protected Password = 123");
            var protection = protectedSheet.Protect("123");
            protection.InsertRows = true;
            protection.InsertColumns = true;

            return ExportExcel(wb, "SheetProtection");
        }

        /// <summary>
        /// https://github.com/ClosedXML/ClosedXML/wiki/Tab-Colors
        /// </summary>
        /// <returns></returns>
        public ActionResult TabColors()
        {
            GetInstance("Red", out XLWorkbook wb, out IXLWorksheet ws);

            ws.SetTabColor(XLColor.Red);

            var wsAccent3 = wb.Worksheets.Add("Accent3");
            wsAccent3.SetTabColor(XLColor.FromTheme(XLThemeColor.Accent3));

            var wsIndexed = wb.Worksheets.Add("Indexed");
            wsIndexed.TabColor = XLColor.FromIndex(24);

            var wsArgb = wb.Worksheets.Add("Argb");
            wsArgb.TabColor = XLColor.FromArgb(23, 23, 23);

            return ExportExcel(wb, "TabColors");
        }

    }
}