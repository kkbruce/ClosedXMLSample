using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using ClosedXML.Excel;

namespace ClosedXMLSample.Controllers
{
    public class AJAXController : Controller
    {
        // GET: AJAX
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult ExportExcel()
        {
            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sample Sheet");
            ws.Cell("A1").Value = "Hello World!";
            wb.SaveAs("HelloWorld.xlsx");

            return ExportExcel(wb, "HelloWorld");
        }

        [HttpPost]
        public ActionResult ExportExcel(int numberId)
        {
            if (numberId == 9527)
            {
                #region SampleData

                var wb = new XLWorkbook();
                var ws = wb.Worksheets.Add("Contacts");

                ws.Cell("B2").Value = "Contacts";
                ws.Cell("B3").Value = "FName";
                ws.Cell("B4").Value = "John";
                ws.Cell("B5").Value = "Hank";
                ws.Cell("B6").SetValue("Dagny");
                ws.Cell("C3").Value = "LName";
                ws.Cell("C4").Value = "Galt";
                ws.Cell("C5").Value = "Rearden";
                ws.Cell("C6").SetValue("Taggart");
                ws.Cell("D3").Value = "Outcast";
                ws.Cell("D4").Value = true;
                ws.Cell("D5").Value = false;
                ws.Cell("D6").SetValue(false);
                ws.Cell("E3").Value = "DOB";
                ws.Cell("E4").Value = new DateTime(1919, 1, 21);
                ws.Cell("E5").Value = new DateTime(1907, 3, 4);
                ws.Cell("E6").SetValue(new DateTime(1921, 12, 15));
                ws.Cell("F3").Value = "Income";
                ws.Cell("F4").Value = 2000;
                ws.Cell("F5").Value = 40000;
                ws.Cell("F6").SetValue(10000);

                var rngTable = ws.Range("B2:F6");
                var rngDates = rngTable.Range("D3:D5");
                var rngNumbers = rngTable.Range("E3:E5");
                rngDates.Style.NumberFormat.NumberFormatId = 15;
                rngNumbers.Style.NumberFormat.Format = "$ #,##0";
                rngTable.FirstCell().Style
                    .Font.SetBold()
                    .Fill.SetBackgroundColor(XLColor.CornflowerBlue)
                    .Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                rngTable.FirstRow().Merge();

                var rngHeaders = rngTable.Range("A2:E2");
                rngHeaders.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                rngHeaders.Style.Font.Bold = true;
                rngHeaders.Style.Font.FontColor = XLColor.DarkBlue;
                rngHeaders.Style.Fill.BackgroundColor = XLColor.Aqua;

                var rngData = ws.Range("B3:F6");
                var excelTable = rngData.CreateTable();

                excelTable.ShowTotalsRow = true;
                excelTable.Field("Income").TotalsRowFunction = XLTotalsRowFunction.Sum;
                excelTable.Field("DOB").TotalsRowLabel = "Sum:";

                ws.RangeUsed().Style.Border.OutsideBorder = XLBorderStyleValues.Thick;

                ws.Columns().AdjustToContents();
                //wb.SaveAs("Showcase.xlsx");

                // use MVC output file
                //return ExportExcel(wb, "Showcase");

                // for AJAX
                string fileId = Guid.NewGuid().ToString();
                using (MemoryStream ms = new MemoryStream())
                {
                    wb.SaveAs(ms);
                    ms.Seek(0, SeekOrigin.Begin);
                    TempData[fileId] = ms.ToArray();
                }

                return new JsonResult() { Data = new { downloadId = fileId } };
                #endregion
            }

            return new EmptyResult();
        }

        public ActionResult ExportExcel(string downloadId)
        {
            if (TempData[downloadId] != null)
            {
                byte[] data = TempData[downloadId] as byte[];
                return File(data, "application/vnd.ms-excel", "Showcase.xlsx");
            }


            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("NoData");
            ws.Cell("A1").Value = "No Data!";
            using (var ms = new MemoryStream())
            {
                wb.SaveAs(ms);
                ms.Seek(0, SeekOrigin.Begin);
                return File(ms.ToArray(), "application/vnd.ms-excel", "NoData.xlsx");
            }
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
    }
}