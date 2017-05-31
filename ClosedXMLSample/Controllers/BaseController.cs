using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using ClosedXML.Excel;

namespace ClosedXMLSample.Controllers
{
    public class BaseController : Controller
    {
        /// <summary>
        /// Exports the excel.
        /// </summary>
        /// <param name="wb">The XLWorkbook object.</param>
        /// <param name="fileName">Name of the file.</param>
        /// <returns></returns>
        [NonAction]
        protected ActionResult ExportExcel(XLWorkbook wb, string fileName)
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
        protected static void GetInstance(string sheetName, out XLWorkbook wb, out IXLWorksheet ws)
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