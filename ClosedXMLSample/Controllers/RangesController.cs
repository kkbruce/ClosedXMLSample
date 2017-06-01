using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using ClosedXML.Excel;

namespace ClosedXMLSample.Controllers
{
    public class RangesController : BaseController
    {
        // GET: Ranges
        public ActionResult Index()
        {
            return View();
        }

        [NonAction]
        public ActionResult CodeBase()
        {
            GetInstance("Collections", out XLWorkbook wb, out IXLWorksheet ws);
            return ExportExcel(wb, "Collections");
        }
    }
}