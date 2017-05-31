using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web.Mvc;
using ClosedXML.Excel;
using ClosedXMLSample.Models;

namespace ClosedXMLSample.Controllers
{
    public class InsertingController : BaseController
    {
        /// <summary>
        /// https://github.com/closedxml/closedxml/wiki/Copying-IEnumerable-Collections
        /// </summary>
        /// <returns></returns>
        public ActionResult Collections()
        {
            GetInstance("Collections", out XLWorkbook wb, out IXLWorksheet ws);

            // From a list of strings
            var listOfStrings = new List<string> {"House", "Car"};
            ws.Cell(1, 1).Value = "Strings";
            ws.Cell(1, 1).AsRange().AddToNamed("Titles");
            ws.Cell(2, 1).Value = listOfStrings;

            // From a list of arrays
            var listOfArr = new List<int[]>
            {
                new[] {1, 2, 3},
                new[] {1},
                new[] {1, 2, 3, 4, 5, 6}
            };
            ws.Cell(1, 3).Value = "Arrays";
            ws.Range(1, 3, 1, 8).Merge().AddToNamed("Titles");
            ws.Cell(2, 3).Value = listOfArr;

            // From a DataTable
            var dataTable = GetTable();
            ws.Cell(6, 1).Value = "DataTable";
            ws.Range(6, 1, 6, 4).Merge().AddToNamed("Titles");
            ws.Cell(7, 1).Value = dataTable.AsEnumerable();     // 記得轉型為 AsEnumerable

            // From a query
            var list = new List<Person>
            {
                new Person() {Name = "John", Age = 30, House = "On Elm St."},
                new Person() {Name = "Mary", Age = 15, House = "On Main St."},
                new Person() {Name = "Luis", Age = 21, House = "On 23rd St."},
                new Person() {Name = "Henry", Age = 45, House = "On 5th Ave."}
            };

            var people = from p in list
                where p.Age >= 21
                select new { p.Name, p.House, p.Age };

            ws.Cell(6, 6).Value = "Query";
            ws.Range(6, 6, 6, 8).Merge().AddToNamed("Titles");
            ws.Cell(7, 6).Value = people.AsEnumerable();    // 記得轉型為 AsEnumerable

            // Prepare the style for the titles
            var titlesStyle = wb.Style;
            titlesStyle.Font.Bold = true;
            titlesStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            titlesStyle.Fill.BackgroundColor = XLColor.Cyan;

            // Format all titles in one shot
            wb.NamedRanges.NamedRange("Titles").Ranges.Style = titlesStyle;

            ws.Columns().AdjustToContents();

            //wb.SaveAs("Collections.xlsx");

            return ExportExcel(wb, "Collections");
        }

        private DataTable GetTable()
        {
            DataTable table = new DataTable();
            table.Columns.Add("Dosage", typeof(int));
            table.Columns.Add("Drug", typeof(string));
            table.Columns.Add("Patient", typeof(string));
            table.Columns.Add("Date", typeof(DateTime));

            table.Rows.Add(25, "Indocin", "David", DateTime.Now);
            table.Rows.Add(50, "Enebrel", "Sam", DateTime.Now);
            table.Rows.Add(10, "Hydralazine", "Christoff", DateTime.Now);
            table.Rows.Add(21, "Combivent", "Janet", DateTime.Now);
            table.Rows.Add(100, "Dilantin", "Melanie", DateTime.Now);
            return table;
        }
        [NonAction]
        public ActionResult CodeBase()
        {
            GetInstance("Collections", out XLWorkbook wb, out IXLWorksheet ws);
            return ExportExcel(wb, "Collections");
        }
    }
}