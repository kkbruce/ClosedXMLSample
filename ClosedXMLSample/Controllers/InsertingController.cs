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
        /// Using the cell.Value = collection method.
        /// </summary>
        /// <returns></returns>
        public ActionResult Collections()
        {
            GetInstance("Collections", out XLWorkbook wb, out IXLWorksheet ws);

            // From a list of strings
            var listOfStrings = new List<string> { "House", "Car" };
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

        /// <summary>
        /// https://github.com/ClosedXML/ClosedXML/wiki/Inserting-Data
        ///  Using the cell.InsertData(collection) method.
        /// </summary>
        /// <returns></returns>
        public ActionResult InsertData()
        {
            // Note:
            // The difference between InsertData and InsertTable is that InsertData
            // doesn't insert column names and returns a range.
            // InsertTable will insert the column names and returns a table.

            GetInstance("Inserting Data", out XLWorkbook wb, out IXLWorksheet ws);

            // From a list of strings
            var listOfStrings = new List<string>
            {
                "House",
                "Car"
            };
            ws.Cell(1, 1).Value = "From Strings";
            ws.Cell(1, 1).AsRange().AddToNamed("Titles");
            // DataRange: A2:A3
            var rangeWithStrings = ws.Cell(2, 1).InsertData(listOfStrings);

            // From a list of arrays
            var listOfArr = new List<int[]>
            {
                new[] { 1, 2, 3 },
                new[] { 1 },
                new[] { 1, 2, 3, 4, 5, 6 }
            };
            ws.Cell(1, 3).Value = "From Arrays";
            ws.Range(1, 3, 1, 8).Merge().AddToNamed("Titles");
            // DataRange: C2:H4
            var rangeWithArrays = ws.Cell(2, 3).InsertData(listOfArr);

            // From a DataTable
            var dataTable = GetTable();
            ws.Cell(6, 1).Value = "From DataTable";
            ws.Range(6, 1, 6, 4).Merge().AddToNamed("Titles");
            // DataRange: A7:D11
            var rangeWithData = ws.Cell(7, 1).InsertData(dataTable.AsEnumerable());

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

            ws.Cell(6, 6).Value = "From Query";
            ws.Range(6, 6, 6, 8).Merge().AddToNamed("Titles");
            // DataRange: F7:H9
            var rangeWithPeople = ws.Cell(7, 6).InsertData(people.AsEnumerable());

            // Prepare the style for the titles
            var titlesStyle = wb.Style;
            titlesStyle.Font.Bold = true;
            titlesStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            titlesStyle.Fill.BackgroundColor = XLColor.Cyan;

            // Format all titles in one shot
            wb.NamedRanges.NamedRange("Titles").Ranges.Style = titlesStyle;

            ws.Columns().AdjustToContents();

            //wb.SaveAs("InsertingData.xlsx");

            return ExportExcel(wb, "InsertingData");
        }

        /// <summary>
        /// https://github.com/ClosedXML/ClosedXML/wiki/Inserting-Tables
        /// Using the cell.InsertTable(collection) method.
        /// </summary>
        /// <returns></returns>
        public ActionResult InsertTable()
        {
            //  InsertTable will insert the column names and returns a table.

            GetInstance("Inserting Tables", out XLWorkbook wb, out IXLWorksheet ws);

            // From a list of strings
            var listOfStrings = new List<string>
            {
                "House",
                "Car"
            };
            ws.Cell(1, 1).Value = "From Strings";
            ws.Cell(1, 1).AsRange().AddToNamed("Titles");
            // Range: A2:A4
            var tableWithStrings = ws.Cell(2, 1).InsertTable(listOfStrings);

            // From a list of arrays
            var listOfArr = new List<int[]>
            {
                new[] { 1, 2, 3 },
                new[] { 1 },
                new[] { 1, 2, 3, 4, 5, 6 }
            };
            ws.Cell(1, 3).Value = "From Arrays";
            ws.Range(1, 3, 1, 8).Merge().AddToNamed("Titles");
            // Range: C2:H5
            var tableWithArrays = ws.Cell(2, 3).InsertTable(listOfArr);

            // From a DataTable
            var dataTable = GetTable();
            ws.Cell(7, 1).Value = "From DataTable";
            ws.Range(7, 1, 7, 4).Merge().AddToNamed("Titles");
            // Range: A8:D13
            var tableWithData = ws.Cell(8, 1).InsertTable(dataTable.AsEnumerable());

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

            ws.Cell(7, 6).Value = "From Query";
            ws.Range(7, 6, 7, 8).Merge().AddToNamed("Titles");
            // Range: F8:H11
            var tableWithPeople = ws.Cell(8, 6).InsertTable(people.AsEnumerable());

            // Prepare the style for the titles
            var titlesStyle = wb.Style;
            titlesStyle.Font.Bold = true;
            titlesStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            titlesStyle.Fill.BackgroundColor = XLColor.Cyan;

            // Format all titles in one shot
            wb.NamedRanges.NamedRange("Titles").Ranges.Style = titlesStyle;

            ws.Columns().AdjustToContents();

            //wb.SaveAs("InsertingTables.xlsx");

            return ExportExcel(wb, "InsertingTables");
        }

        /// <summary>
        /// https://github.com/ClosedXML/ClosedXML/wiki/Adding-DataTable-as-Worksheet
        /// </summary>
        /// <returns></returns>
        public ActionResult DataTable()
        {
            var wb = new XLWorkbook();
            var dataTable = GetTable("Information");

            // Add a DataTable as a worksheet
            wb.Worksheets.Add(dataTable);

            //wb.SaveAs("AddingDataTableAsWorksheet.xlsx");

            return ExportExcel(wb, "AddingDataTableAsWorksheet");
        }

        /// <summary>
        /// https://github.com/ClosedXML/ClosedXML/wiki/Adding-DataSet
        /// </summary>
        /// <returns></returns>
        public ActionResult DataSet()
        {
            var wb = new XLWorkbook();
            var dataSet = GetDataSet();

            // Add all DataTables in the DataSet as a worksheets
            wb.Worksheets.Add(dataSet);

            //wb.SaveAs("AddingDataSet.xlsx");

            return ExportExcel(wb, "AddingDataSet");
        }

        private DataSet GetDataSet()
        {
            var ds = new DataSet();
            ds.Tables.Add(GetTable("Patients"));
            ds.Tables.Add(GetTable("Employees"));
            ds.Tables.Add(GetTable("Information"));
            return ds;
        }

        private DataTable GetTable()
        {
            DataTable table = new DataTable();
            TableData(ref table);
            return table;
        }

        private DataTable GetTable(string tableName)
        {
            DataTable table = new DataTable
            {
                TableName = tableName
            };
            TableData(ref table);
            return table;
        }

        private void TableData(ref DataTable table)
        {
            table.Columns.Add("Dosage", typeof(int));
            table.Columns.Add("Drug", typeof(string));
            table.Columns.Add("Patient", typeof(string));
            table.Columns.Add("Date", typeof(DateTime));

            table.Rows.Add(25, "Indocin", "David", DateTime.Now);
            table.Rows.Add(50, "Enebrel", "Sam", DateTime.Now);
            table.Rows.Add(10, "Hydralazine", "Christoff", DateTime.Now);
            table.Rows.Add(21, "Combivent", "Janet", DateTime.Now);
            table.Rows.Add(100, "Dilantin", "Melanie", DateTime.Now);
        }
    }
}