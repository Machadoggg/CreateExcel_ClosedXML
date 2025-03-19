using ClosedXML.Excel;
using CreateExcel_ClosedXML.Models;
using System.Collections.Generic;

namespace CreateExcel_ClosedXML
{
    public class Excel
    {
        public static void CreateExcel(string path)
        {
            var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("Sheet 1");
            worksheet.Cell(1, 1).Value = "Hello World";
            workbook.SaveAs(path);
        }

        public static void CreateExcel2(string path2)
        {
            var workbook2 = new XLWorkbook();
            var worksheet2 = workbook2.Worksheets.Add("Sheet 2");
            var customers = new List<Customer>
            {
                new Customer{Name="Juan",Age=18,Address="Calle33"},
                new Customer{Name="Juan",Age=18,Address="Calle33"}
            };

            //Headers
            worksheet2.Cell("A1").Value = "Name";
            worksheet2.Cell("B1").Value = "Age";
            worksheet2.Cell("C1").Value = "Address";

            //Insert data directly from the list
            worksheet2.Cell("A2").InsertData(customers);

            workbook2.SaveAs(path2);
        }
    }
}
