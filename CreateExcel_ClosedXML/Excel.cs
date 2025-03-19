using ClosedXML.Excel;
using CreateExcel_ClosedXML.Models;
using System.Collections.Generic;
using System.Data;

namespace CreateExcel_ClosedXML
{
    public class Excel
    {
        //Initial Format
        public static void CreateExcel(string path)
        {
            var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("Sheet 1");
            worksheet.Cell(1, 1).Value = "Hello World";
            workbook.SaveAs(path);
        }

        //Format with headers and initial cell data
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


        //With format
        public static byte[] CreateExcel3(string path) 
        {
            var _cusdata = GetCustomerData();
            using (XLWorkbook wb = new XLWorkbook())
            {
                wb.AddWorksheet(_cusdata, "Customer Records");
                using (MemoryStream ms = new MemoryStream())
                {
                    wb.SaveAs(path);
                    return ms.ToArray();
                }
                
            }
            
        }
        private static DataTable GetCustomerData() 
        {
            DataTable dt = new DataTable();
            dt.TableName = "CustomerData";
            dt.Columns.Add("Name", typeof(string));
            dt.Columns.Add("Age", typeof(int));
            dt.Columns.Add("Address", typeof(string));

            var customers = new List<Customer>()
            {
                new Customer{Name="Andres",Age=18,Address="Carrera 22"},
                new Customer{Name="Ana",Age=19,Address="Calle 45"},
                new Customer{Name="Camilo",Age=25,Address="Carrera 33"},
                new Customer{Name="Daniela",Age=34,Address="Avenida 22"},
                new Customer{Name="Juan",Age=65,Address="Carrera 57"},
                new Customer{Name="Jorge",Age=15,Address="Carrera 100"},
                new Customer{Name="Gabriela",Age=87,Address="Carrera 44"},
                new Customer{Name="Zaul",Age=56,Address="Avenida 93"},
                new Customer{Name="Xavier",Age=46,Address="Calle 21"}
            };

            if (customers.Count > 0) 
            {
                customers.ForEach(item => 
                { 
                    dt.Rows.Add(item.Name, item.Age, item.Address);
                });
            }

            return dt;
        }

    }
}
