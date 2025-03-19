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
            var _empdata = GetEmployeeData();
            using (XLWorkbook wb = new XLWorkbook())
            {
                //wb.AddWorksheet(_empdata, "Employee Records");
                var sheet1 = wb.AddWorksheet(_empdata, "Employee Records");

                sheet1.Column(1).Style.Font.FontColor = XLColor.Red;
                sheet1.Columns(2, 4).Style.Font.FontColor = XLColor.Blue;

                sheet1.Row(1).CellsUsed().Style.Fill.BackgroundColor = XLColor.Black;
                sheet1.Row(1).Style.Font.FontColor = XLColor.White;

                sheet1.Rows(2, 4).Style.Font.FontColor = XLColor.AshGrey;

                using (MemoryStream ms = new MemoryStream())
                {
                    wb.SaveAs(path);
                    return ms.ToArray();
                }
                
            }
            
        }
        private static DataTable GetEmployeeData() 
        {
            DataTable dt = new DataTable();
            dt.TableName = "EmployeeeData";
            dt.Columns.Add("Name", typeof(string));
            dt.Columns.Add("Age", typeof(int));
            dt.Columns.Add("Address", typeof(string));
            dt.Columns.Add("Email", typeof(string));
            dt.Columns.Add("Salary", typeof(int));
            dt.Columns.Add("Birthday", typeof(string));

            var employees = new List<Employee>()
            {
                new Employee{Name="Andres",Age=18,Address="Carrera 22",Email="andres@gmail.com",Salary=1500000,Birthday="2020-12-12"},
                new Employee{Name="Ana",Age=19,Address="Calle 45", Email = "andres@gmail.com", Salary = 1500000, Birthday = "2020-12-12"},
                new Employee{Name="Camilo",Age=25,Address="Carrera 33", Email = "andres@gmail.com", Salary = 1500000, Birthday = "2020-12-12"},
                new Employee{Name="Daniela",Age=34,Address="Avenida 22", Email = "andres@gmail.com", Salary = 1500000, Birthday = "2020-12-12"},
                new Employee{Name="Juan",Age=65,Address="Carrera 57", Email = "andres@gmail.com", Salary = 1500000, Birthday = "2020-12-12"},
                new Employee{Name="Jorge",Age=15,Address="Carrera 100", Email = "andres@gmail.com", Salary = 1500000, Birthday = "2020-12-12"},
                new Employee{Name="Gabriela",Age=87,Address="Carrera 44", Email = "andres@gmail.com", Salary = 1500000, Birthday = "2020-12-12"},
                new Employee{Name="Zaul",Age=56,Address="Avenida 93", Email = "andres@gmail.com", Salary = 1500000, Birthday = "2020-12-12"},
                new Employee{Name="Xavier",Age=46,Address="Calle 21", Email = "andres@gmail.com", Salary = 1500000, Birthday = "2020-12-12"}
            };

            if (employees.Count > 0) 
            {
                employees.ForEach(item => 
                { 
                    dt.Rows.Add(item.Name, item.Age, item.Address, item.Email, item.Salary, item.Birthday);
                });
            }

            return dt;
        }

    }
}
