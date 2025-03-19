using ClosedXML.Excel;

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
    }
}
