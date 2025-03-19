
string pathFile = AppDomain.CurrentDomain.BaseDirectory + "myExcel.xlsx";

SLDocument oSLDocument = new SLDocument();

System.Data.DataTable dt = new System.Data.DataTable();


//Add Columns
dt.Columns.Add("Name", typeof(string));
dt.Columns.Add("Age", typeof(int));
dt.Columns.Add("Gender", typeof(string));

//Add Rows
dt.Rows.Add("Pedro", 13, "Hombre");
dt.Rows.Add("Ana", 23, "Mujer");
dt.Rows.Add("Luis", 33, "Hombre");
dt.Rows.Add("Lina", 13, "Mujer");
dt.Rows.Add("Leon", 43, "Hombre");

//Generate matrix excel init write
oSLDocument.ImportDataTable(1, 1, dt, true);

//Save document
oSLDocument.SaveAs(pathFile);
