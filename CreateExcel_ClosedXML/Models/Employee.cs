namespace CreateExcel_ClosedXML.Models
{
    public class Employee
    {
        public string Name { get; set; } = default!;
        public int Age { get; set; }
        public string Address { get; set; } = default!;
        public string Email { get; set; } = default!;
        public int Salary { get; set; }
        public string Birthday { get; set; } = default!;
    }
}
