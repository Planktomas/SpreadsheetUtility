using SpreadsheetUtility;

var employees = new[]
{
    new Employee("John", "CEO", 10000),
    new Employee("Steve", "Manager", 6000),
    new Employee("Will", "Senior Software Engineer", 4000),
    new Employee("Kate", "Software Engineer", 2000),
    new Employee("Paul", "Quality Assurance", 1000)
};

using (var spreadsheet = new Spreadsheet("Company.xlsx"))
{
    spreadsheet.Write(employees);

    foreach (var employee in spreadsheet.Read<Employee>())
    {
        Console.WriteLine($"Salary: {employee.Salary} \t Position: {employee.Position}");
    }
}

class Employee
{
    public string? Name { get; set; }
    public string? Position { get; set; }

    [Format("0$")]
    public decimal Salary { get; set; }

    public Employee() { }

    public Employee(string name, string position, decimal salary)
    {
        Name = name;
        Position = position;
        Salary = salary;
    }
}