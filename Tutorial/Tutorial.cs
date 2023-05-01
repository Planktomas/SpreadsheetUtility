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

    IEnumerable<(decimal Salary, string Position)> salaries;

    salaries = spreadsheet.Read<decimal, string>(typeof(Employee),
        nameof(Employee.Salary), nameof(Employee.Position));

    foreach (var item in salaries)
    {
        Console.WriteLine($"Salary: {item.Salary} \t Position: {item.Position}");
    }
}

class Employee
{
    public string Name { get; set; }
    public string Position { get; set; }
    public decimal Salary { get; set; }

    public Employee(string name, string position, decimal salary)
    {
        Name = name;
        Position = position;
        Salary = salary;
    }
}