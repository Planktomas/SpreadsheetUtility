# SpreadsheetUtility
Ultra lightweight spreadsheet utility to display processed collections of data and occasionally reading it

#### Features
+ Uses XLSX file format
+ Writes public properties of a collection into a dedicated worksheet
+ Reads worksheet data into tuple enumerator
+ Supports multiple worksheets
+ Auto fits columns for comfortable viewing

#### Tutorial

Let's create an employee class to store in a spreadsheet.

```cs
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
```

Now we can make an array of company's employees.

```cs
var employees = new[]
{
    new Employee("John", "CEO", 10000),
    new Employee("Kate", "Software Engineer", 2000),
    new Employee("Paul", "Quality Assurance", 1000)
};
```

This array can now go into the spreadsheet.

```cs
using (var spreadsheet = new Spreadsheet("Company.xlsx"))
{
    spreadsheet.Write(employees);
}
```

Here is how this data looks in the spreadsheet.

<img src="https://user-images.githubusercontent.com/94010480/235367459-c488f500-2f01-440e-9653-e3a8f895550d.png" width="350" height="220" />

And if we need to read some of that data back, we can do it too.

```cs
using (var spreadsheet = new Spreadsheet("Company.xlsx"))
{
    IEnumerable<(decimal Salary, string Position)> salaries;

    salaries = spreadsheet.Read<decimal, string>(typeof(Employee),
        nameof(Employee.Salary), nameof(Employee.Position));
}
```

We can verify the results just to be sure we got the right data.

```cs
foreach (var item in salaries)
{
    Console.WriteLine($"Salary: {item.Salary} \t Position: {item.Position}");
}
```

<img src="https://user-images.githubusercontent.com/94010480/235367385-1bedc612-0d15-410e-b262-cb82b61601ae.png" width="400" height="90" />
