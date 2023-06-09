# SpreadsheetUtility
Ultra lightweight spreadsheet utility to display processed collections of data and occasionally reading it

![license](https://img.shields.io/github/license/planktomas/spreadsheetutility.svg)
![GitHub release (latest by date)](https://img.shields.io/github/v/release/planktomas/spreadsheetutility)

### Features
+ Uses XLSX file format
+ Writes public properties of a collection into a dedicated sheet
+ Reads sheet data into an enumerator(List)
+ Supports multiple sheets
+ Auto fits columns for comfortable viewing
+ Can set a startup sheet
+ Supports type independent sheet names
+ Supports custom string formatting
+ Supports color scale formatting
+ Supports horizontal and vertical data layout
+ Can exclude specific properties from writing to the spreadsheet
+ Supports formulas referencing values in the same entry

### Tutorial

Let's create an employee class to store in a spreadsheet.

```cs
class Employee
{
    public string? Name { get; set; }
    public string? Position { get; set; }

    [Format("0$")]
    [ColorScale("red", "#00FF00" /* green */)]
    public decimal Salary { get; set; }

    public Employee() { }

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
    new Employee("Steve", "Manager", 6000),
    new Employee("Will", "Senior Software Engineer", 4000),
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

<img src="https://github.com/Planktomas/SpreadsheetUtility/assets/94010480/155379da-b753-4069-a057-4022192345e5.png" width="350" height="220" />

And if we need to read some of that data back, we can do it too.

```cs
using (var spreadsheet = new Spreadsheet("Company.xlsx"))
{
    foreach (var employee in spreadsheet.Read<Employee>())
    {
        Console.WriteLine($"Salary: {employee.Salary} \t Position: {employee.Position}");
    }
}
```

<img src="https://github.com/Planktomas/SpreadsheetUtility/assets/94010480/5354153c-b40e-436d-9619-9652f3082cc0.png" width="520" height="160" />

[You can review the whole tutorial here](https://github.com/Planktomas/SpreadsheetUtility/blob/main/Tutorial/Program.cs)

### Additional features
#### Layout
By default all sheets will have a horizontal data layout but we can change it to vertical using Layout attribute.

```cs
[Layout(Flow.Vertical)]
class Employee
{
    public string? Name { get; set; }
    public string? Position { get; set; }

    [Format("0$")]
    [ColorScale("red", "#00FF00" /* green */)]
    public decimal Salary { get; set; }

    public Employee() { }

    public Employee(string name, string position, decimal salary)
    {
        Name = name;
        Position = position;
        Salary = salary;
    }
}
```

Here is how it looks in the spreadsheet.

<img src="https://github.com/Planktomas/SpreadsheetUtility/assets/94010480/7aff29ad-88f5-413e-8be6-bb6d73773327.png" width="700" height="160" />

#### Hidden attribute

If there is no need to export a property to the spreadsheet we can exclude it via Hidden attribute.

```cs
class Employee
{
    [Hidden]
    public string? Name { get; set; }
    public string? Position { get; set; }
    ...
}
```

<img src="https://github.com/Planktomas/SpreadsheetUtility/assets/94010480/853d573e-a25a-40e3-a65a-5c50c7ddbcbc.png" width="360" height="180" />

#### Formula

Sometimes we want to have cells that update in real time or react to the changes we make in the spreadsheet. For this case we can use formulas. Keep in mind though that formulas can only reference properties in the same line. Also note that we don't declare a setter for formula property as we don't really need to read the formula back.

```cs

class Employee
{
    ...
    [Format("0$")]
    [ColorScale("red", "#00FF00" /* green */)]
    public decimal Salary { get; set; }

    public string DesiredSalary => $"= {nameof(Salary)} * 2";
    ...
}
```

<img src="https://github.com/Planktomas/SpreadsheetUtility/assets/94010480/1440bcbf-e5b4-417b-be68-80d2a64afd5a.png" width="400" height="180" />
