# SpreadsheetUtility
Ultra lightweight spreadsheet utility to display processed collections of data and occasionally reading it

#### Features
+ Uses XLSX file format
+ Writes public properties of a collection into a dedicated worksheet
+ Reads worksheet data into tuple enumerator
+ Supports multiple worksheets
+ Auto fits columns for comfortable viewing

#### Preview
<img src="https://user-images.githubusercontent.com/94010480/235344261-6c207066-a73a-4abd-9ac4-7c0eec31ff17.png" width="500" height="300" />

```cs
class TestObject
{
    public string Name { get; set; }
    public string Description { get; set; }
    public decimal Price { get; set; }

    public static TestObject[] Make(int count)
    {
        var result = new TestObject[count];

        for (int i = 0; i < result.Length; i++)
        {
            result[i] = new TestObject()
            {
                Name = i.ToString(),
                Description = Guid.NewGuid().ToString(),
                Price = i * i * i,
            };
        }

        return result;
    }
}

static void Main(string[] args)
{
    var spreadsheet = new Spreadsheet("test.xlsx");
    spreadsheet.Write(TestObject.Make(10));

    IEnumerable<(string Description, decimal Price)> info = spreadsheet.Read<string, decimal>(
        typeof(TestObject), nameof(TestObject.Description), nameof(TestObject.Price));

    foreach (var item in info)
        Console.WriteLine($"Description: {item.Description} \t Price: {item.Price}");

    Console.WriteLine();

    // Alternative reading though this does not attach nice property names to results
    var info2 = spreadsheet.Read<string, decimal>(
        typeof(TestObject), nameof(TestObject.Description), nameof(TestObject.Price));

    foreach (var item in info2)
        Console.WriteLine($"Description: {item.Item1} \t Price: {item.Item2}");
}
```
