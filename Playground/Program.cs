using SpreadsheetUtility;
using System;
using System.Collections.Generic;

namespace Playground
{
    internal class Program
    {
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
    }
}
