using SpreadsheetLight;
using System.Reflection;

namespace SpreadsheetUtility
{
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
    public class FormatAttribute : Attribute
    {
        public string FormatCode { get; }

        public FormatAttribute(string formatCode)
        {
            FormatCode = formatCode;
        }
    }

    internal static class StyleExtensions
    {
        public static void ApplyWorksheetAttributes<T>(this Spreadsheet spreadsheet, PropertyInfo?[] properties)
        {
            for (int i = 0; i < properties.Length; i++)
            {
                var attributes = properties[i].GetCustomAttributes(true);

                foreach (var attribute in attributes)
                {
                    switch (attribute)
                    {
                        case FormatAttribute formatAttribute:
                            SLStyle style = spreadsheet.document.GetColumnStyle(i + 1);
                            style.FormatCode = formatAttribute.FormatCode;
                            spreadsheet.document.SetColumnStyle(i + 1, style);
                            break;
                    }
                }
            }
        }

        public static void AutoFit(this Spreadsheet spreadsheet)
        {
            foreach (var worksheet in spreadsheet.document.GetWorksheetNames())
            {
                spreadsheet.document.SelectWorksheet(worksheet);

                for (int x = 0; true; x++)
                {
                    if (!spreadsheet.document.HasCellValue(Spreadsheet.Cell(x, 0)))
                        break;

                    spreadsheet.document.AutoFitColumn(Spreadsheet.Cell(x, 0));
                }
            }
        }
    }
}
