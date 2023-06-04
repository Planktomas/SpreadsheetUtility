using SpreadsheetLight;
using System.Drawing;
using System.Reflection;

namespace SpreadsheetUtility
{
    /// <summary>
    /// Applies specific format to column. 
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = true)]
    public class FormatAttribute : Attribute
    {
        public string FormatCode { get; }

        /// <summary>
        /// Applies specific format to column. 
        /// <see cref="https://support.microsoft.com/en-au/office/number-format-codes-5026bbd6-04bc-48cd-bf33-80f18b4eae68">See Microsoft format code documentation.</see>
        /// </summary>
        public FormatAttribute(string formatCode)
        {
            FormatCode = formatCode;
        }
    }

    /// <summary>
    /// Applies color scale formatting to column.
    /// Able to apply both HTML color codes (For example: #FF00FF)
    /// and color names (For example: AliceBlue).
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = true)]
    public class ColorScaleAttribute : Attribute
    {
        const string k_NeutralColor = "#FFFFFF";

        public string ColorLow { get; }
        public string ColorMiddle { get; } = k_NeutralColor;
        public string ColorHigh { get; }

        /// <summary>
        /// Applies color scale formatting to column.
        /// Able to apply both HTML color codes (For example: #FF00FF)
        /// and color names (For example: AliceBlue).
        /// <see cref="https://learn.microsoft.com/en-us/power-platform/power-fx/reference/function-colors">See Microsoft color value documentation.</see>
        /// </summary>
        public ColorScaleAttribute(string colorLow, string colorHigh)
        {
            ColorLow = colorLow;
            ColorHigh = colorHigh;
        }

        /// <summary>
        /// Applies color scale formatting to column.
        /// Able to apply both HTML color codes (For example: #FF00FF)
        /// and color names (For example: AliceBlue).
        /// <see cref="https://learn.microsoft.com/en-us/power-platform/power-fx/reference/function-colors">See Microsoft color value documentation.</see>
        /// </summary>
        public ColorScaleAttribute(string colorLow, string colorMiddle, string colorHigh)
        {
            ColorLow = colorLow;
            ColorMiddle = colorMiddle;
            ColorHigh = colorHigh;
        }
    }

    internal static class StyleExtensions
    {
        public static void ApplySheetAttributes<T>(this Spreadsheet spreadsheet, List<PropertyInfo> properties)
        {
            for (int i = 0; i < properties.Count; i++)
            {
                var attributes = properties[i].GetCustomAttributes(true);

                foreach (var attribute in attributes)
                {
                    switch (attribute)
                    {
                        case FormatAttribute formatAttribute:
                            var style = spreadsheet.Document.GetColumnStyle(i + 1);
                            style.FormatCode = formatAttribute.FormatCode;

                            if(Spreadsheet.IsVerticalFlow)
                                spreadsheet.Document.SetRowStyle(i + 1, style);
                            else
                                spreadsheet.Document.SetColumnStyle(i + 1, style);
                            break;

                        case ColorScaleAttribute colorScaleAttribute:

                            SLConditionalFormatting? colorScale = null;
                            if(Spreadsheet.IsVerticalFlow)
                                colorScale = new SLConditionalFormatting(i + 1, 0, i + 1, int.MaxValue);
                            else
                                colorScale = new SLConditionalFormatting(0, i + 1, int.MaxValue, i + 1);

                            colorScale?.SetCustom3ColorScale(
                                SLConditionalFormatMinMaxValues.Percentile, "0", ColorTranslator.FromHtml(colorScaleAttribute.ColorLow),
                                SLConditionalFormatRangeValues.Percentile, "50", ColorTranslator.FromHtml(colorScaleAttribute.ColorMiddle),
                                SLConditionalFormatMinMaxValues.Percentile, "100", ColorTranslator.FromHtml(colorScaleAttribute.ColorHigh));
                            spreadsheet.Document.AddConditionalFormatting(colorScale);
                            break;
                    }
                }
            }
        }

        public static void AutoFit(this Spreadsheet spreadsheet)
        {
            foreach (var worksheet in spreadsheet.Document.GetWorksheetNames())
            {
                spreadsheet.Document.SelectWorksheet(worksheet);

                for (int x = 0; true; x++)
                {
                    if (!spreadsheet.Document.HasCellValue(Spreadsheet.Cell(x, 0)))
                        break;

                    spreadsheet.Document.AutoFitColumn(Spreadsheet.Cell(x, 0));
                }
            }
        }
    }
}
