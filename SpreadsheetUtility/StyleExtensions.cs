using SpreadsheetLight;
using System.Drawing;
using System.Reflection;

namespace SpreadsheetUtility
{
    public abstract class SpreadsheetAttribute : Attribute
    {
        internal abstract void Apply(SLDocument document, int column);
    }

    /// <summary>
    /// Applies specific format to column. 
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = true)]
    public class FormatAttribute : SpreadsheetAttribute
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

        internal override void Apply(SLDocument document, int column)
        {
            var style = document.GetColumnStyle(column);
            style.FormatCode = FormatCode;

            if (Spreadsheet.IsVerticalFlow)
                document.SetRowStyle(column, style);
            else
                document.SetColumnStyle(column, style);
        }
    }

    /// <summary>
    /// Applies color scale formatting to column.
    /// Able to apply both HTML color codes (For example: #FF00FF)
    /// and color names (For example: AliceBlue).
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = true)]
    public class ColorScaleAttribute : SpreadsheetAttribute
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

        internal override void Apply(SLDocument document, int column)
        {
            SLConditionalFormatting? colorScale = null;
            if (Spreadsheet.IsVerticalFlow)
                colorScale = new SLConditionalFormatting(column, 0, column, int.MaxValue);
            else
                colorScale = new SLConditionalFormatting(0, column, int.MaxValue, column);

            colorScale?.SetCustom3ColorScale(
                SLConditionalFormatMinMaxValues.Percentile, "0", ColorTranslator.FromHtml(ColorLow),
                SLConditionalFormatRangeValues.Percentile, "50", ColorTranslator.FromHtml(ColorMiddle),
                SLConditionalFormatMinMaxValues.Percentile, "100", ColorTranslator.FromHtml(ColorHigh));
            document.AddConditionalFormatting(colorScale);
        }
    }

    /// <summary>
    /// Applies tooltip to column name. 
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = true)]
    public class TooltipAttribute : SpreadsheetAttribute
    {
        public string Tooltip { get; }

        /// <summary>
        /// Applies specific format to column. 
        /// <see cref="https://support.microsoft.com/en-au/office/number-format-codes-5026bbd6-04bc-48cd-bf33-80f18b4eae68">See Microsoft format code documentation.</see>
        /// </summary>
        public TooltipAttribute(string tooltip)
        {
            Tooltip = tooltip;
        }

        internal override void Apply(SLDocument document, int column)
        {
            var tooltip = document.CreateComment();
            tooltip.SetText(Tooltip);
            document.InsertComment(Spreadsheet.Cell(column - 1, 0), tooltip);
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
                    if (attribute is not SpreadsheetAttribute spreadsheetAttribute)
                        continue;

                    spreadsheetAttribute.Apply(spreadsheet.Document, i + 1);
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
