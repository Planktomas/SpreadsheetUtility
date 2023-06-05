using System.Reflection;

namespace SpreadsheetUtility
{
    public enum Flow
    {
        Horizontal,
        Vertical,
    }

    /// <summary>
    /// Makes collection have vertical layout in the sheet.
    /// </summary>
    [AttributeUsage(AttributeTargets.Class, AllowMultiple = false, Inherited = true)]
    public class LayoutAttribute : Attribute
    {
        public Flow Direction { get; }

        /// <summary>
        /// Makes collection have vertical layout in the sheet.
        /// </summary>
        public LayoutAttribute(Flow direction)
        {
            Direction = direction;
        }
    }

    /// <summary>
    /// Excludes property when writing and reading sheets. 
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = true)]
    public class HiddenAttribute : Attribute
    {
        public string[] SheetNames { get; }

        /// <summary>
        /// Excludes property when writing and reading sheets. 
        /// </summary>
        /// <param name="sheetNames">Names of sheets where this property is going to be excluded from.</param>
        public HiddenAttribute(params string[] sheetNames)
        {
            SheetNames = sheetNames;
        }
    }

    class LayoutScope : IDisposable
    {
        const Flow k_DefaultFlow = Flow.Horizontal;

        internal static Flow s_Flow = k_DefaultFlow;

        readonly Flow m_PreviousFlow;

        public LayoutScope(Type type)
        {
            var attribute = type.GetCustomAttribute(typeof(LayoutAttribute)) as LayoutAttribute;

            if (attribute == null)
                return;

            m_PreviousFlow = s_Flow;
            s_Flow = attribute.Direction;
        }

        public void Dispose()
        {
            s_Flow = m_PreviousFlow;
        }
    }
}
