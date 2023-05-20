using SpreadsheetLight;
using System.Globalization;
using System.Reflection;

namespace SpreadsheetUtility
{
    public class Spreadsheet : IDisposable
    {
        static ObjectDisposedException m_DisposedException = new($"Spreadsheet has been disposed");

        readonly string m_Path;

        SLDocument? m_Document;

        internal SLDocument document
        {
            get
            {
                if (m_Document == null)
                    throw m_DisposedException;

                return m_Document;
            }
        }

        /// <summary>
        /// Creates or opens an XLSX format spreadsheet.
        /// </summary>
        /// <param name="path">Path to file.</param>
        public Spreadsheet(string path)
        {
            m_Path = path;

            if (File.Exists(m_Path))
                m_Document = new SLDocument(m_Path);
            else
                m_Document = new SLDocument();
        }

        public void Dispose()
        {
            this.AutoFit();
            document.SaveAs(m_Path);
            document.Dispose();
            m_Document = null;

            GC.SuppressFinalize(this);
        }

        internal static string Cell(int x, int y)
        {
            const int Range = 'Z' - 'A' + 1;

            var coord = string.Empty;

            do
            {
                coord = (char)('A' + (x % Range)) + coord;
                x /= Range;
                x -= 1;
            } while (x >= 0);

            return coord + (y + 1);
        }

        /// <summary>
        /// Reads the worksheet of the type provided.
        /// </summary>
        /// <typeparam name="T">Type identifying a worksheet</typeparam>
        /// <returns>A collection of provided type instances with public instance properties having assigned read values.</returns>
        public IEnumerable<T>? Read<T>()
        {
            var worksheet = typeof(T).Name;
            var properties = typeof(T).GetProperties().ToDictionary(p => p.Name);

            if (!SelectWorksheet(worksheet, false))
                return null;

            var worksheetProperties = GetWorksheetProperties<T>();
            return ReadData<T>(worksheetProperties);
        }

        /// <summary>
        /// Creates or updates worksheet with data from the collection provided.
        /// </summary>
        /// <typeparam name="T">Type identifying worksheet.</typeparam>
        /// <param name="source">Collection to be used as data source.</param>
        public void Write<T>(IEnumerable<T> source)
        {
            var worksheet = typeof(T).Name;
            var properties = typeof(T).GetProperties();

            SelectAndClearWorksheet(worksheet);
            WriteHeaders(properties);
            WriteData(properties, source);
            this.ApplyWorksheetAttributes<T>(properties);
        }

        bool SelectWorksheet(string name, bool canCreate = true)
        {
            var selectResult = document.SelectWorksheet(name);

            if (!canCreate || selectResult)
                return selectResult;

            return document.AddWorksheet(name);
        }

        void SelectAndClearWorksheet(string name)
        {
            document.AddWorksheet(SLDocument.DefaultFirstSheetName);

            document.DeleteWorksheet(name);
            SelectWorksheet(name);

            document.DeleteWorksheet(SLDocument.DefaultFirstSheetName);
        }

        void WriteHeaders(PropertyInfo[] properties)
        {
            for (int i = 0; i < properties.Length; i++)
                document.SetCellValue(Cell(i, 0), properties[i].Name);
        }

        void WriteData<T>(PropertyInfo[] properties, IEnumerable<T> source)
        {
            for (int y = 0; y < source.Count(); y++)
                WriteEntry(y + 1, properties, source.ElementAt(y));
        }

        void WriteEntry<T>(int row, PropertyInfo[] properties, T source)
        {
            for (int x = 0; x < properties.Length; x++)
            {
                var value = (string?)Convert.ChangeType(properties[x].GetValue(source),
                    typeof(string), CultureInfo.InvariantCulture);

                if (properties[x].PropertyType == typeof(string))
                    document.SetCellValue(Cell(x, row), value);
                else
                    document.SetCellValueNumeric(Cell(x, row), value);
            }
        }

        Dictionary<PropertyInfo, int> GetWorksheetProperties<T>()
        {
            var typeProperties = typeof(T).GetProperties();
            var worksheetProperties = new Dictionary<PropertyInfo, int>();

            for (var i = 0; true; i++)
            {
                var label = document.GetCellValueAsString(Cell(i, 0));

                if (string.IsNullOrEmpty(label))
                    break;

                var labelProperty = typeProperties.FirstOrDefault(p => p.Name == label);

                if (labelProperty == null)
                    continue;

                worksheetProperties[labelProperty] = i;
            }

            return worksheetProperties;
        }

        IEnumerable<T> ReadData<T>(Dictionary<PropertyInfo, int> properties)
        {
            var data = new List<T>(document.GetWorksheetStatistics().EndRowIndex + 1);

            for (int y = 1; y < document.GetWorksheetStatistics().EndRowIndex; y++)
            {
                T entry = Activator.CreateInstance<T>();

                foreach (var property in properties)
                {
                    var value = document.GetCellValueAsString(Cell(property.Value, y));

                    property.Key.SetValue(entry, Convert.ChangeType(value,
                        property.Key.PropertyType, CultureInfo.InvariantCulture));
                }

                data.Add(entry);
            }

            return data;
        }
    }
}
