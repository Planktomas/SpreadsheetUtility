using SpreadsheetLight;
using System.Globalization;
using System.Reflection;

namespace SpreadsheetUtility
{
    public class Spreadsheet : IDisposable
    {
        static readonly ObjectDisposedException k_DisposedException = new($"Spreadsheet has been disposed.");
        static readonly ArgumentException k_UnknownSheetException = new("Sheet does not exist in the spreadsheet.");

        readonly string k_Path;

        SLDocument? m_Document;
        string? m_StartupSheet;

        internal SLDocument Document
        {
            get
            {
                if (m_Document == null)
                    throw k_DisposedException;

                return m_Document;
            }
        }

        /// <summary>
        /// Creates or opens an XLSX format spreadsheet.
        /// </summary>
        /// <param name="path">Path to file.</param>
        public Spreadsheet(string path)
        {
            k_Path = path;

            if (File.Exists(k_Path))
                m_Document = new SLDocument(k_Path);
            else
                m_Document = new SLDocument();
        }

        public void Dispose()
        {
            this.AutoFit();

            if(m_StartupSheet != null)
                Document.SelectWorksheet(m_StartupSheet);

            Document.SaveAs(k_Path);
            Document.Dispose();
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
        /// Reads the sheet of the type provided.
        /// </summary>
        /// <typeparam name="T">Type identifying a sheet</typeparam>
        /// <returns>A collection of provided type instances with public instance properties having assigned read values.</returns>
        public IEnumerable<T>? Read<T>()
        {
            var sheetName = typeof(T).Name;
            var properties = typeof(T).GetProperties().ToDictionary(p => p.Name);

            if (!SelectSheet(sheetName, false))
                return null;

            var sheetProperties = GetSheetProperties<T>();
            return ReadData<T>(sheetProperties);
        }

        /// <summary>
        /// Creates or updates sheet with data from the collection provided.
        /// </summary>
        /// <typeparam name="T">Type identifying sheet.</typeparam>
        /// <param name="source">Collection to be used as data source.</param>
        public void Write<T>(IEnumerable<T> source)
        {
            var sheetName = typeof(T).Name;
            var properties = typeof(T).GetProperties();

            SelectAndClearSheet(sheetName);
            WriteHeaders(properties);
            WriteData(properties, source);
            this.ApplySheetAttributes<T>(properties);
        }

        /// <summary>
        /// Sets a sheet that will be selected when opening *.xlsx file.
        /// </summary>
        /// <typeparam name="T">Type identifying sheet.</typeparam>
        public void SetStartupSheet<T>()
        {
            var sheetName = typeof(T).Name;

            if (!Document.GetWorksheetNames().Contains(sheetName))
                throw k_UnknownSheetException;

            m_StartupSheet = sheetName;
        }

        bool SelectSheet(string name, bool canCreate = true)
        {
            var selectResult = Document.SelectWorksheet(name);

            if (!canCreate || selectResult)
                return selectResult;

            return Document.AddWorksheet(name);
        }

        void SelectAndClearSheet(string name)
        {
            Document.AddWorksheet(SLDocument.DefaultFirstSheetName);

            Document.DeleteWorksheet(name);
            SelectSheet(name);

            Document.DeleteWorksheet(SLDocument.DefaultFirstSheetName);
        }

        void WriteHeaders(PropertyInfo[] properties)
        {
            for (int i = 0; i < properties.Length; i++)
                Document.SetCellValue(Cell(i, 0), properties[i].Name);
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
                    Document.SetCellValue(Cell(x, row), value);
                else
                    Document.SetCellValueNumeric(Cell(x, row), value);
            }
        }

        Dictionary<PropertyInfo, int> GetSheetProperties<T>()
        {
            var typeProperties = typeof(T).GetProperties();
            var sheetProperties = new Dictionary<PropertyInfo, int>();

            for (var i = 0; true; i++)
            {
                var label = Document.GetCellValueAsString(Cell(i, 0));

                if (string.IsNullOrEmpty(label))
                    break;

                var labelProperty = typeProperties.FirstOrDefault(p => p.Name == label);

                if (labelProperty == null)
                    continue;

                sheetProperties[labelProperty] = i;
            }

            return sheetProperties;
        }

        IEnumerable<T> ReadData<T>(Dictionary<PropertyInfo, int> properties)
        {
            var data = new List<T>(Document.GetWorksheetStatistics().EndRowIndex + 1);

            for (int y = 1; y < Document.GetWorksheetStatistics().EndRowIndex; y++)
            {
                T entry = Activator.CreateInstance<T>();

                foreach (var property in properties)
                {
                    var value = Document.GetCellValueAsString(Cell(property.Value, y));

                    property.Key.SetValue(entry, Convert.ChangeType(value,
                        property.Key.PropertyType, CultureInfo.InvariantCulture));
                }

                data.Add(entry);
            }

            return data;
        }
    }
}
