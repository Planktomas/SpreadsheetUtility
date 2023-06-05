﻿using SpreadsheetLight;
using System.Globalization;
using System.Reflection;

namespace SpreadsheetUtility
{
    public class Spreadsheet : IDisposable
    {
        static readonly ObjectDisposedException k_DisposedException = new($"Spreadsheet is disposed.");

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

        internal static bool IsVerticalFlow => LayoutScope.s_Flow == Flow.Vertical;

        int Rows
        {
            get
            {
                var statistics = Document.GetWorksheetStatistics();
                return IsVerticalFlow ? statistics.EndColumnIndex : statistics.EndRowIndex;
            }
        }

        int Columns
        {
            get
            {
                var statistics = Document.GetWorksheetStatistics();
                return IsVerticalFlow ? statistics.EndRowIndex : statistics.EndColumnIndex;
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

            if (m_StartupSheet != null)
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

            if(IsVerticalFlow)
            {
                var temp = x;
                x = y;
                y = temp;
            }

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
        /// <typeparam name="T">Type identifying the sheet</typeparam>
        /// <param name="sheetName">Name identifying the sheet.</param>
        /// <returns>A collection of provided type instances with public instance properties having assigned read values.</returns>
        public IEnumerable<T>? Read<T>(string? sheetName = null)
        {
            var name = sheetName ?? typeof(T).Name;

            if (!SelectSheet(name, false))
                return null;

            using var layoutScope = new LayoutScope(typeof(T));
            var properties = GetPropertiesFromSheet<T>();
            return ReadData<T>(properties);
        }

        /// <summary>
        /// Creates or updates sheet with data from the collection provided.
        /// </summary>
        /// <typeparam name="T">Type identifying the sheet.</typeparam>
        /// <param name="source">Collection to be used as data source.</param>
        /// <param name="sheetName">Name identifying the sheet.</param>
        public void Write<T>(IEnumerable<T> source, string? sheetName = null)
        {
            var name = sheetName ?? typeof(T).Name;

            SelectAndClearSheet(name);

            using var layoutScope = new LayoutScope(typeof(T));
            var properties = GetPropertiesFromType<T>();
            WriteData(properties, source);
            this.ApplySheetAttributes<T>(properties);
        }

        /// <summary>
        /// Deletes sheet from spreadsheet.
        /// </summary>
        /// <typeparam name="T">Type identifying the sheet.</typeparam>
        /// <returns>Returns `true` if the operation succeeded. Otherwise `false`.</returns>
        public bool Delete<T>() => Delete(typeof(T).Name);

        /// <summary>
        /// Deletes sheet from spreadsheet.
        /// </summary>
        /// <param name="sheetName">Name identifying the sheet.</param>
        /// <returns>Returns `true` if the operation succeeded. Otherwise `false`.</returns>
        public bool Delete(string sheetName)
        {
            Document.AddWorksheet(SLDocument.DefaultFirstSheetName);
            var result = Document.DeleteWorksheet(sheetName);
            Document.DeleteWorksheet(SLDocument.DefaultFirstSheetName);

            return result;
        }

        /// <summary>
        /// Sets a sheet that will be selected when opening *.xlsx file.
        /// </summary>
        /// <typeparam name="T">Type identifying the sheet.</typeparam>
        public void SetStartupSheet<T>() => SetStartupSheet(typeof(T).Name);

        /// <summary>
        /// Sets a sheet that will be selected when opening *.xlsx file.
        /// </summary>
        /// <param name="sheetName">Type identifying the sheet.</param>
        public void SetStartupSheet(string sheetName)
        {
            // Ensure a valid document
            _ = Document;
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

        List<PropertyInfo> GetPropertiesFromType<T>(string? sheetName = null)
        {
            var name = sheetName ?? typeof(T).Name;
            var properties = typeof(T).GetProperties();
            var sheetProperties = new List<PropertyInfo>();

            foreach (var property in properties)
            {
                var hiddenAttribute = property.GetCustomAttribute<HiddenAttribute>();

                if (hiddenAttribute != null &&
                    (hiddenAttribute.SheetNames.Length == 0 || hiddenAttribute.SheetNames.Contains(name)))
                    continue;

                sheetProperties.Add(property);
            }

            return sheetProperties;
        }

        Dictionary<PropertyInfo, int> GetPropertiesFromSheet<T>(string? sheetName = null)
        {
            var name = sheetName ?? typeof(T).Name;
            var properties = typeof(T).GetProperties();
            var sheetProperties = new Dictionary<PropertyInfo, int>();
            var columnCount = Columns;

            for (var i = 0; i < columnCount; i++)
            {
                var label = Document.GetCellValueAsString(Cell(i, 0));

                if (string.IsNullOrEmpty(label))
                    break;

                var labelProperty = properties.FirstOrDefault(p => p.Name == label);

                if (labelProperty == null)
                    continue;

                var hiddenAttribute = labelProperty.GetCustomAttribute<HiddenAttribute>();

                if (hiddenAttribute != null &&
                    (hiddenAttribute.SheetNames.Length == 0 || hiddenAttribute.SheetNames.Contains(name)))
                    continue;

                sheetProperties[labelProperty] = i;
            }

            return sheetProperties;
        }

        void WriteData<T>(List<PropertyInfo> properties, IEnumerable<T> source)
        {
            for (int i = 0; i < properties.Count(); i++)
                Document.SetCellValue(Cell(i, 0), properties[i].Name);

            for (int y = 0; y < source.Count(); y++)
            {
                for (int x = 0; x < properties.Count(); x++)
                {
                    var row = y + 1;
                    var value = (string?)Convert.ChangeType(properties[x].GetValue(source.ElementAt(y)),
                        typeof(string), CultureInfo.InvariantCulture);

                    if (properties[x].PropertyType == typeof(string))
                    {
                        if (value?[0] == '=')
                            for (int i = 0; i < properties.Count; i++)
                                value = value.Replace(properties[i].Name, Cell(i, row));

                        Document.SetCellValue(Cell(x, row), value);
                    }
                    else
                    {
                        Document.SetCellValueNumeric(Cell(x, row), value);
                    }
                }
            }
        }

        IEnumerable<T> ReadData<T>(Dictionary<PropertyInfo, int> properties)
        {
            var rowCount = Rows;
            var data = new List<T>(rowCount - 1);

            for (int y = 1; y < rowCount; y++)
            {
                T entry = Activator.CreateInstance<T>();

                foreach (var property in properties)
                {
                    var value = Document.GetCellValueAsString(Cell(property.Value, y));

                    if (!property.Key.CanWrite)
                        continue;

                    property.Key.SetValue(entry, Convert.ChangeType(value,
                        property.Key.PropertyType, CultureInfo.InvariantCulture));
                }

                data.Add(entry);
            }

            return data;
        }
    }
}
