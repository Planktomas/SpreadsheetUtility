using SpreadsheetLight;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;

namespace SpreadsheetUtility
{
    public class Spreadsheet : IDisposable
    {
        string m_Path;
        bool m_IsDirty;
        internal SLDocument m_Document;

        ObjectDisposedException disposedException =>
            new ObjectDisposedException("This spreadsheet has been disposed");

        public Spreadsheet(string path)
        {
            m_Path = path;

            if (File.Exists(m_Path))
                m_Document = new SLDocument(m_Path);
            else
                m_Document = new SLDocument();
        }

        ~Spreadsheet()
        {
            Dispose();
        }

        public void Dispose()
        {
            if(m_Document == null)
                return;

            if(m_IsDirty)
            {
                AutoFit();

                if (File.Exists(m_Path))
                    m_Document.Save();
                else
                    m_Document.SaveAs(m_Path);
            }

            m_Document.Dispose();
            m_Document = null;
        }

        public IEnumerable<T> Read<T>(Type worksheetType, string column)
        {
            if (m_Document == null)
                throw disposedException;

            var data = ReadData(worksheetType, column);

            return data[0].Select(d => 
                ((T)Convert.ChangeType(d, typeof(T), CultureInfo.InvariantCulture)));
        }

        public IEnumerable<(T1, T2)> Read<T1, T2>(Type worksheetType, string column1, string column2)
        {
            if (m_Document == null)
                throw disposedException;

            var data = ReadData(worksheetType, column1, column2);

            return data[0].Zip(data[1], (x, y) => 
                (((T1)Convert.ChangeType(x, typeof(T1), CultureInfo.InvariantCulture)),
                ((T2)Convert.ChangeType(y, typeof(T2), CultureInfo.InvariantCulture))));
        }

        public IEnumerable<(T1, T2, T3)> Read<T1, T2, T3>(Type worksheetType, string column1,
            string column2, string column3)
        {
            if (m_Document == null)
                throw disposedException;

            var data = ReadData(worksheetType, column1, column2, column3);

            var zip1 = data[0].Zip(data[1], (x, y) =>
                (((T1)Convert.ChangeType(x, typeof(T1), CultureInfo.InvariantCulture)),
                ((T2)Convert.ChangeType(y, typeof(T2), CultureInfo.InvariantCulture))));

            return zip1.Zip(data[2], (x, y) =>
                (x.Item1, x.Item2,
                ((T3)Convert.ChangeType(y, typeof(T3), CultureInfo.InvariantCulture))));
        }

        public IEnumerable<(T1, T2, T3, T4)> Read<T1, T2, T3, T4>(Type worksheetType, string column1,
            string column2, string column3, string column4)
        {
            if (m_Document == null)
                throw disposedException;

            var data = ReadData(worksheetType, column1, column2, column3, column4);

            var zip1 = data[0].Zip(data[1], (x, y) =>
                (((T1)Convert.ChangeType(x, typeof(T1), CultureInfo.InvariantCulture)),
                ((T2)Convert.ChangeType(y, typeof(T2), CultureInfo.InvariantCulture))));

            var zip2 = data[2].Zip(data[3], (x, y) =>
                (((T3)Convert.ChangeType(x, typeof(T3), CultureInfo.InvariantCulture)),
                ((T4)Convert.ChangeType(y, typeof(T4), CultureInfo.InvariantCulture))));

            return zip1.Zip(zip2, (x, y) =>
                (x.Item1, x.Item2, y.Item1, y.Item2));
        }

        public void Write<TSource>(IEnumerable<TSource> source)
        {
            if (m_Document == null)
                throw disposedException;

            m_IsDirty = true;
            var worksheet = typeof(TSource).Name;
            var properties = typeof(TSource)
                .GetProperties(System.Reflection.BindingFlags.Instance
                | System.Reflection.BindingFlags.Public);

            var worksheets = m_Document.GetWorksheetNames();

            if (worksheets.Count < 2 && !m_Document.AddWorksheet(SLDocument.DefaultFirstSheetName))
                m_Document.SelectWorksheet(SLDocument.DefaultFirstSheetName);

            var wasDefaultWorksheet = m_Document.GetCurrentWorksheetName() == SLDocument.DefaultFirstSheetName;

            if (worksheets.Any(w => w == worksheet))
                m_Document.DeleteWorksheet(worksheet);

            m_Document.AddWorksheet(worksheet);

            if(wasDefaultWorksheet)
                m_Document.DeleteWorksheet(SLDocument.DefaultFirstSheetName);

            WriteHeaders(properties);
            WriteData(properties, source);
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

        void WriteHeaders(PropertyInfo[] properties)
        {
            for (int i = 0; i < properties.Length; i++)
                m_Document.SetCellValue(Cell(i, 0), properties[i].Name);
        }

        void WriteData<TSource>(PropertyInfo[] properties, IEnumerable<TSource> source)
        {
            for (int y = 0; y < source.Count(); y++)
            {
                var entry = source.ElementAt(y);
                WriteEntry(y + 1, properties, entry);
            }
        }

        void WriteEntry<TSource>(int row, PropertyInfo[] properties, TSource source)
        {
            for (int x = 0; x < properties.Length; x++)
            {
                var value = (string)Convert.ChangeType(properties[x].GetValue(source),
                    typeof(string), CultureInfo.InvariantCulture);

                if (properties[x].PropertyType == typeof(string))
                    m_Document.SetCellValue(Cell(x, row), value);
                else
                    m_Document.SetCellValueNumeric(Cell(x, row), value);
            }
        }

        List<string>[] ReadData(Type worksheetType, params string[] columns)
        {
            var worksheetName = worksheetType.Name;

            if (!m_Document.GetWorksheetNames().Contains(worksheetName))
                return null;

            m_Document.SelectWorksheet(worksheetName);

            var columnIndices = GetColumnIndices(columns);
            return GetData(columns, columnIndices);
        }

        int[] GetColumnIndices(string[] columns)
        {
            var result = new int[columns.Length];

            for (int i = 0; true; i++)
            {
                var value = m_Document.GetCellValueAsString(Cell(i, 0));

                if (string.IsNullOrEmpty(value))
                    break;

                for (int j = 0; j < columns.Length; j++)
                {
                    if (value == columns[j])
                        result[j] = i;
                }
            }

            return result;
        }

        List<string>[] GetData(string[] columns, int[] columnIndices)
        {
            var result = new List<string>[columns.Length];

            for (int y = 1; true; y++)
            {
                var rowValues = new string[columnIndices.Length];

                for (int i = 0; i < columnIndices.Length; i++)
                    rowValues[i] = m_Document.GetCellValueAsString(Cell(columnIndices[i], y));

                if (rowValues.All(v => string.IsNullOrEmpty(v)))
                    break;

                for (int i = 0; i < result.Length; i++)
                {
                    if (result[i] == null)
                        result[i] = new List<string>();

                    result[i].Add(rowValues[i]);
                }
            }

            return result;
        }

        internal void AutoFit()
        {
            foreach (var worksheet in m_Document.GetWorksheetNames())
            {
                m_Document.SelectWorksheet(worksheet);

                for (int x = 0; true; x++)
                {
                    if (!m_Document.HasCellValue(Cell(x, 0)))
                        break;

                    m_Document.AutoFitColumn(Cell(x, 0));
                }
            }
        }
    }
}
