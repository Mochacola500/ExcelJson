using ExcelDataReader;
using System.Globalization;

namespace ExcelJson
{
    public abstract class ExcelJsonReader
    {
        protected readonly ExcelJsonOptions m_Options;

        protected ExcelJsonReader(ExcelJsonOptions options)
        {
            m_Options = options;
        }

        protected List<string> ReadDefinitions(IExcelDataReader reader)
        {
            int fieldCount = reader.FieldCount;
            var definitions = new List<string>(fieldCount);
            reader.Read(); // Read header row.
            for (int i = 0; i < fieldCount; ++i)
            {
                var value = reader.GetValue(i);
                var name = Convert.ToString(value, CultureInfo.CurrentCulture);
                if (string.IsNullOrEmpty(name))
                {
                    break;
                }
                definitions.Add(name);
            }
            return definitions;
        }

        protected HeaderRow ReadHeaderRow(string sheetName, List<string> definitions)
        {
            int count = definitions.Count;
            var headerHashSet = new Dictionary<string, HeaderField>(count);
            var prevName = "";

            for (int i = 0; i < count; ++i)
            {
                var definition = definitions[i];

                if (definition.StartsWith(m_Options.FilterColumnToken))
                {
                    continue;
                }
                if (definition == m_Options.ArrayPlaceholder)
                {
                    if (prevName == "")
                    {
                        throw new Exception($"Definition is required befor the array token.\nName:{sheetName}\nIndex:{i}");
                    }
                    if (headerHashSet.TryGetValue(prevName, out var array))
                    {
                        ++array.Length;
                    }
                    continue;
                }
                var nameAndType = definition.Split(m_Options.TypeToken, StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries);
                // In most cases, length is 2.
                if (nameAndType.Length != 2)
                {
                    if (nameAndType.Length == 1)
                    {
                        throw new Exception($"Missing type token.\nTableName:{sheetName}\nDefinition:{definition}\nToken:{m_Options.TypeToken}");
                    }
                    else
                    {
                        throw new Exception($"Too many type token exist.\nTableName:{sheetName}\nDefinition:{definition}\nToken:{m_Options.TypeToken}");
                    }
                }
                var type = nameAndType[1];
                var name = nameAndType[0];
                var headerField = new HeaderField(type, name, 1);
                if (!headerHashSet.TryAdd(name, headerField))
                {
                    throw new Exception($"Duplicated definition name.\nTableName:{sheetName}\nDefinition:{definition}");
                }
                prevName = name;
            }
            var headerArray = headerHashSet.Values.ToArray();
            return new(sheetName, headerArray);
        }

        protected IEnumerable<string[]> ReadRows(IExcelDataReader reader, int definitionCount)
        {
            int emptyRowCount = 0;

            while (reader.Read())
            {
                if (IsEmptyRow(reader))
                {
                    ++emptyRowCount;
                    continue;
                }

                for (int i = 0; i < emptyRowCount; ++i)
                {
                    yield return new string[definitionCount];
                }

                var row = new string[definitionCount];
                for (int i = 0; i < definitionCount; ++i)
                {
                    var obj = reader.GetValue(i);
                    var value = Convert.ToString(obj);
                    row[i] = value ?? "";
                }
                yield return row;
            }
        }

        bool IsEmptyRow(IExcelDataReader reader)
        {
            for (var i = 0; i < reader.FieldCount; ++i)
            {
                var value = reader.GetValue(i);
                if (value != null)
                {
                    return false;
                }
            }
            return true;
        }
    }
}