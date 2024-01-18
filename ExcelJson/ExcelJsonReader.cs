using System.Globalization;
using ExcelDataReader;

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
            // Read header row.
            reader.Read();
            int fieldCount = reader.FieldCount;
            var definitions = new List<string>(fieldCount);
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
                        throw new ArrayPlaceholderException(sheetName, i);
                    }
                    if (headerHashSet.TryGetValue(prevName, out var array))
                    {
                        ++array.Length; 
                    }
                    continue;
                }
                var nameAndType = definition.Split(m_Options.TypeToken, StringSplitOptions.TrimEntries);
                // In most cases, length is 2.
                if (nameAndType.Length != 2)
                {
                    throw new TypeTokenException(sheetName, definition, m_Options.TypeToken);
                }
                var type = nameAndType[1];
                var name = nameAndType[0];
                var headerField = new HeaderField(type, name, 1);
                if (!headerHashSet.TryAdd(name, headerField))
                {
                    throw new DuplicatedException(sheetName, definition);
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
                if (IsEmptyRow(reader, definitionCount))
                {
                    ++emptyRowCount;
                    continue;
                }
                for (int i = 0; i < emptyRowCount; ++i)
                {
                    yield return new string[definitionCount];
                }
                emptyRowCount = 0;
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

        bool IsEmptyRow(IExcelDataReader reader, int definitionCount)
        {
            for (var i = 0; i < definitionCount; ++i)
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