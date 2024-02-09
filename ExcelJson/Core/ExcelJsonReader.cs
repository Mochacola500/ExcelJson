using ExcelDataReader;

namespace ExcelJson.Core
{
    public class ExcelJsonReader : IDisposable
    {
        readonly ExcelJsonOptions m_ExcelJsonOptions;
        readonly IExcelDataReader m_ExcelDataReader;

        public ExcelJsonReader(IExcelDataReader excelDataReader, ExcelJsonOptions excelJsonOptions)
        {
            m_ExcelDataReader = excelDataReader;
            m_ExcelJsonOptions = excelJsonOptions;
        }

        public IEnumerable<ExcelJsonSheet> ReadExcel()
        {
            m_ExcelDataReader.Reset();
            do
            {
                var sheetName = m_ExcelDataReader.Name;
                if (sheetName.StartsWith(m_ExcelJsonOptions.FilterSheetToken))
                {
                    continue;
                }
                yield return ReadCurrentSheet();
            } while (m_ExcelDataReader.NextResult());
        }

        public void Dispose()
        {
            m_ExcelDataReader.Dispose();
        }

        ExcelJsonSheet ReadCurrentSheet()
        {
            m_ExcelDataReader.Read();
            var definitions = ReadDefinitions();
            var header = ConvertToHeader(definitions);
            var rows = ReadRows(definitions.Length);
            return new(m_ExcelDataReader.Name, header, rows);
        }

        ExcelHeader ConvertToHeader(string[] definitions)
        {
            int totalCount = definitions.Length;
            var headerLookup = new Dictionary<string, HeaderField>(totalCount);
            var sheetName = m_ExcelDataReader.Name;
            var prevDefinition = string.Empty;

            for (int i = 0; i < totalCount; ++i)
            {
                var definition = definitions[i];
                if (definition.StartsWith(m_ExcelJsonOptions.FilterColumnToken))
                {
                    continue;
                }
                if (definition == m_ExcelJsonOptions.ArrayPlaceholder)
                {
                    if (prevDefinition == "")
                    {
                        throw new ArrayPlaceholderException(sheetName, i);
                    }
                    if (headerLookup.TryGetValue(prevDefinition, out var array))
                    {
                        ++array.ArrayLength; 
                    }
                }
                else
                {
                    var nameAndType = definition.Split(m_ExcelJsonOptions.TypeToken, StringSplitOptions.TrimEntries);
                    // In most cases, length is 2.
                    if (nameAndType.Length != 2)
                    {
                        throw new TypeTokenException(sheetName, definition, m_ExcelJsonOptions.TypeToken);
                    }
                    var type = nameAndType[1];
                    var name = nameAndType[0];
                    var headerField = new HeaderField(type, name, 1);
                    if (!headerLookup.TryAdd(name, headerField))
                    {
                        throw new DuplicatedException(sheetName, definition);
                    }
                    prevDefinition = name;
                }
            }
            var headerArray = headerLookup.Values.ToArray();
            return new(sheetName, headerArray, totalCount);
        }

        string[] ReadDefinitions()
        {
            int fieldCount = m_ExcelDataReader.FieldCount;
            var definitions = new List<string>();
            for (int i = 0; i < fieldCount; ++i)
            {
                var value = m_ExcelDataReader.GetString(i);
                // Separate the data area and non-data using an empty column.
                if (string.IsNullOrEmpty(value))
                {
                    break;
                }
                definitions.Add(value);
            }
            return definitions.ToArray();
        }

        IEnumerable<string[]> ReadRows(int definitionCount)
        {   
            while (m_ExcelDataReader.Read())
            {
                var row = new string[definitionCount];
                for (int i = 0; i < definitionCount; ++i)
                {
                    var value = m_ExcelDataReader.GetString(i);
                    row[i] = value ?? "";
                }
                yield return row;
            }
        }
    }
}