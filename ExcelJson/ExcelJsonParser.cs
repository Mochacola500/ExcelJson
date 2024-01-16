using ExcelDataReader;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json;

namespace ExcelJson
{
    public class ExcelJsonParser : ExcelJsonReader
    {
        public ExcelJsonParser() : this(new())
        {

        }

        public ExcelJsonParser(ExcelJsonOptions options) : base(options)
        {

        }

        public IEnumerable<ExcelJsonSheet> ReadExcel(Stream stream)
        {
            using var reader = ExcelReaderFactory.CreateReader(stream);
            reader.Reset();
            do
            {
                var sheetName = reader.Name;
                if (sheetName.StartsWith(m_Options.FilterSheetToken))
                {
                    continue;
                }
                yield return ReadExcelJsonSheet(reader, sheetName);
            } while (reader.NextResult());
        }

        ExcelJsonSheet ReadExcelJsonSheet(IExcelDataReader reader, string sheetName)
        {
            var definitions = ReadDefinitions(reader);
            var headerRow = ReadHeaderRow(sheetName, definitions);
            var rows = ReadRows(reader, definitions.Count);

            var items = new Dictionary<string, JObject>();
            var fields = headerRow.Fields;
            foreach (var row in rows)
            {
                var jObject = new JObject();
                for (int i = 0; i < fields.Length; ++i)
                {
                    var field = fields[i];
                    if (field.Length == 1)
                    {
                        AddObject(jObject, row[i], field);
                    }
                    else
                    {
                        var jArray = new JArray();
                        for (int j = i; j < field.Length; ++j)
                        {
                            AddPrimitiveToArray(jArray, field.Type, row[j]);
                        }
                        jObject.Add(field.Name, jArray);
                    }
                }
                items.Add(row[0], jObject);
            }
            var json = JsonConvert.SerializeObject(items);
            return new(sheetName, json);
        }

        // Need optimization.
        void AddObject(JObject root, string value, HeaderField field)
        {
            root.Add(field.Name, Parse(field.Type, value));
        }

        // Need optimization.
        void AddPrimitiveToArray(JArray root, string type, string value)
        {
            root.Add(Parse(type, value));
        }

        JToken Parse(string type, string value)
        {
            if (type == "int")
            {
                return int.Parse(value);
            }
            else if (type == "float")
            {
                return float.Parse(value);
            }
            else if (type == "bool")
            {
                return bool.Parse(value);
            }
            else
            {
                return value;
            }
        }
    }
}