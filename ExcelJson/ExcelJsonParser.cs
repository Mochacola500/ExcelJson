using ExcelDataReader;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json;
using Microsoft.VisualBasic.FileIO;

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
                            AddPrimitiveToArray(jArray, row[j], field.Type);
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
            if (field.Type == "int")
            {
                root.Add(field.Name, int.Parse(value));
            }
            else
            {
                root.Add(field.Name, value);
            }
        }

        // Need optimization.
        void AddPrimitiveToArray(JArray root,string value, string type)
        {
            if (type == "int")
            {
                root.Add(int.Parse(value));
            }
            else
            {
                root.Add(value);
            }
        }
    }
}