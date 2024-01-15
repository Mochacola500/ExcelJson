using ExcelDataReader;
using Newtonsoft.Json.Linq;
using ExcelJson.DataModel;
using Newtonsoft.Json;

namespace ExcelJson
{
    public class ExcelJsonParser : ExcelJsonReader
    {
        ExcelJsonParser(ExcelJsonOptions options) : base(options)
        {

        }

        public static ExcelJsonParser Create(ExcelJsonOptions? options = null)
        {
            options ??= new();
            return new ExcelJsonParser(options);
        }

        public IEnumerable<ExcelJsonField> ToExcelJsonSet(Stream stream)
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
                yield return ToExcelJsonResult(reader);
            } while (reader.NextResult());
        }

        public ExcelJsonField ToExcelJsonResult(IExcelDataReader reader)
        {
            var sheetName = reader.Name;
            var definitions = ReadDefinitions(reader);
            var headerRow = ReadHeaderRow(sheetName, definitions);
            var rows = ReadRows(reader, definitions.Count);

            var items = new Dictionary<string, JObject>();
            var key = headerRow.Key;
            var fields = headerRow.Fields;
            foreach (var row in rows)
            {
                var jObject = new JObject();
                for (int i = 0; i < fields.Length; ++i)
                {
                    var field = fields[i];
                    if (field.Length == 1)
                    {
                        jObject.Add(field.Name, row[i]);
                    }
                    else
                    {
                        var jArray = new JArray();
                        for (int j = i; j < field.Length; ++j)
                        {
                            jArray.Add(row[j]);
                        }
                        jObject.Add(field.Name, jArray);
                    }
                }
                items.Add(key.Name, jObject);
            }
            var json = JsonConvert.SerializeObject(items);
            return new(sheetName, json);
        }
    }
}