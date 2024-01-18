using ExcelDataReader;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json;

namespace ExcelJson
{
    public class ExcelJsonParser : ExcelJsonReader
    {
        readonly ExcelJsonTokenizer m_Tokenizer;
        readonly JsonSerializerSettings m_Settings;

        public ExcelJsonParser() : this(new(), new()) { }
        public ExcelJsonParser(ExcelJsonOptions options) : this(options, new()) { }
        public ExcelJsonParser(JsonSerializerSettings settings) : this(new(), settings) { }
        public ExcelJsonParser(ExcelJsonOptions options, JsonSerializerSettings settings) : base(options)
        {
            m_Tokenizer = new();
            m_Settings = settings;
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
            var tokenizers = ReadTokenizers(headerRow);
            var rows = ReadRows(reader, definitions.Count);

            var items = new Dictionary<JToken, JObject>();
            var fields = headerRow.Fields;
            foreach (var row in rows)
            {
                var jObject = new JObject();
                for (int i = 0; i < fields.Length; ++i)
                {
                    var field = fields[i];
                    var fieldName = field.Name;
                    var tokenizer = tokenizers[i];
                    if (field.Length == 1)
                    {
                        var token = tokenizer.Invoke(row[i]);
                        jObject.Add(fieldName, token);
                    }
                    else
                    {
                        var jArray = new JArray();
                        for (int j = i; j < field.Length; ++j)
                        {
                            var token = tokenizer.Invoke(row[j]);
                            jArray.Add(token);
                        }
                        jObject.Add(fieldName, jArray);
                    }
                }
                var key = tokenizers[0].Invoke(row[0]);
                items.Add(key, jObject);
            }
            var json = JsonConvert.SerializeObject(items, m_Settings);
            return new(sheetName, json);
        }

        ToJTokenFunction[] ReadTokenizers(HeaderRow headerRow)
        {
            var fields = headerRow.Fields;
            var functions = new ToJTokenFunction[fields.Length];
            for (int i = 0; i <  functions.Length; ++i)
            {
                functions[i] = m_Tokenizer.GetTokenizeFunction(fields[i].Type);
            }
            return functions;
        }
    }
}