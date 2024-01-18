using ExcelDataReader;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json;

namespace ExcelJson
{
    public class ExcelJsonParser : ExcelJsonReader
    {
        readonly ExcelJsonTokenizer m_Tokenizer;
        readonly JsonSerializerSettings m_JsonSettings;

        public ExcelJsonTokenizer Tokenizer
        {
            get
            {
                return m_Tokenizer;
            }
        }

        public ExcelJsonParser() : this(new(), new()) { }
        public ExcelJsonParser(ExcelJsonOptions jsonOptions) : this(jsonOptions, new()) { }
        public ExcelJsonParser(JsonSerializerSettings jsonSettings) : this(new(), jsonSettings) { }
        public ExcelJsonParser(ExcelJsonOptions jsonOptions, JsonSerializerSettings jsonSettings) : base(jsonOptions)
        {
            m_Tokenizer = new(jsonSettings);
            m_JsonSettings = jsonSettings;
        }

        public IEnumerable<ExcelJsonSheet> ParseExcel(Stream stream)
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
            var jTokenFunctions = ReadJTokenFunctions(headerRow);

            var items = new Dictionary<JToken, JObject>();
            var fields = headerRow.Fields;
            foreach (var row in rows)
            {
                var jObject = new JObject();
                for (int i = 0; i < fields.Length; ++i)
                {
                    var field = fields[i];
                    var fieldName = field.Name;
                    var jTokenFunction = jTokenFunctions[i];
                    if (field.Length == 1)
                    {
                        var token = jTokenFunction.Invoke(row[i]);
                        jObject.Add(fieldName, token);
                    }
                    else
                    {
                        var jArray = new JArray();
                        for (int j = i; j < field.Length; ++j)
                        {
                            var token = jTokenFunction.Invoke(row[j]);
                            jArray.Add(token);
                        }
                        jObject.Add(fieldName, jArray);
                    }
                }
                var key = jTokenFunctions[0].Invoke(row[0]);
                items.Add(key, jObject);
            }
            var json = JsonConvert.SerializeObject(items, m_JsonSettings);
            return new(sheetName, json);
        }

        JTokenFunction[] ReadJTokenFunctions(HeaderRow headerRow)
        {
            var fields = headerRow.Fields;
            var functions = new JTokenFunction[fields.Length];
            for (int i = 0; i <  functions.Length; ++i)
            {
                functions[i] = m_Tokenizer.GetTokenizeFunction(fields[i].Type);
            }
            return functions;
        }
    }
}