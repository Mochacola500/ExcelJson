using ExcelDataReader;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json;

namespace ExcelJson
{
    public class ExcelJsonParser : ExcelJsonReader
    {
        readonly ExcelJsonTokenizer m_Tokenizer;

        public ExcelJsonParser() : this(new())
        {

        }

        public ExcelJsonParser(ExcelJsonOptions options) : base(options)
        {
            m_Tokenizer = new();
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
            var tokenizers = ReadTokenizers(headerRow).ToArray();
            var rows = ReadRows(reader, definitions.Count);

            var items = new Dictionary<string, JObject>();
            var fields = headerRow.Fields;
            foreach (var row in rows)
            {
                var jObject = new JObject();
                for (int i = 0; i < fields.Length; ++i)
                {
                    var field = fields[i];
                    var tokenizer = tokenizers[i];
                    if (field.Length == 1)
                    {
                        var token = tokenizer.Invoke(row[i]);
                        jObject.Add(field.Name, token);
                    }
                    else
                    {
                        var jArray = new JArray();
                        for (int j = i; j < field.Length; ++j)
                        {
                            var token = tokenizer.Invoke(row[j]);
                            jArray.Add(token);
                        }
                        jObject.Add(field.Name, jArray);
                    }
                }
                items.Add(row[0], jObject);
            }
            var json = JsonConvert.SerializeObject(items);
            return new(sheetName, json);
        }

        IEnumerable<ExcelJsonTokenizer.ToJToken> ReadTokenizers(HeaderRow headerRow)
        {
            foreach (var field in headerRow.Fields)
            {
                yield return m_Tokenizer.FindTokenizeFunction(field.Type);
            }
        }
    }
}