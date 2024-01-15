using ExcelDataReader;
using ExcelJson.DataModel;

namespace ExcelJson
{
    public class ExcelJsonCodeGenerator : ExcelJsonReader
    {
        public ExcelJsonCodeGenerator(ExcelJsonOptions options) : base(options)
        {

        }

        public static ExcelJsonCodeGenerator Create(ExcelJsonOptions? options = null)
        {
            options ??= new();
            return new ExcelJsonCodeGenerator(options);
        }

        //public IEnumerable<ExcelJsonCode> ToCodes(Stream stream)
        //{
        //    using var reader = ExcelReaderFactory.CreateReader(stream);
        //    reader.Reset();
        //    var headerRows = new List<HeaderRow>();
        //    do
        //    {
        //        var sheetName = reader.Name;
        //        if (sheetName.StartsWith(m_Options.FilterSheetToken))
        //        {
        //            continue;
        //        }
        //        var definitions = ReadDefinitions(reader);
        //        var headerRow = ReadHeaderRow(sheetName, definitions);
        //        headerRows.Add(headerRow);
        //    } while (reader.NextResult());
        //
        //    
        //
        //}
    }
}