using ExcelDataReader;

namespace ExcelJson
{
    public class ExcelJsonCodeGenerator : ExcelJsonReader
    {

        public ExcelJsonCodeGenerator() : this(new())
        {

        }

        public ExcelJsonCodeGenerator(ExcelJsonOptions options) : base(options)
        {

        }

        public void WriteCodes(Stream stream, string directory)
        {
            using var reader = ExcelReaderFactory.CreateReader(stream);
            reader.Reset();
            var headerRows = new List<HeaderRow>();
            do
            {
                var sheetName = reader.Name;
                if (sheetName.StartsWith(m_Options.FilterSheetToken))
                {
                    continue;
                }
                var definitions = ReadDefinitions(reader);
                var headerRow = ReadHeaderRow(sheetName, definitions);
                headerRows.Add(headerRow);
            } while (reader.NextResult());



        }
    }
}