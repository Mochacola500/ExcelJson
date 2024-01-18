using ExcelDataReader;

namespace ExcelJson
{
    public interface IExcelJsonParser
    {
        IEnumerable<ExcelJsonSheet> ParseExcel(Stream stream);
        IEnumerable<ExcelJsonSheet> ParseExcel(Stream stream, ExcelReaderConfiguration excelReaderConfig);
    }
}