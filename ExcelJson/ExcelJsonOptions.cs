
namespace ExcelJson
{
    public class ExcelJsonOptions
    {
        public string ArrayPlaceholder { get; set; } = "//";
        public char TypeToken { get; set; } = ':';
        public char FilterSheetToken { get; set; } = '#';
        public char FilterColumnToken { get; set; } = '#';
    }
}