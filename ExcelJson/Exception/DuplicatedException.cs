
namespace ExcelJson
{
    public class DuplicatedException : Exception
    {
        public string SheetName {  get; set; }
        public string Definition {  get; set; }

        public DuplicatedException(string sheetName, string definition) : base($"Duplicated definition name.\nTableName:{sheetName}\nDefinition:{definition}")
        {
            SheetName = sheetName;
            Definition = definition;
        }
    }
}