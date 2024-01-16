
namespace ExcelJson
{
    public class TypeTokenException : Exception
    {
        public string SheetName {  get; set; }
        public string Definition { get; set; }
        public char TypeToken {  get; set; }

        public TypeTokenException(string sheetName, string definition, char typeToken) : base($"Missing type token.\nTableName:{sheetName}\nDefinition:{definition}\nToken:{typeToken}")
        {
            SheetName = sheetName;
            Definition = definition;
            TypeToken = typeToken;
        }
    }
}