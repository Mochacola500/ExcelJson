
namespace ExcelJson
{
    public class ArrayPlaceholderException : Exception
    {
        public string SheetName { get; set; }
        public int Index { get; set; }

        public ArrayPlaceholderException(string sheetName, int index) : base($"Definition is required befor the array token.\nName:{sheetName}\nIndex:{index}")
        {
            SheetName = sheetName;
            Index = index;
        }
    }
}