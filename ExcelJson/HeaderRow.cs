namespace ExcelJson
{
    public class HeaderRow
    {
        public string Name { get; }
        public HeaderField Key { get; }
        public HeaderField[] Fields { get; }

        public HeaderRow(string name, HeaderField key, HeaderField[] fields)
        {
            Name = name;
            Key = key;
            Fields = fields;
        }
    }
}