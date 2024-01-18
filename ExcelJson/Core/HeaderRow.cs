namespace ExcelJson
{
    internal class HeaderRow
    {
        public string Name { get; }
        public HeaderField[] Fields { get; }

        public HeaderRow(string name, HeaderField[] fields)
        {
            Name = name;
            Fields = fields;
        }
    }
}