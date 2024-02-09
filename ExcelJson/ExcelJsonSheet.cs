
namespace ExcelJson
{
    public class ExcelJsonSheet
    {
        public string Name { get; }
        public ExcelHeader Header { get; }
        public IEnumerable<string[]> Rows { get; }

        public ExcelJsonSheet(string name, ExcelHeader header, IEnumerable<string[]> rows)
        {
            Name = name;
            Header = header;
            Rows = rows;
        }
    }

    public class ExcelHeader
    {
        public string Name { get; }
        public HeaderField[] Fields { get; }
        public int TotalCount { get; }

        public ExcelHeader(string name, HeaderField[] fields, int totalCount)
        {
            Name = name;
            Fields = fields;
            TotalCount = totalCount;
        }
    }

    public class HeaderField
    {
        public string Type { get; }
        public string Name { get; }
        public int ArrayLength { get; set; }

        public HeaderField(string type, string name, int length)
        {
            Type = type;
            Name = name;
            ArrayLength = length;
        }
    }
}