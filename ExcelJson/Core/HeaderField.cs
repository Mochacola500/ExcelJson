namespace ExcelJson
{
    internal class HeaderField
    {
        public string Type { get; }
        public string Name { get; }
        public int Length { get; set; }

        public HeaderField(string type, string name, int length)
        {
            Type = type;
            Name = name;
            Length = length;
        }
    }
}
