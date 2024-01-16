namespace ExcelJson
{
    public class ExcelJsonSheet
    {
        public string Name { get; }
        public string Json { get; }

        public ExcelJsonSheet(string name, string json)
        {
            Name = name;
            Json = json;
        }
    }
}