namespace ExcelJson.DataModel
{
    public class ExcelJsonField
    {
        public string Name { get; }
        public string Json { get; }

        public ExcelJsonField(string name, string json)
        {
            Name = name;
            Json = json;
        }
    }
}