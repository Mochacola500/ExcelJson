
namespace ExcelJson.DataModel
{
    public class ExcelJsonCode
    {
        public string Name { get; }
        public string Code { get; }

        public ExcelJsonCode(string name, string code)
        {
            Name = name;
            Code = code;
        }
    }
}