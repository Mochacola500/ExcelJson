
namespace ExcelJson
{
    internal class ExcelCodeParser : ExcelJsonReader
    {
        readonly ExcelCodeOptions m_CodeOptions;
        readonly ExcelCodeEmitter m_Emitter;

        internal ExcelCodeParser(ExcelCodeOptions codeOptions, ExcelJsonOptions jsonOptions) : base(jsonOptions)
        {
            m_CodeOptions = codeOptions;
            m_Emitter = new();
        }
    }
}