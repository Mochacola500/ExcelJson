using System.Text;
using Newtonsoft.Json;

namespace ExcelJson
{
    public static class ExcelJsonFactory
    {
        static int m_IsRegisterCodePage1252 = 0;

        public static IExcelJsonParser CreateJsonParser(ExcelJsonOptions jsonOptions, JsonSerializerSettings jsonSettings)
        {
            MakeSureRegisterCP1252();
            return new ExcelJsonParser(jsonOptions, jsonSettings);
        }

        static void MakeSureRegisterCP1252()
        {
            if (Interlocked.Exchange(ref m_IsRegisterCodePage1252, 1) != 0)
            {
                var cp1252 = Encoding.GetEncodings()
                    .Where(x => x.CodePage == 1252)
                    .FirstOrDefault();

                if (cp1252 == null)
                {
                    Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
                }
            }
        }
    }
}