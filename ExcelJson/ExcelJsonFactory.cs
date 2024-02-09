using System.Text;
using ExcelDataReader;
using ExcelJson.Core;
using Newtonsoft.Json;

namespace ExcelJson
{
    public static class ExcelJsonFactory
    {
        static int m_IsRegisterCodePage1252 = 0;

        public static ExcelJsonReader CreateJsonReader(Stream stream, ExcelJsonOptions excelJsonOptions)
        {
            MakeSureRegisterCP1252();
            var excelDataReader = ExcelReaderFactory.CreateReader(stream);
            return new ExcelJsonReader(excelDataReader, excelJsonOptions);
        }

        public static ExcelJsonSerializer CreateSerializer(JsonSerializerSettings jsonSettings)
        {
            return new ExcelJsonSerializer(jsonSettings);
        }

        static void MakeSureRegisterCP1252()
        {
            if (Interlocked.Exchange(ref m_IsRegisterCodePage1252, 1) == 0)
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