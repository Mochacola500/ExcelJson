using Newtonsoft.Json.Linq;

namespace ExcelJson
{
    public class ExcelJsonTokenizer
    {
        public delegate JToken ToJToken(string value);

        static readonly Dictionary<string, ToJToken> g_DefaultTokenizer = new()
        {
            { "int", ToInt32 },
            { "long", ToInt64 },
            { "float", ToSingle },
            { "double", ToDouble },
            { "decimal", ToDecimal },
            { "uint", ToUInt32 },
            { "ulong", ToUInt64 },
            { "string", ToString },
        };

        public ToJToken FindTokenizeFunction(string type)
        {
            if (g_DefaultTokenizer.TryGetValue(type, out var function))
            {
                return function;
            }
            return ToString;
        }

        static JToken ToInt32(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return 0;
            }
            return int.Parse(value);
        }

        static JToken ToInt64(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return 0L;
            }
            return long.Parse(value);
        }

        static JToken ToSingle(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return 0f;
            }
            return float.Parse(value);
        }

        static JToken ToDouble(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return 0D;
            }
            return double.Parse(value);
        }

        static JToken ToDecimal(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return 0m;
            }
            return decimal.Parse(value);
        }

        static JToken ToUInt32(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return 0u;
            }
            return uint.Parse(value);
        }

        static JToken ToUInt64(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return 0UL;
            }
            return ulong.Parse(value);
        }

        static JToken ToString(string value)
        {
            return value;
        }
    }
}