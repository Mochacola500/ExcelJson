using Newtonsoft.Json.Linq;

namespace ExcelJson
{
    public delegate JToken ToJTokenFunction(string value);

    public class ExcelJsonTokenizer
    {
        static readonly Dictionary<string, ToJTokenFunction> m_DefaultTokenizer = new()
        {
            { "int", ToInt32 },
            { "long", ToInt64 },
            { "float", ToSingle },
            { "double", ToDouble },
            { "decimal", ToDecimal },
            { "uint", ToUInt32 },
            { "ulong", ToUInt64 },
            { "string", ToString },
            { "bool", ToBoolean },
            { "DateTime", ToDateTime }
        };

        public ToJTokenFunction GetTokenizeFunction(string type)
        {
            if (m_DefaultTokenizer.TryGetValue(type, out var function))
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

        static JToken ToBoolean(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return false;
            }
            return bool.Parse(value);
        }

        static JToken ToString(string value)
        {
            return value;
        }

        static JToken ToDateTime(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                throw new ArgumentNullException("Date time is null or empty.");
            }
            return DateTime.Parse(value);
        }
    }
}