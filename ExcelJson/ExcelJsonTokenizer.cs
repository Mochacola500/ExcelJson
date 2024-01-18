using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace ExcelJson
{
    public delegate JToken JTokenFunction(string value);

    public class ExcelJsonTokenizer
    {
        static readonly Dictionary<string, JTokenFunction> m_DefaultJTokens = new()
        {
            { "short", ToInt16 },
            { "int", ToInt32 },
            { "long", ToInt64 },
            { "float", ToSingle },
            { "double", ToDouble },
            { "decimal", ToDecimal },
            { "ushort", ToUInt16 },
            { "uint", ToUInt32 },
            { "ulong", ToUInt64 },
            { "bool", ToBoolean },
            { "string", ToString },
            { "byte", ToByte },
        };

        readonly Dictionary<string, JTokenFunction> m_CustomJTokens;
        readonly JsonSerializerSettings m_Settings;

        public ExcelJsonTokenizer() : this(new()) { }
        public ExcelJsonTokenizer(JsonSerializerSettings settings)
        {
            m_Settings = settings;
            m_CustomJTokens = new();
        }

        public void AddJTokenFunction(string key, JTokenFunction value)
        {
            m_CustomJTokens.Add(key, value);
        }

        public JTokenFunction GetTokenizeFunction(string type)
        {
            if (m_CustomJTokens.TryGetValue(type, out var function))
            {
                return function;
            }
            if (m_DefaultJTokens.TryGetValue(type, out function))
            {
                return function;
            }
            if (type == "DateTime")
            {
                return ToDateTime;
            }
            return ToString;
        }

        static JToken ToInt16(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return (short)0;
            }
            return short.Parse(value);
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

        static JToken ToUInt16(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return (ushort)0;
            }
            return ushort.Parse(value);
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

        static JToken ToByte(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return ' ';
            }
            return Convert.ToByte(value);
        }

        JToken ToDateTime(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                throw new ArgumentNullException("Date time is null or empty.");
            }
            return DateTime.Parse(value, m_Settings.Culture);
        }
    }
}