using Newtonsoft.Json.Linq;

namespace ExcelJson.Core
{
    delegate JToken JTokenFunction(string value);

    internal static class JTokenUtils
    {
        static readonly Dictionary<string, JTokenFunction> m_PrimitiveTypes = new()
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

        public static bool TryGetJToken(string type, out JTokenFunction function)
        {
            if (m_PrimitiveTypes.TryGetValue(type, out var successFunction))
            {
                function = successFunction;
                return true;
            }
            else
            {
                function = ToString;
                return false;
            }
        }

        public static JToken ToInt16(string value)
        {
            return Convert.ToInt16(value);
        }

        public static JToken ToInt32(string value)
        {
            return Convert.ToInt32(value);
        }

        public static JToken ToInt64(string value)
        {
            return Convert.ToInt64(value);
        }

        public static JToken ToSingle(string value)
        {
            return Convert.ToSingle(value);
        }

        public static JToken ToDouble(string value)
        {
            return Convert.ToDouble(value);
        }

        public static JToken ToDecimal(string value)
        {
            return Convert.ToDecimal(value);
        }

        public static JToken ToUInt16(string value)
        {
            return Convert.ToUInt16(value);
        }

        public static JToken ToUInt32(string value)
        {
            return Convert.ToUInt32(value);
        }

        public static JToken ToUInt64(string value)
        {
            return Convert.ToUInt64(value);
        }

        public static JToken ToBoolean(string value)
        {
            return Convert.ToBoolean(value);
        }

        public static JToken ToByte(string value)
        {
            return Convert.ToByte(value);
        }

        public static JToken ToString(string value)
        {
            return value;
        }
    }
}