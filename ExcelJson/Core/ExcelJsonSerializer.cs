using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace ExcelJson.Core
{
    public class ExcelJsonSerializer
    {
        readonly JsonSerializerSettings m_JsonSettings;
        readonly Dictionary<string, JTokenFunction> m_CustomTypes;

        public ExcelJsonSerializer(JsonSerializerSettings jsonSettings)
        {
            m_JsonSettings = jsonSettings;
            m_CustomTypes = new()
            {
                { "DateTime", ToDateTime }
            };
        }

        public string ToJson(ExcelJsonSheet jsonSheet)
        {
            var fields = jsonSheet.Header.Fields;
            var values = new Dictionary<JToken, JObject>();

            foreach (var row in jsonSheet.Rows)
            {
                var jObject = new JObject();
                for (int i = 0; i < fields.Length; ++i)
                {
                    var field = fields[i];
                    var name = field.Name;
                    var type = field.Type;
                    if (field.ArrayLength == 1)
                    {
                        var token = ToJToken(type, row[i]);
                        jObject.Add(name, token);
                    }
                    else
                    {
                        var jArray = new JArray();
                        for (int j = i; j < field.ArrayLength; ++j)
                        {
                            var token = ToJToken(type, row[j]);
                            jArray.Add(token);
                        }
                        jObject.Add(name, jArray);
                    }
                }
                var key = ToJToken(fields[0].Type, row[0]);
                values.Add(key, jObject);
            }
            return JsonConvert.SerializeObject(values, m_JsonSettings);
        }

        JToken ToJToken(string type, string value)
        {
            return GetJTokenFunction(type).Invoke(value);
        }

        JTokenFunction GetJTokenFunction(string type)
        {
            if (JTokenUtils.TryGetJToken(type, out var function))
            {
                return function;
            }
            if (m_CustomTypes.TryGetValue(type, out function))
            {
                return function;
            }
            return JTokenUtils.ToString;
        }

        JToken ToDateTime(string value)
        {
            return Convert.ToDateTime(value, m_JsonSettings.Culture);
        }
    }
}