using Windows.Data.Json;

namespace GraphMobileSync.Services
{
    internal static class JsonValueReader
    {
        public static string GetString(JsonObject jsonObject, string name, string defaultValue)
        {
            IJsonValue value;

            if (!TryGetValue(jsonObject, name, JsonValueType.String, out value))
            {
                return defaultValue;
            }

            string result = value.GetString();
            return string.IsNullOrEmpty(result) ? defaultValue : result;
        }

        public static JsonArray GetArray(JsonObject jsonObject, string name)
        {
            IJsonValue value;
            return TryGetValue(jsonObject, name, JsonValueType.Array, out value)
                ? value.GetArray()
                : new JsonArray();
        }

        public static JsonObject GetObject(JsonObject jsonObject, string name)
        {
            IJsonValue value;
            return TryGetValue(jsonObject, name, JsonValueType.Object, out value)
                ? value.GetObject()
                : null;
        }

        public static bool GetBoolean(JsonObject jsonObject, string name, bool defaultValue)
        {
            IJsonValue value;
            return TryGetValue(jsonObject, name, JsonValueType.Boolean, out value)
                ? value.GetBoolean()
                : defaultValue;
        }

        private static bool TryGetValue(
            JsonObject jsonObject,
            string name,
            JsonValueType expectedType,
            out IJsonValue value)
        {
            value = null;

            if (jsonObject == null || !jsonObject.ContainsKey(name))
            {
                return false;
            }

            IJsonValue candidate = jsonObject.GetNamedValue(name);

            if (candidate == null || candidate.ValueType != expectedType)
            {
                return false;
            }

            value = candidate;
            return true;
        }
    }
}
