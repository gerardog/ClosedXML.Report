using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;

namespace ClosedXML.Report
{
    internal static class JsonVariableParser
    {
        public static object ParseJson(Stream jsonStream, JsonVariableOptions options)
        {
            using (var document = JsonDocument.Parse(jsonStream))
            {
                var selected = SelectPath(document.RootElement, options?.RootPath).Clone();
                return ConvertElement(selected);
            }
        }

        public static object ParseJson(TextReader jsonReader, JsonVariableOptions options)
        {
            var reader = CreateReader(jsonReader);
            using (var document = JsonDocument.ParseValue(ref reader))
            {
                var selected = SelectPath(document.RootElement, options?.RootPath).Clone();
                return ConvertElement(selected);
            }
        }

        public static IEnumerable<object> ParseJsonLines(TextReader reader, JsonLinesVariableOptions options, bool disposeReader = false)
        {
            try
            {
                int lineNumber = 0;
                var skipEmptyLines = options == null || options.SkipEmptyLines;
                string line;
                while ((line = reader.ReadLine()) != null)
                {
                    lineNumber++;
                    if (skipEmptyLines && string.IsNullOrWhiteSpace(line))
                        continue;

                    JsonDocument doc;
                    try
                    {
                        doc = JsonDocument.Parse(line);
                    }
                    catch (JsonException ex)
                    {
                        throw new JsonException($"Invalid JSONL at line {lineNumber}.", ex);
                    }

                    using (doc)
                    {
                        yield return ConvertElement(doc.RootElement.Clone());
                    }
                }
            }
            finally
            {
                if (disposeReader)
                    reader.Dispose();
            }
        }

        public static object ConvertNode(JsonNode value)
        {
            if (value == null)
                return null;

            var element = JsonSerializer.SerializeToElement(value);
            return ConvertElement(element);
        }

        public static object ConvertElement(JsonElement value)
        {
            switch (value.ValueKind)
            {
                case JsonValueKind.Object:
                    IDictionary<string, object> dictionary = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
                    foreach (var property in value.EnumerateObject())
                    {
                        dictionary[property.Name] = ConvertElement(property.Value);
                    }
                    return dictionary;

                case JsonValueKind.Array:
                    return value.EnumerateArray().Select(item => ConvertElement(item.Clone()));

                case JsonValueKind.String:
                    return value.GetString();

                case JsonValueKind.Number:
                    if (value.TryGetInt64(out var longValue))
                        return longValue;
                    if (value.TryGetDecimal(out var decimalValue))
                        return decimalValue;
                    return value.GetDouble();

                case JsonValueKind.True:
                    return true;

                case JsonValueKind.False:
                    return false;

                case JsonValueKind.Null:
                case JsonValueKind.Undefined:
                    return null;

                default:
                    return value.GetRawText();
            }
        }

        private static JsonElement SelectPath(JsonElement root, string rootPath)
        {
            if (string.IsNullOrWhiteSpace(rootPath))
                return root;

            var current = root;
            var parts = rootPath.Split(new[] { '.' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var part in parts)
            {
                if (current.ValueKind != JsonValueKind.Object || !current.TryGetProperty(part, out var next))
                    throw new ArgumentException(
                        string.Format(CultureInfo.InvariantCulture, "JSON root path '{0}' was not found.", rootPath),
                        nameof(rootPath));

                current = next;
            }

            return current;
        }

        private static Utf8JsonReader CreateReader(TextReader jsonReader)
        {
            var data = jsonReader.ReadToEnd();
            var utf8 = Encoding.UTF8.GetBytes(data);
            return new Utf8JsonReader(utf8, true, default);
        }
    }
}
