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
            using (var document = JsonDocument.Parse(jsonStream, GetDocumentOptions(options)))
            {
                var selected = SelectPath(document.RootElement, options?.RootPath).Clone();
                return ConvertElement(selected);
            }
        }

        public static object ParseJson(TextReader jsonReader, JsonVariableOptions options)
        {
            var data = Encoding.UTF8.GetBytes(jsonReader.ReadToEnd());
            using (var document = JsonDocument.Parse(data, GetDocumentOptions(options)))
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
                        doc = JsonDocument.Parse(line, GetDocumentOptions(options));
                    }
                    catch (JsonException ex)
                    {
                        throw new JsonException($"Invalid JSONL at line {lineNumber}: {ex.Message}", ex);
                    }

                    using (doc)
                    {
                        yield return ConvertElement(SelectPath(doc.RootElement, options?.RootPath).Clone());
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

        private static JsonDocumentOptions GetDocumentOptions(JsonVariableOptions options)
        {
            var serializerOptions = options?.SerializerOptions;
            return new JsonDocumentOptions
            {
                AllowTrailingCommas = serializerOptions?.AllowTrailingCommas ?? false,
                CommentHandling = serializerOptions?.ReadCommentHandling ?? JsonCommentHandling.Disallow
            };
        }
    }
}
