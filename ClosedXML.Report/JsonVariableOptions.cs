using System.Text;
using System.Text.Json;

namespace ClosedXML.Report
{
    public class JsonVariableOptions
    {
        public Encoding Encoding { get; set; }

        public string RootPath { get; set; }

        public JsonSerializerOptions SerializerOptions { get; set; }

        public bool LeaveOpen { get; set; } = true;
    }

    public class JsonLinesVariableOptions : JsonVariableOptions
    {
        public bool SkipEmptyLines { get; set; } = true;
    }
}
