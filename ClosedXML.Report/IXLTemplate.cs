using System;
using System.IO;
using System.Text.Json;
using System.Text.Json.Nodes;
using ClosedXML.Excel;

namespace ClosedXML.Report
{
    public interface IXLTemplate : IDisposable
    {
        public IXLWorkbook Workbook { get; }

        public XLGenerateResult Generate();

        public void AddVariable(object value);

        public void AddVariable(string alias, object value);

        public void AddVariable(JsonElement value);

        public void AddVariable(string alias, JsonElement value);

        public void AddVariable(JsonNode value);

        public void AddVariable(string alias, JsonNode value);

        public void AddJsonVariable(string alias, Stream jsonStream, JsonVariableOptions options = null);

        public void AddJsonVariable(string alias, TextReader jsonReader, JsonVariableOptions options = null);

        public void AddJsonLinesVariable(string alias, Stream jsonLinesStream, JsonLinesVariableOptions options = null);

        public void AddJsonLinesVariable(string alias, TextReader jsonLinesReader, JsonLinesVariableOptions options = null);

        public void SaveAs(string file);

        public void SaveAs(string file, SaveOptions options);

        public void SaveAs(string file, bool validate, bool evaluateFormulae = false);

        public void SaveAs(Stream stream);

        public void SaveAs(Stream stream, SaveOptions options);

        public void SaveAs(Stream stream, bool validate, bool evaluateFormulae = false);
    }
}
