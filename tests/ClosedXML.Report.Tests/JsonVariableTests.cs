using System;
using System.Collections.Generic;
using System.Dynamic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using ClosedXML.Excel;
using FluentAssertions;
using Xunit;

namespace ClosedXML.Report.Tests
{
    public class JsonVariableTests
    {
        [Fact]
        public void AddJsonVariable_should_bind_root_object_properties()
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.AddWorksheet("Sheet1");
                worksheet.Cell("A1").Value = "{{person.Name}}";
                worksheet.Cell("B1").Value = "{{person.Age}}";

                var json = "{\"Name\":\"Alice\",\"Age\":30}";
                using (var stream = new MemoryStream(Encoding.UTF8.GetBytes(json)))
                using (var template = new XLTemplate(workbook))
                {
                    template.AddJsonVariable("person", stream);
                    template.Generate();

                    worksheet.Cell("A1").GetString().Should().Be("Alice");
                    worksheet.Cell("B1").GetValue<int>().Should().Be(30);
                }
            }
        }

        [Fact]
        public void AddJsonVariable_should_bind_root_array_to_named_range()
        {
            using (var template = CreateItemsTemplate())
            {
                var worksheet = template.Workbook.Worksheet(1);
                var json = "[{\"Name\":\"A\",\"Age\":1},{\"Name\":\"B\",\"Age\":2}]";
                using (var stream = new MemoryStream(Encoding.UTF8.GetBytes(json)))
                {
                    template.AddJsonVariable("Persons", stream);
                    template.Generate();
                }

                worksheet.Cell("A2").GetString().Should().Be("A");
                worksheet.Cell("B2").GetValue<int>().Should().Be(1);
                worksheet.Cell("A3").GetString().Should().Be("B");
                worksheet.Cell("B3").GetValue<int>().Should().Be(2);
            }
        }

        [Fact]
        public void AddJsonVariable_should_support_root_path_selection()
        {
            using (var template = CreateItemsTemplate())
            {
                var worksheet = template.Workbook.Worksheet(1);
                var json = "{\"data\":{\"items\":[{\"Name\":\"A\",\"Age\":1},{\"Name\":\"B\",\"Age\":2}]}}";
                using (var stream = new MemoryStream(Encoding.UTF8.GetBytes(json)))
                {
                    template.AddJsonVariable("Persons", stream, new JsonVariableOptions { RootPath = "data.items" });
                    template.Generate();
                }

                worksheet.Cell("A2").GetString().Should().Be("A");
                worksheet.Cell("A3").GetString().Should().Be("B");
            }
        }

        [Fact]
        public void AddJsonLinesVariable_should_bind_jsonl_stream_to_named_range()
        {
            using (var template = CreateItemsTemplate())
            {
                var worksheet = template.Workbook.Worksheet(1);
                var jsonl = "{\"Name\":\"A\",\"Age\":1}\n\n{\"Name\":\"B\",\"Age\":2}\n";
                using (var stream = new MemoryStream(Encoding.UTF8.GetBytes(jsonl)))
                {
                    template.AddJsonLinesVariable("Persons", stream);
                    template.Generate();
                }

                worksheet.Cell("A2").GetString().Should().Be("A");
                worksheet.Cell("B2").GetValue<int>().Should().Be(1);
                worksheet.Cell("A3").GetString().Should().Be("B");
                worksheet.Cell("B3").GetValue<int>().Should().Be(2);
            }
        }

        [Fact]
        public void AddJsonLinesVariable_should_not_read_all_lines_eagerly()
        {
            var reader = new SequentialJsonLinesReader(100000);
            var values = JsonVariableParser.ParseJsonLines(reader, new JsonLinesVariableOptions());

            using (var enumerator = values.GetEnumerator())
            {
                enumerator.MoveNext().Should().BeTrue();
                reader.ReadLineCalls.Should().Be(1);

                enumerator.MoveNext().Should().BeTrue();
                reader.ReadLineCalls.Should().Be(2);
            }
        }

        [Fact]
        public void AddJsonLinesVariable_should_include_line_number_for_malformed_line()
        {
            var reader = new StringReader("{\"Name\":\"A\"}\n{bad json}\n{\"Name\":\"B\"}");
            var values = JsonVariableParser.ParseJsonLines(reader, new JsonLinesVariableOptions()).ToList;

            Action act = () => values();

            act.Should()
                .Throw<JsonException>()
                .WithMessage("*line 2*");
        }

        [Fact]
        public void AddJsonLinesVariable_should_support_null_and_mixed_values()
        {
            var reader = new StringReader("null\n{\"Name\":\"A\"}\n42\n");
            var values = JsonVariableParser.ParseJsonLines(reader, new JsonLinesVariableOptions()).ToList();

            values[0].Should().BeNull();
            ((IDictionary<string, object>) values[1]).Should().ContainKey("Name");
            ((IDictionary<string, object>) values[1])["Name"].Should().Be("A");
            values[2].Should().Be(42L);
        }

        [Fact]
        public void AddVariable_should_support_json_element_and_json_node()
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.AddWorksheet("Sheet1");
                worksheet.Cell("A1").Value = "{{Name}}";
                worksheet.Cell("B1").Value = "{{Node.Name}}";

                using (var template = new XLTemplate(workbook))
                using (var doc = JsonDocument.Parse("{\"Name\":\"Element\"}"))
                {
                    template.AddVariable(doc.RootElement);
                    template.AddVariable("Node", JsonNode.Parse("{\"Name\":\"Node\"}"));
                    template.Generate();

                    worksheet.Cell("A1").GetString().Should().Be("Element");
                    worksheet.Cell("B1").GetString().Should().Be("Node");
                }
            }
        }

        [Fact]
        public void AddJsonVariable_should_keep_stream_open_by_default()
        {
            using (var workbook = new XLWorkbook())
            {
                workbook.AddWorksheet("Sheet1").Cell("A1").Value = "{{person.Name}}";
                var json = "{\"Name\":\"Alice\"}";
                using (var stream = new MemoryStream(Encoding.UTF8.GetBytes(json)))
                using (var template = new XLTemplate(workbook))
                {
                    template.AddJsonVariable("person", stream);
                    stream.CanRead.Should().BeTrue();
                }
            }
        }

        private static XLTemplate CreateItemsTemplate()
        {
            var workbook = new XLWorkbook();
            var worksheet = workbook.AddWorksheet("Sheet1");
            worksheet.Cell("A2").Value = "{{item.Name}}";
            worksheet.Cell("B2").Value = "{{item.Age}}";
            worksheet.Range("A2:B3").AddToNamed("Persons");
            return new XLTemplate(workbook);
        }

        private class SequentialJsonLinesReader : TextReader
        {
            private readonly int _totalLines;
            private int _lineIndex;

            public SequentialJsonLinesReader(int totalLines)
            {
                _totalLines = totalLines;
            }

            public int ReadLineCalls { get; private set; }

            public override string ReadLine()
            {
                ReadLineCalls++;
                if (_lineIndex >= _totalLines)
                    return null;

                var value = _lineIndex;
                _lineIndex++;
                return string.Format("{{\"Name\":\"N{0}\",\"Age\":{0}}}", value);
            }
        }
    }
}
