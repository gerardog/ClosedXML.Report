using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Diagnostics;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using ClosedXML.Excel;
using FluentAssertions;
using Xunit;
using Xunit.Abstractions;

namespace ClosedXML.Report.Tests
{
    public class JsonVariableTests
    {
        private readonly ITestOutputHelper _output;

        public JsonVariableTests(ITestOutputHelper output)
        {
            _output = output;
        }

        [Fact]
        public void AddJsonVariable_should_bind_root_object_properties()
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.AddWorksheet("Sheet1");
                worksheet.Cell("A1").Value = "{{person[\"Name\"]}}";
                worksheet.Cell("B1").Value = "{{person[\"Age\"]}}";

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
            Action act = () => JsonVariableParser.ParseJsonLines(reader, new JsonLinesVariableOptions()).ToList();

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
                worksheet.Cell("B1").Value = "{{Node[\"Name\"]}}";

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
                workbook.AddWorksheet("Sheet1").Cell("A1").Value = "{{person[\"Name\"]}}";
                var json = "{\"Name\":\"Alice\"}";
                using (var stream = new MemoryStream(Encoding.UTF8.GetBytes(json)))
                using (var template = new XLTemplate(workbook))
                {
                    template.AddJsonVariable("person", stream);
                    stream.CanRead.Should().BeTrue();
                }
            }
        }

        [Theory]
        [InlineData(1)]
        [InlineData(3)]
        [InlineData(10)]
        public void AddJsonVariable_should_process_large_json_payloads_end_to_end(int targetSizeInMegabytes)
        {
            var payload = CreateLargeJsonPayload(targetSizeInMegabytes);
            using (var jsonStream = new MemoryStream(payload.Utf8Bytes, writable: false))
            using (var templateStream = CreateLargeItemsTemplateStream())
            {
                var baseline = MemorySnapshot.Capture();

                using (var template = new XLTemplate(templateStream))
                {
                    template.AddJsonVariable("Persons", jsonStream);
                    var afterAddVariable = MemorySnapshot.Capture();

                    template.Generate();
                    var afterGenerate = MemorySnapshot.Capture();

                    using (var outputStream = new MemoryStream())
                    {
                        template.SaveAs(outputStream);
                        var afterSave = MemorySnapshot.Capture();

                        outputStream.Position = 0;
                        using (var workbook = new XLWorkbook(outputStream))
                        {
                            var worksheet = workbook.Worksheet(1);
                            var lastRow = payload.ItemCount + 1;

                            worksheet.Cell("A2").GetValue<long>().Should().Be(0);
                            worksheet.Cell("B2").GetString().Should().Be("Name0");
                            worksheet.Cell("C2").GetString().Should().Be("Group0");
                            worksheet.Cell("D2").GetValue<long>().Should().Be(20);

                            worksheet.Cell(lastRow, 1).GetValue<long>().Should().Be(payload.ItemCount - 1);
                            worksheet.Cell(lastRow, 2).GetString().Should().Be($"Name{payload.ItemCount - 1}");
                            worksheet.Cell(lastRow, 3).GetString().Should().Be($"Group{(payload.ItemCount - 1) % 10}");
                            worksheet.Cell(lastRow, 4).GetValue<long>().Should().Be(20 + ((payload.ItemCount - 1) % 50));
                        }

                        var peakManagedBytes = new[]
                        {
                            baseline.ManagedBytes,
                            afterAddVariable.ManagedBytes,
                            afterGenerate.ManagedBytes,
                            afterSave.ManagedBytes
                        }.Max();

                        var peakWorkingSetBytes = new[]
                        {
                            baseline.WorkingSetBytes,
                            afterAddVariable.WorkingSetBytes,
                            afterGenerate.WorkingSetBytes,
                            afterSave.WorkingSetBytes
                        }.Max();

                        _output.WriteLine(
                            $"JSON size={payload.Utf8Bytes.Length:N0} bytes (~{targetSizeInMegabytes} MB), items={payload.ItemCount:N0}, " +
                            $"managed baseline={baseline.ManagedBytes:N0}, after add={afterAddVariable.ManagedBytes:N0}, after generate={afterGenerate.ManagedBytes:N0}, after save={afterSave.ManagedBytes:N0}, " +
                            $"peak managed delta={peakManagedBytes - baseline.ManagedBytes:N0}, peak working set delta={peakWorkingSetBytes - baseline.WorkingSetBytes:N0}");
                    }
                }
            }
        }

        private static XLTemplate CreateItemsTemplate()
        {
            var workbook = new XLWorkbook();
            var worksheet = workbook.AddWorksheet("Sheet1");
            worksheet.Cell("A2").Value = "{{item[\"Name\"]}}";
            worksheet.Cell("B2").Value = "{{item[\"Age\"]}}";
            worksheet.Range("A2:B3").AddToNamed("Persons");
            return new XLTemplate(workbook);
        }

        private static MemoryStream CreateLargeItemsTemplateStream()
        {
            var workbook = new XLWorkbook();
            try
            {
                var worksheet = workbook.AddWorksheet("Sheet1");
                worksheet.Cell("A2").Value = "{{item[\"Id\"]}}";
                worksheet.Cell("B2").Value = "{{item[\"Name\"]}}";
                worksheet.Cell("C2").Value = "{{item[\"Category\"]}}";
                worksheet.Cell("D2").Value = "{{item[\"Age\"]}}";
                worksheet.Range("A2:D3").AddToNamed("Persons");

                var stream = new MemoryStream();
                workbook.SaveAs(stream);
                stream.Position = 0;
                return stream;
            }
            finally
            {
                workbook.Dispose();
            }
        }

        private static LargeJsonPayload CreateLargeJsonPayload(int targetSizeInMegabytes)
        {
            const int paddingLength = 2048;
            var targetBytes = targetSizeInMegabytes * 1024 * 1024;
            var padding = new string('P', paddingLength);
            var builder = new StringBuilder(targetBytes + 1024);
            builder.Append('[');

            var itemCount = 0;
            while (builder.Length < targetBytes)
            {
                if (itemCount > 0)
                    builder.Append(',');

                builder.Append("{\"Id\":")
                    .Append(itemCount)
                    .Append(",\"Name\":\"Name")
                    .Append(itemCount)
                    .Append("\",\"Category\":\"Group")
                    .Append(itemCount % 10)
                    .Append("\",\"Age\":")
                    .Append(20 + (itemCount % 50))
                    .Append(",\"Padding\":\"")
                    .Append(padding)
                    .Append("\"}");

                itemCount++;
            }

            builder.Append(']');
            return new LargeJsonPayload(Encoding.UTF8.GetBytes(builder.ToString()), itemCount);
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

        private sealed class LargeJsonPayload
        {
            public LargeJsonPayload(byte[] utf8Bytes, int itemCount)
            {
                Utf8Bytes = utf8Bytes;
                ItemCount = itemCount;
            }

            public byte[] Utf8Bytes { get; }

            public int ItemCount { get; }
        }

        private readonly struct MemorySnapshot
        {
            public MemorySnapshot(long managedBytes, long workingSetBytes)
            {
                ManagedBytes = managedBytes;
                WorkingSetBytes = workingSetBytes;
            }

            public long ManagedBytes { get; }

            public long WorkingSetBytes { get; }

            public static MemorySnapshot Capture()
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();

                var process = Process.GetCurrentProcess();
                process.Refresh();
                return new MemorySnapshot(GC.GetTotalMemory(true), process.WorkingSet64);
            }
        }
    }
}
