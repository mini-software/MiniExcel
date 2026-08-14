using System.Text.Json;
using MiniExcelLib.Core.Exceptions;
using MiniExcelLib.Tests.Common.Utils;

namespace MiniExcelLib.OpenXml.Tests.Compatibility;

public class RustParityContractTests
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNameCaseInsensitive = true
    };

    private readonly OpenXmlImporter _excelImporter = MiniExcel.Importers.GetOpenXmlImporter();

    [Fact]
    public void DynamicQueriesMatchSharedRustContract()
    {
        var contract = LoadContract();
        Assert.Equal(1, contract.Version);

        foreach (var testCase in contract.DynamicCases)
        {
            var path = PathHelper.GetFile($"xlsx/{testCase.Fixture}");
            if (testCase.ExpectedSheetNames is not null)
            {
                Assert.Equal(testCase.ExpectedSheetNames, _excelImporter.GetSheetNames(path));
            }

            var configuration = testCase.IgnoreEmptyRows
                ? new OpenXmlConfiguration { IgnoreEmptyRows = true }
                : null;
            var rows = (testCase.EndCell is null
                ? _excelImporter.Query(
                    path,
                    hasHeaderRow: testCase.HasHeader,
                    sheetName: testCase.SheetName,
                    startCell: testCase.StartCell,
                    configuration: configuration)
                : _excelImporter.QueryRange(
                    path,
                    hasHeaderRow: testCase.HasHeader,
                    sheetName: testCase.SheetName,
                    startCell: testCase.StartCell,
                    endCell: testCase.EndCell,
                    configuration: configuration))
                .Cast<IDictionary<string, object?>>()
                .ToList();

            using var stream = File.OpenRead(path);
            var streamRows = (testCase.EndCell is null
                ? _excelImporter.Query(
                    stream,
                    hasHeaderRow: testCase.HasHeader,
                    sheetName: testCase.SheetName,
                    startCell: testCase.StartCell,
                    configuration: configuration)
                : _excelImporter.QueryRange(
                    stream,
                    hasHeaderRow: testCase.HasHeader,
                    sheetName: testCase.SheetName,
                    startCell: testCase.StartCell,
                    endCell: testCase.EndCell,
                    configuration: configuration))
                .Cast<IDictionary<string, object?>>()
                .ToList();
            Assert.True(
                rows.Count == streamRows.Count,
                $"{testCase.Name}: path returned {rows.Count} rows, stream returned {streamRows.Count}");

            Assert.True(
                rows.Count == testCase.RowCount,
                $"{testCase.Name}: expected {testCase.RowCount} rows, got {rows.Count}");

            if (testCase.ExpectedColumns is not null)
            {
                Assert.NotEmpty(rows);
                Assert.Equal(testCase.ExpectedColumns, rows[0].Keys);
                if (testCase.EndCell is null)
                {
                    Assert.Equal(
                        testCase.ExpectedColumns,
                        _excelImporter.GetColumnNames(
                            path,
                            hasHeaderRow: testCase.HasHeader,
                            sheetName: testCase.SheetName,
                            startCell: testCase.StartCell));
                }
            }

            AssertSamples(testCase.Name, rows, testCase.Samples);
        }
    }

    [Fact]
    public void TypedQueriesMatchSharedRustContract()
    {
        var contract = LoadContract();

        foreach (var testCase in contract.TypedCases)
        {
            var rows = ReadTypedRows(testCase);
            Assert.True(
                rows.Count == testCase.RowCount,
                $"{testCase.Name}: expected {testCase.RowCount} rows, got {rows.Count}");
            AssertSamples(testCase.Name, rows, testCase.Samples);
        }
    }

    [Fact]
    public void ConversionErrorsMatchSharedRustContract()
    {
        var contract = LoadContract();

        foreach (var testCase in contract.ErrorCases)
        {
            var path = PathHelper.GetFile($"xlsx/{testCase.Fixture}");
            var exception = testCase.Model switch
            {
                "invalidSequence" => Assert.Throws<ValueNotAssignableException>(
                    () => _excelImporter.Query<ParityInvalidSequence>(path).ToList()),
                _ => throw new InvalidOperationException(
                    $"{testCase.Name}: unsupported error model '{testCase.Model}'")
            };

            Assert.Equal(testCase.ExpectedRow, exception.Row);
            Assert.Equal(testCase.ExpectedValue, Convert.ToString(exception.Value));
        }
    }

    private IReadOnlyList<IDictionary<string, object?>> ReadTypedRows(TypedCase testCase)
    {
        var path = PathHelper.GetFile($"xlsx/{testCase.Fixture}");
        return testCase.Model switch
        {
            "userAccount" => _excelImporter.Query<ParityUserAccount>(path)
                .Select(row => (IDictionary<string, object?>)new Dictionary<string, object?>
                {
                    ["ID"] = row.ID,
                    ["Name"] = row.Name,
                    ["BoD"] = row.BoD,
                    ["Age"] = row.Age,
                    ["VIP"] = row.VIP,
                    ["Points"] = row.Points
                })
                .ToList(),
            "simpleAccount" => _excelImporter.Query<ParitySimpleAccount>(path)
                .Select(row => (IDictionary<string, object?>)new Dictionary<string, object?>
                {
                    ["Name"] = row.Name,
                    ["Age"] = row.Age,
                    ["Mail"] = row.Mail,
                    ["Points"] = row.Points
                })
                .ToList(),
            _ => throw new InvalidOperationException(
                $"{testCase.Name}: unsupported typed model '{testCase.Model}'")
        };
    }

    private static void AssertSamples(
        string caseName,
        IReadOnlyList<IDictionary<string, object?>> rows,
        IReadOnlyList<RowSample> samples)
    {
        foreach (var sample in samples)
        {
            Assert.True(
                sample.RowIndex >= 0 && sample.RowIndex < rows.Count,
                $"{caseName}: missing sample row {sample.RowIndex}");
            var row = rows[sample.RowIndex];

            foreach (var (column, expected) in sample.Cells)
            {
                Assert.True(
                    row.TryGetValue(column, out var value),
                    $"{caseName}: row {sample.RowIndex} does not contain '{column}'");
                Assert.True(
                    Normalize(value) == expected,
                    $"{caseName}: row {sample.RowIndex}, '{column}' expected '{expected}', got '{Normalize(value)}'");
            }
        }
    }

    private static string Normalize(object? value)
    {
        if (value is null or DBNull)
        {
            return "empty:";
        }

        return value switch
        {
            bool boolean => $"bool:{boolean.ToString().ToLowerInvariant()}",
            DateTime dateTime => NormalizeDateTime(dateTime),
            TimeSpan duration => $"duration:{duration.TotalMilliseconds.ToString("0", CultureInfo.InvariantCulture)}",
            Guid guid => $"guid:{guid.ToString("D").ToUpperInvariant()}",
            string text => NormalizeString(text),
            byte or sbyte or short or ushort or int or uint or long or ulong or float or double or decimal =>
                $"number:{Convert.ToDecimal(value, CultureInfo.InvariantCulture).ToString("0.#############################", CultureInfo.InvariantCulture)}",
            _ => $"string:{Convert.ToString(value, CultureInfo.InvariantCulture)}"
        };
    }

    private static string NormalizeString(string value)
    {
        if (Guid.TryParse(value, out var guid))
        {
            return $"guid:{guid.ToString("D").ToUpperInvariant()}";
        }

        string[] formats = ["yyyy-MM-dd'T'HH:mm:ss", "yyyy-MM-dd HH:mm:ss", "yyyy-MM-dd"];
        if (DateTime.TryParseExact(
                value,
                formats,
                CultureInfo.InvariantCulture,
                DateTimeStyles.None,
                out var dateTime))
        {
            return NormalizeDateTime(dateTime);
        }

        return $"string:{value}";
    }

    private static string NormalizeDateTime(DateTime value)
    {
        var text = value.ToString("yyyy-MM-dd'T'HH:mm:ss.fffffff", CultureInfo.InvariantCulture)
            .TrimEnd('0')
            .TrimEnd('.');
        return $"datetime:{text}";
    }

    private static ParityContract LoadContract()
    {
        var path = PathHelper.GetFile("contracts/xlsx-parity-v1.json");
        return JsonSerializer.Deserialize<ParityContract>(File.ReadAllText(path), JsonOptions)
            ?? throw new InvalidOperationException("The XLSX parity contract is empty");
    }

    private sealed class ParityContract
    {
        public int Version { get; set; }
        public List<DynamicCase> DynamicCases { get; set; } = [];
        public List<TypedCase> TypedCases { get; set; } = [];
        public List<ErrorCase> ErrorCases { get; set; } = [];
    }

    private sealed class DynamicCase
    {
        public string Name { get; set; } = "";
        public string Fixture { get; set; } = "";
        public bool HasHeader { get; set; }
        public string? SheetName { get; set; }
        public string StartCell { get; set; } = "A1";
        public string? EndCell { get; set; }
        public bool IgnoreEmptyRows { get; set; }
        public string[]? ExpectedSheetNames { get; set; }
        public int RowCount { get; set; }
        public string[]? ExpectedColumns { get; set; }
        public List<RowSample> Samples { get; set; } = [];
    }

    private sealed class TypedCase
    {
        public string Name { get; set; } = "";
        public string Model { get; set; } = "";
        public string Fixture { get; set; } = "";
        public int RowCount { get; set; }
        public List<RowSample> Samples { get; set; } = [];
    }

    private sealed class RowSample
    {
        public int RowIndex { get; set; }
        public Dictionary<string, string> Cells { get; set; } = [];
    }

    private sealed class ErrorCase
    {
        public string Name { get; set; } = "";
        public string Model { get; set; } = "";
        public string Fixture { get; set; } = "";
        public int ExpectedRow { get; set; }
        public string ExpectedValue { get; set; } = "";
    }

    private sealed class ParityUserAccount
    {
        public Guid ID { get; set; }
        public string? Name { get; set; }
        public DateTime BoD { get; set; }
        public int Age { get; set; }
        public bool VIP { get; set; }
        public decimal Points { get; set; }
    }

    private sealed class ParitySimpleAccount
    {
        public string? Name { get; set; }
        public int Age { get; set; }
        public string? Mail { get; set; }
        public decimal Points { get; set; }
    }

    private sealed class ParityInvalidSequence
    {
        public int ID { get; set; }
        public string? Name { get; set; }
        public int SEQ { get; set; }
    }
}
