using System.Text.Json;
using MiniExcelLibs.Cli;
using Xunit;

namespace MiniExcelLibs.Cli.Tests;

public sealed class CliApplicationTests : IDisposable
{
    private readonly string _directory = Path.Combine(Path.GetTempPath(), $"MiniExcel.Cli.Tests-{Guid.NewGuid():N}");

    public CliApplicationTests() => Directory.CreateDirectory(_directory);

    [Fact]
    public void QueryWritesRowsAsJson()
    {
        var input = Path.Combine(_directory, "input.csv");
        File.WriteAllText(input, "Name,Count\r\nMiniExcel,42");
        using var output = new StringWriter();
        using var error = new StringWriter();

        var exitCode = CliApplication.Run(["query", "--input", input], output, error);

        Assert.Equal(0, exitCode);
        Assert.Equal(string.Empty, error.ToString());
        using var document = JsonDocument.Parse(output.ToString());
        var row = Assert.Single(document.RootElement.EnumerateArray());
        Assert.Equal("MiniExcel", row.GetProperty("Name").GetString());
        Assert.Equal("42", row.GetProperty("Count").GetString());
    }

    [Fact]
    public void ConvertCreatesReadableXlsxFile()
    {
        var input = Path.Combine(_directory, "input.csv");
        var outputPath = Path.Combine(_directory, "output.xlsx");
        File.WriteAllText(input, "Name,Count\r\nMiniExcel,42");
        using var output = new StringWriter();
        using var error = new StringWriter();

        var exitCode = CliApplication.Run(["convert", "--input", input, "--output", outputPath], output, error);

        Assert.Equal(0, exitCode);
        Assert.Equal(string.Empty, error.ToString());
        var row = Assert.Single(MiniExcel.Query(outputPath, useHeaderRow: true).Cast<object>());
        var values = Assert.IsAssignableFrom<IDictionary<string, object>>(row);
        Assert.Equal("MiniExcel", values["Name"]);
        Assert.Equal("42", values["Count"]);
    }

    public void Dispose() => Directory.Delete(_directory, recursive: true);
}