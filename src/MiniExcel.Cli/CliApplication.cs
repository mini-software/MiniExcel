using System.CommandLine;
using System.Globalization;
using System.Text.Json;

namespace MiniExcelLibs.Cli;

public static class CliApplication
{
    public static int Run(string[] args, TextWriter output, TextWriter error)
    {
        var rootCommand = new RootCommand("Query and convert Excel and CSV files with MiniExcel.");
        rootCommand.Subcommands.Add(CreateConvertCommand(output, error));
        rootCommand.Subcommands.Add(CreateQueryCommand(output, error));
        return rootCommand.Parse(args).Invoke();
    }

    private static Command CreateConvertCommand(TextWriter output, TextWriter error)
    {
        var inputOption = RequiredFileOption("--input", "Input .xlsx, .xlsm, or .csv file.");
        var outputOption = RequiredFileOption("--output", "Output .xlsx or .csv file.");
        var sheetOption = new Option<string?>("--sheet") { Description = "Worksheet name to read." };
        var noHeaderOption = new Option<bool>("--no-header") { Description = "Treat the first row as data instead of column names." };
        var overwriteOption = new Option<bool>("--overwrite") { Description = "Overwrite the output file if it exists." };

        var command = new Command("convert", "Convert between Excel and CSV formats.");
        command.Options.Add(inputOption);
        command.Options.Add(outputOption);
        command.Options.Add(sheetOption);
        command.Options.Add(noHeaderOption);
        command.Options.Add(overwriteOption);
        command.SetAction(parseResult => Execute(
            () => Convert(
                parseResult.GetValue(inputOption)!,
                parseResult.GetValue(outputOption)!,
                parseResult.GetValue(sheetOption),
                !parseResult.GetValue(noHeaderOption),
                parseResult.GetValue(overwriteOption)),
            output,
            error));
        return command;
    }

    private static Command CreateQueryCommand(TextWriter output, TextWriter error)
    {
        var inputOption = RequiredFileOption("--input", "Input .xlsx, .xlsm, or .csv file.");
        var sheetOption = new Option<string?>("--sheet") { Description = "Worksheet name to read." };
        var noHeaderOption = new Option<bool>("--no-header") { Description = "Treat the first row as data instead of column names." };
        var startCellOption = new Option<string>("--start-cell") { Description = "First cell to read.", DefaultValueFactory = _ => "A1" };

        var command = new Command("query", "Write rows as a JSON array.");
        command.Options.Add(inputOption);
        command.Options.Add(sheetOption);
        command.Options.Add(noHeaderOption);
        command.Options.Add(startCellOption);
        command.SetAction(parseResult => Execute(
            () => Query(
                parseResult.GetValue(inputOption)!,
                parseResult.GetValue(sheetOption),
                !parseResult.GetValue(noHeaderOption),
                parseResult.GetValue(startCellOption)!,
                output),
            output,
            error));
        return command;
    }

    private static Option<FileInfo> RequiredFileOption(string name, string description) =>
        new(name) { Description = description, Required = true };

    private static int Execute(Action action, TextWriter output, TextWriter error)
    {
        try
        {
            action();
            return 0;
        }
        catch (Exception exception)
        {
            error.WriteLine($"miniexcel: {exception.Message}");
            return 1;
        }
    }

    private static void Convert(FileInfo input, FileInfo output, string? sheet, bool useHeaderRow, bool overwrite)
    {
        EnsureInputExists(input);
        EnsureSupportedOutput(output);
        var rows = MiniExcel.Query(input.FullName, useHeaderRow, sheet);
        MiniExcel.SaveAs(output.FullName, rows, overwriteFile: overwrite);
    }

    private static void Query(FileInfo input, string? sheet, bool useHeaderRow, string startCell, TextWriter output)
    {
        EnsureInputExists(input);
        using var stream = new MemoryStream();
        using (var writer = new Utf8JsonWriter(stream, new JsonWriterOptions { Indented = true }))
        {
            writer.WriteStartArray();
            foreach (var item in MiniExcel.Query(input.FullName, useHeaderRow, sheet, startCell: startCell).Cast<object>())
            {
                var row = (IDictionary<string, object?>)item;
                writer.WriteStartObject();
                foreach (var cell in row)
                {
                    writer.WritePropertyName(cell.Key);
                    WriteValue(writer, cell.Value);
                }
                writer.WriteEndObject();
            }
            writer.WriteEndArray();
        }
        output.WriteLine(System.Text.Encoding.UTF8.GetString(stream.ToArray()));
    }

    private static void WriteValue(Utf8JsonWriter writer, object? value)
    {
        switch (value)
        {
            case null: writer.WriteNullValue(); break;
            case bool boolean: writer.WriteBooleanValue(boolean); break;
            case byte number: writer.WriteNumberValue(number); break;
            case sbyte number: writer.WriteNumberValue(number); break;
            case short number: writer.WriteNumberValue(number); break;
            case ushort number: writer.WriteNumberValue(number); break;
            case int number: writer.WriteNumberValue(number); break;
            case uint number: writer.WriteNumberValue(number); break;
            case long number: writer.WriteNumberValue(number); break;
            case ulong number: writer.WriteNumberValue(number); break;
            case float number: writer.WriteNumberValue(number); break;
            case double number: writer.WriteNumberValue(number); break;
            case decimal number: writer.WriteNumberValue(number); break;
            case DateTime dateTime: writer.WriteStringValue(dateTime); break;
            case DateTimeOffset dateTimeOffset: writer.WriteStringValue(dateTimeOffset); break;
            case TimeSpan timeSpan: writer.WriteStringValue(timeSpan.ToString("c", CultureInfo.InvariantCulture)); break;
            case Guid guid: writer.WriteStringValue(guid); break;
            default: writer.WriteStringValue(System.Convert.ToString(value, CultureInfo.InvariantCulture)); break;
        }
    }

    private static void EnsureInputExists(FileInfo input)
    {
        if (!input.Exists)
            throw new FileNotFoundException($"Input file not found: {input.FullName}", input.FullName);
    }

    private static void EnsureSupportedOutput(FileInfo output)
    {
        var extension = output.Extension.ToLowerInvariant();
        if (extension is not ".xlsx" and not ".csv")
            throw new ArgumentException("Output file must use the .xlsx or .csv extension.", nameof(output));
    }
}