using MiniExcelLib;
using MiniExcelLib.OpenXml;

if (args.Length is < 1 or > 2)
{
    Console.Error.WriteLine("Usage: MiniExcel.StressTests <xlsx-path> [passes]");
    return 2;
}

var path = Path.GetFullPath(args[0]);
var passes = args.Length == 2 && int.TryParse(args[1], out var parsedPasses) && parsedPasses > 0
    ? parsedPasses
    : 1;

var importer = MiniExcel.Importers.GetOpenXmlImporter();
long rowCount = 0;

for (var pass = 0; pass < passes; pass++)
{
    foreach (var row in importer.Query(path))
    {
        _ = row;
        rowCount++;
    }
}

Console.WriteLine(rowCount);
return 0;