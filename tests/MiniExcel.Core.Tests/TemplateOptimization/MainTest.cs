namespace MiniExcelLib.Tests.TemplateOptimization;

public class MainTest
{
    private readonly ITestOutputHelper _testOutputHelper;

    public MainTest(ITestOutputHelper testOutputHelper)
    {
        _testOutputHelper = testOutputHelper;
    }

    internal const int GenerateLimit = 1000;
    
    [Fact]
    public void OptimizationTest()
    {
        var tmpFilePath = Path.Combine(Environment.CurrentDirectory, @"C:\Users\computer\source\repos\MiniExcel\tests\MiniExcel.Core.Tests\TemplateOptimization\template.xlsx");

        var testList = DummyFiller.GenerateDummyFileSystemEntries(GenerateLimit).ToList();
        _testOutputHelper.WriteLine("Getting SIDs from result(owner)");
        var identities = testList.Select(x => x.Owner).ToList();
        _testOutputHelper.WriteLine("Getting SIDs from result(ACLs)");
        foreach (var securityIdentifier in testList.Select(x => x.Acl).SelectMany(sid => sid.Keys))
        {
            identities.Append(securityIdentifier);
        }

        _testOutputHelper.WriteLine("Creating Matrix List");
        var timer = Stopwatch.StartNew();
        var matrixCreator = new MatrixCreator(testList, identities);
        var matrixRows = matrixCreator.Creation().ToList();
        timer.Stop();
        _testOutputHelper.WriteLine($"Created in {timer.Elapsed}");

        var sheets = SheetCreation(matrixRows);
        tmpFilePath = matrixCreator.ExcelFileCreator(tmpFilePath);

        var timestamp = Stopwatch.GetTimestamp();
        var memory = Environment.WorkingSet / (1024 * 1024);
        TestingMethod(matrixCreator, sheets, tmpFilePath, "result_template.xlsx");
        memory = Environment.WorkingSet / (1024 * 1024) - memory;
        var elapsed = Stopwatch.GetElapsedTime(timestamp);
        
        _testOutputHelper.WriteLine($"\n \n Memory used: {memory}MB \nWith Time: {elapsed.TotalMilliseconds} ms \n \n");    }

    private static void TestingMethod(MatrixCreator matrixCreator, Dictionary<string, object> sheets, string templatePath = "", string saveAsName = "result.xlsx")
    {
        var path = Path.Combine(Environment.CurrentDirectory, saveAsName);
        if (File.Exists(path)) File.Delete(path);
        MiniExcel.Templater.GetOpenXmlTemplater().ApplyTemplate(saveAsName, templatePath, sheets);
    }

    private static Dictionary<string, object> SheetCreation(List<ReportHelper.MatrixRow>? matrixRows)
    {
        var sheets = new Dictionary<string, object>();
        if (matrixRows is null)
            return sheets;
        
        var pages = (int)Math.Ceiling(matrixRows.Count / ReportHelper.MaxMatrixRow);
        var matrixLength = matrixRows.First().Access.Length;
        var identityParts = (int)Math.Ceiling(matrixLength / ReportHelper.MaxMatrixIdentityPart);

        for (var page = 0; page < pages; page++)
        {
            for (var part = 0; part < identityParts; part++)
            {
                var startPart = (int)(part * ReportHelper.MaxMatrixIdentityPart);
                var endPart = Math.Min((int)(startPart + ReportHelper.MaxMatrixIdentityPart), matrixLength);

                var partName = identityParts == 1 ? "" : $"-{part + 1}";
                var sheetName = "Matrix" + (page == 0 && identityParts == 1 ? "" : $"_{page + 1}{partName}");

                var chunk = ReportHelper.Page(matrixRows, (int)ReportHelper.MaxMatrixRow, page)
                    .Select(x => x.ToDictionary(startPart, endPart));
                sheets.TryAdd(sheetName, chunk);
                Console.WriteLine($"Adding {sheetName} Sheet");
            }
        }

        return sheets;
    }
}