using BenchmarkDotNet.Attributes;
using ClosedXML.Report;
using MiniExcelLib.Benchmarks.Utils;
using MiniExcelLib.Core;

namespace MiniExcelLib.Benchmarks.BenchmarkSections;

public class TemplateExcelBenchmark : BenchmarkBase
{
    private OpenXmlTemplater _templater;

    [GlobalSetup]
    public void Setup()
    {
        _templater = MiniExcel.Templaters.GetOpenXmlTemplater();
    }
    
    [Benchmark(Description = "MiniExcel Template Generate")]
    public void MiniExcel_Template_Generate_Test()
    {
        const string templatePath = "TestTemplateBasicIEmumerableFill.xlsx";
        
        using var path = AutoDeletingPath.Create();
        var value = new
        {
            employees = Enumerable.Range(1, RowCount)
                .Select(s => new
                {
                    name = "Jack",
                    department = "HR"
                })
        };

        _templater.ApplyTemplate(path.FilePath, templatePath, value);
    }

    [Benchmark(Description = "ClosedXml.Report Template Generate")]
    public void ClosedXml_Report_Template_Generate_Test()
    {
        const string templatePath = "TestTemplateBasicIEmumerableFill_ClosedXML_Report.xlsx";
        
        using var path = AutoDeletingPath.Create();
        var value = new
        {
            employees = Enumerable.Range(1, RowCount)
                .Select(_ => new
                {
                    name = "Jack",
                    department = "HR"
                })
        };

        var template = new XLTemplate(templatePath);
        template.AddVariable(value);
        template.Generate();

        template.SaveAs(path.FilePath);
    }
}