using BenchmarkDotNet.Attributes;
using ClosedXML.Report;
using MiniExcelLibs.Benchmarks.Utils;

namespace MiniExcelLibs.Benchmarks.BenchmarkSections;

public class TemplateXlsxBenchmark : BenchmarkBase
{
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

        MiniExcel.SaveAsByTemplate(path.FilePath, templatePath, value);
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