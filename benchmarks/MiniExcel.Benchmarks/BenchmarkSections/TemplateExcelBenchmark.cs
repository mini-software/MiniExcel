using BenchmarkDotNet.Attributes;
using ClosedXML.Report;
using MiniExcelLib.Benchmarks.Utils;
using MiniExcelLib.Core;
using MiniExcelLib.Core.Mapping;

namespace MiniExcelLib.Benchmarks.BenchmarkSections;

public class TemplateExcelBenchmark : BenchmarkBase
{
    private OpenXmlTemplater _templater;
    private MappingTemplater _mappingTemplater;
    private OpenXmlExporter _exporter;

    public class Employee
    {
        public string Name { get; set; } = "";
        public string Department { get; set; } = "";
    }

    [GlobalSetup]
    public void Setup()
    {
        _templater = MiniExcel.Templaters.GetOpenXmlTemplater();
        _exporter = MiniExcel.Exporters.GetOpenXmlExporter();
        
        var registry = new MappingRegistry();
        registry.Configure<Employee>(config =>
        {
            config.Property(x => x.Name).ToCell("A2");
            config.Property(x => x.Department).ToCell("B2");
        });
        _mappingTemplater = MiniExcel.Templaters.GetMappingTemplater(registry);
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

    [Benchmark(Description = "MiniExcel Mapping Template Generate")]
    public void MiniExcel_Mapping_Template_Generate_Test()
    {
        using var templatePath = AutoDeletingPath.Create();
        var templateData = new[]
        {
            new { A = "Name", B = "Department" },
            new { A = "", B = "" } // Empty row for data
        };
        _exporter.Export(templatePath.FilePath, templateData);
        
        using var outputPath = AutoDeletingPath.Create();
        var employees = Enumerable.Range(1, RowCount)
            .Select(s => new Employee
            {
                Name = "Jack",
                Department = "HR"
            });

        _mappingTemplater.ApplyTemplate(outputPath.FilePath, templatePath.FilePath, employees);
    }
}