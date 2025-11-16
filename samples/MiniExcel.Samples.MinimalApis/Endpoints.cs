using System.Text.Json;
using Microsoft.AspNetCore.Mvc;
using MiniExcelLib.Core;

namespace MiniExcel.Samples.MinimalApis;

internal static class Endpoints
{
    private static readonly OpenXmlImporter Importer = MiniExcelLib.Core.MiniExcel.Importers.GetOpenXmlImporter();
    private static readonly OpenXmlExporter Exporter = MiniExcelLib.Core.MiniExcel.Exporters.GetOpenXmlExporter();
    private static readonly OpenXmlTemplater Templater = MiniExcelLib.Core.MiniExcel.Templaters.GetOpenXmlTemplater();

    internal static RouteGroupBuilder MapExcelApi(this IEndpointRouteBuilder builder)
    {
        var group = builder.MapGroup("api");

        group.MapImportExcel();
        group.MapExportExcel();
        group.MapApplyExcelTemplate();

        return group;
    }

    private static RouteHandlerBuilder MapExportExcel(this IEndpointRouteBuilder builder)
    {
        return builder.MapGet("export", async () =>
        {
            object[] values = 
            [
                new { Column1 = "MiniExcel", Column2 = 1 },
                new { Column1 = "Github", Column2 = 2}
            ];
            
            var memoryStream = new MemoryStream();
            await Exporter.ExportAsync(memoryStream, values);
            memoryStream.Seek(0, SeekOrigin.Begin);
            
            return Results.Stream(memoryStream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "export_demo.xlsx");
        });
    }

    private static RouteHandlerBuilder MapApplyExcelTemplate(this IEndpointRouteBuilder builder)
    {
        return builder.MapGet("template", async () =>
        {
            var value = new Dictionary<string, object>
            {
                ["title"] = "FooCompany",
                
                ["managers"] = new[] 
                {
                    new { name = "Jack", department = "HR" },
                    new { name = "Loan", department = "IT" }
                },
                
                ["employees"] = new[] 
                {
                    new { name = "Wade", department = "HR" },
                    new { name = "Felix", department = "HR" },
                    new { name = "Eric", department = "IT" },
                    new { name = "Keaton", department = "IT" }
                }
            };

            var memoryStream = new MemoryStream();
            await Templater.ApplyTemplateAsync("TestTemplateComplex.xlsx", memoryStream, value);
            memoryStream.Seek(0, SeekOrigin.Begin);
            
            return Results.Stream(memoryStream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "template_demo.xlsx");
        });
    }

    private static RouteHandlerBuilder MapImportExcel(this IEndpointRouteBuilder builder)
    {
        return builder.MapPost("import", async ([FromBody] IFormFile file) =>
        {
            var stream = new MemoryStream();
            await file.CopyToAsync(stream);

            var result = new List<dynamic>();
            await foreach (var item in Importer.QueryAsync(stream, useHeaderRow: true))
            {
                // your logic here
                result.Add(item);
            }

            var limit = Math.Min(10, result.Count);
            return Results.Ok("File uploaded successfully\ndata: "+ JsonSerializer.Serialize(result.Take(limit)));
        });
    }
}