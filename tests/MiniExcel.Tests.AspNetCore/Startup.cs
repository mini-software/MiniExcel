using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.DependencyInjection;
using System.IO;
using MiniExcelLibs;
using System;
using System.Net;
using System.Collections.Generic;

public class Startup
{
    public void ConfigureServices(IServiceCollection services) => services.AddMvc();
    public void Configure(IApplicationBuilder app, IWebHostEnvironment env)
    {
        app.UseStaticFiles();
        app.UseRouting();
        app.UseEndpoints(endpoints => endpoints.MapDefaultControllerRoute());
    }
}
public class HomeController : Controller
{
    public IActionResult Index()
    {
        return new ContentResult
        {
            ContentType = "text/html",
            StatusCode = (int)HttpStatusCode.OK,
            Content = @"<html><body>
<a href='Home/DownloadExcel'>DownloadExcel</a><br>
<a href='Home/DownloadExcelFromTmplate'>DownloadExcelFromTmplate</a>
</body></html>"
        };
    }

    public IActionResult DownloadExcel()
    {
        var values = new[] {
            new { Column1 = "MiniExcel", Column2 = 1 },
            new { Column1 = "Github", Column2 = 2}
        };
        var stream = new MemoryStream();
        stream.SaveAs(values);
        stream.Position = 0;
        return File(stream,
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            "demo.xlsx");
    }

    public IActionResult DownloadExcelFromTmplate()
    {
        var templatePath = "TestTemplateComplex.xlsx";
        var value = new Dictionary<string, object>()
        {
            ["title"] = "FooCompany",
            ["managers"] = new[] {
                new {name="Jack",department="HR"},
                new {name="Loan",department="IT"}
            },
            ["employees"] = new[] {
                new {name="Wade",department="HR"},
                new {name="Felix",department="HR"},
                new {name="Eric",department="IT"},
                new {name="Keaton",department="IT"}
            }
        };
        var stream = new MemoryStream();
        stream.SaveAsByTemplate(templatePath, value);
        stream.Position = 0;
        return File(stream,
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            "demo.xlsx");
    }
}
