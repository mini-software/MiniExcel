using MiniExcel.Samples.MinimalApis;

var builder = WebApplication.CreateBuilder(args);
var app = builder.Build();

app.MapExcelApi();
app.Run();
