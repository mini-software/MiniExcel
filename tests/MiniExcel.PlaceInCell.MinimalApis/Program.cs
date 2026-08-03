using MiniExcel.PlaceInCell.MinimalApis;

var builder = WebApplication.CreateBuilder(args);
var app = builder.Build();

app.MapGet("/", () => Results.Content(
    """
    <html>
    <head><title>PlaceInCell API</title></head>
    <body style="font-family:sans-serif;margin:2rem">
      <h1>Place in Cell 测试接口</h1>
      <ul>
        <li><a href="/api/place-in-cell/stream">GET /api/place-in-cell/stream</a> — 不落地，直接下载 xlsx</li>
        <li><a href="/api/place-in-cell/save">GET /api/place-in-cell/save</a> — 保存到桌面并返回路径</li>
      </ul>
    </body>
    </html>
    """,
    "text/html; charset=utf-8"));

app.MapPlaceInCellApi();
app.Run();
