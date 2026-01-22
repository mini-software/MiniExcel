## MiniExcel 2.0 Upgrade Notes

- Support for .NET Framework 4.5 was dropped, the minimum supported Framework version is now 4.6.2. 
- The root namespace was changed from `MiniExcelLibs` to `MiniExcelLib`.  
- Instead of having all methods being part of the `MiniExcel` static class, the functionalities are now split into 3 providers:
`MiniExcel.Importers`, `MiniExcel.Exporters` and `MiniExcel.Templaters` will give you access to, respectively, the `MiniExcelImporterProvider`, `MiniExcelExporterProvider` and `MiniExcelTemplaterProvider`.
- This way Excel and Csv query methods are split between the `OpenXmlImporter` and the `CsvImporter`, accessible from the `MiniExcelImporterProvider`.
- The same structure was adopted for export methods through `OpenXmlExporter` and `CsvExporter`, while template methods are instead currently only found in `OpenXmlTemplater`.
- Csv methods are only available if the MiniExcel.Csv package is installed, which is pulled down automatically when the full MiniExcel package is downloaded.
- You can only access the conversion methods `ConvertCsvToXlsx` and `ConvertXlsxToCsv` from the `MiniExcelConverter` utility class, which is part of the full MiniExcel package.
- If the full MiniExcel package is downloaded, the previous namespace will coexist along the new one, containing the original static methods' signatures, which have become a facade for the aferomentioned providers.
- `IConfiguration` is now `IMiniExcelConfiguration`, but most methods now require the proper implementation (`OpenXmlConfiguration` or `CsvConfiguration`) to be provided rather than the interface
- MiniExcel now fully supports asynchronous streaming the queries, 
so the return type for `OpenXmlImporter.QueryAsync` is `IAsyncEnumerable<T>` instead of `Task<IEnumerable<T>>`