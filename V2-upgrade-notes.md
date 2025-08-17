## MiniExcel 2.0 Migration Guide

- The root namespace was changed from `MiniExcelLibs` to `MiniExcelLib`. If the full MiniExcel package is downloaded, the previous namespace will still exist and will contain the now old and deprecated methods' signatures
- Instead of having all methods being part of the `MiniExcel` static class, the functionalities are now split into 3 providers accessible from the same class:
`MiniExcel.Importers`, `MiniExcel.Exporters` and `MiniExcel.Templaters` will give you access to, respectively, the `MiniExcelImporterProvider`, `MiniExcelExporterProvider` and `MiniExcelTemplaterProvider`
- This way Excel and Csv query methods are split between the `OpenXmlImporter` and the `CsvImporter`, accessible from the `MiniExcelImporterProvider`
- The same division was adopted for export methods with `OpenXmlExporter` and `CsvExporter`
- Template methods are instead currently only found in `OpenXmlTemplater`
- Csv methods are only available if the MiniExcel.Csv package is installed, which is pulled down automatically when the full MiniExcel package is downloaded
- `IConfiguration` is now `IMiniExcelConfiguration`, but most methods now require the proper implementation (`OpenXmlConfiguration` or `CsvConfiguration`) to be provided rather than the interface
- MiniExcel now fully supports asynchronous streaming the queries, 
so the return type for `OpenXmlImporter.QueryAsync` is `IAsyncEnumerable<T>` instead of `Task<IEnumerable<T>>` 