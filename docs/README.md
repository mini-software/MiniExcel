

## Release  Notes

### 0.2.1  
- Optimize type mapping bool and datetime auto check
- Query Support xl/worksheets/Sheet Xml Xml `<c>` without `r` attribute or without `<dimension>` but `<c>` with `r` attribute, but now performance is slow than with dimension ([](https://github.com/shps951023/MiniExcel/issues/2))


### 0.2.0  
- Release to nuget.org

### 0.1.0-preview
- Add Query strongly typed mapping
- Add QueryFirstOrDefault、QuerySingle、QuerySingleOrDefault

### 0.0.7-beta
- Add QueryFirst method

### 0.0.6-beta
- [Breaking Changes]Replace Create by SavaAs

### 0.0.5-beta
- Release remove `assembly: InternalsVisibleTo`

### 0.0.4-beta
- Support SaveAs Stream

### 0.0.3-beta
- Support Query dynamic and IEnumrable lazy loading to avoid OOM
- MiniExcelHelper.Create value type change to ICollection
- Encode XML Value `&apos; &quot; &gt; &lt; &amp;`
- Check Multiple Sheet Index Order
- Dynamic Query support A,B,C.. column name key
- Support insert empty Rows between rows

### 0.0.2-beta
- Remove System.IO.Packaging.Package Dependency, and replaced by System.IO.Compression.ZipArchive
- Add MiniExcelHelper.Read Method

### 0.0.1-beta
- Add MiniExcelHelper.Create

### 0.0.0
- Init