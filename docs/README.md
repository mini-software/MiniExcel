

## Release  Notes

### 0.11.0
- [New] Added GetSheetNames method support multi-sheets Query
- [New] Query support by sheet name
- Optimiz SaveAs/Query excel file type auto-check

### 0.10.3 
- [Bug] Fix Query SharedStrings control character not encoding (issue [Issue #149](https://github.com/shps951023/MiniExcel/issues/149))

### 0.10.2
- [Bug] Fix SharedStrings get wrong index (issue [#153](https://github.com/shps951023/MiniExcel/issues/153))
- [Bug] SaveAs support control character encoding (issue [Issue #149](https://github.com/shps951023/MiniExcel/issues/149))

### 0.10.1
- [New] SaveAs support POCO excel column name/ignore attribute

### 0.10.0
- [New] Query dynamic with first head will ignore blank/whitespace columns
- [New] Query type mapping support Custom POCO excel column name/ignore attribute

### 0.9.1 
- [Bug] Solve cannot convert Cell value to Nullable<T> (issue #138)

### 0.9.0
- [Bug] Solve System.IO.Compression referencing twice  (issue #97)
- [Bug] StrongTypeMapping Query empty row will be generated repeatedly

### 0.8.0
- [New] Add MiniExcel.Query by file path method

### 0.7.0
- Optimize SaveAs logic
- [Breaking Change] SaveAs value parameter change type check logic

### 0.6.0
- [New] SaveAs support parameter IEnumerable deferred execution
- [Breaking Change] Remove SaveAs by object, now only support Datatable,IEnumerable<T>,ICollection<T>
- [Bug] Fix empty rows generate excel error (issue: #128)

### 0.5.0
- [New] Support OpenXml Xlsx SaveAs writer mode that avoids OOM
- [Breaking Change] Remove SaveAs startCell parameter
- [Bug] Fix SaveAs dimension printHeader:true not correct 

### 0.4.0
- [New] Support create CSV by file path or stream 
- [New] Support csv custom configuration setting
- [New] Support auto/manual specify excel type (xlsx or csv)
- [Breaking Changes] Remove Query First/FirstOrDefault/Single/SingleOrDefault, user can use LINQ method do it.

### 0.3.0
- [New] Support SaveAs by IEnumerable of DapperRow and IDictionary<string,object>
- [New] Support dynamic query timespan style format mapping to timespan type.

### 0.2.3
- [Bug] Fix ShMemory leak and static problem.
- [New] Support style datetime format mapping to datetime type.

### 0.2.2 
- SavaAs support xl/sheet dimension
- [Breaking Changes] SaveAs value type from object to DataTable & ICollection
- [Bug] Fix ICollection with type but no data error (https://github.com/shps951023/MiniExcel/issues/105)

### 0.2.1  
- [Optimize] Optimize type mapping bool and datetime auto check
- [New] Query Support xl/worksheets/Sheet Xml Xml `<c>` without `r` attribute or without `<dimension>` but `<c>` with `r` attribute, but now performance is slow than with dimension ([](https://github.com/shps951023/MiniExcel/issues/2))

### 0.2.0  
- Release to nuget.org

### 0.1.0-preview
- [New] Add Query strongly typed mapping
- [New] Add QueryFirstOrDefault、QuerySingle、QuerySingleOrDefault

### 0.0.7-beta
- [New] Add QueryFirst method

### 0.0.6-beta
- [Breaking Changes] Replace Create by SavaAs

### 0.0.5-beta
- Release remove `assembly: InternalsVisibleTo`

### 0.0.4-beta
- [New] Support SaveAs Stream

### 0.0.3-beta
- [New] Support Query dynamic and IEnumrable Deferred Execution to avoid OOM
- [New] MiniExcelHelper.Create value type change to ICollection
- [New] Encode XML Value `&apos; &quot; &gt; &lt; &amp;`
- [New] Check Multiple Sheet Index Order
- [New] Dynamic Query support A,B,C.. column name key
- [New] Support insert empty Rows between rows

### 0.0.2-beta
- [New] Add MiniExcelHelper.Read Method
- [Breaking Changes] Remove System.IO.Packaging.Package Dependency, and replaced by System.IO.Compression.ZipArchive

### 0.0.1-beta
- [New] Add MiniExcelHelper.Create

### 0.0.0
- Init