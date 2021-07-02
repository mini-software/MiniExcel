## Release  Notes

---

<div align="center">
<a href="https://www.nuget.org/packages/MiniExcel"><img src="https://img.shields.io/nuget/v/MiniExcel.svg" alt="NuGet"></a>  <a href="https://www.nuget.org/packages/MiniExcel"><img src="https://img.shields.io/nuget/dt/MiniExcel.svg" alt=""></a>  <a href="https://ci.appveyor.com/project/shps951023/miniexcel/branch/master"><img src="https://ci.appveyor.com/api/projects/status/b2vustrwsuqx45f4/branch/master?svg=true" alt="Build status"></a>
</div>

<div align="center">
<strong><a href="README.md">English</a> | <a href="README.zh-CN.md">简体中文</a> | <a href="README.zh-Hant.md">繁體中文</a></strong>
</div>

<div align="center">
🙌 <a href="https://github.com/shps951023/MiniExcel">Star</a> can make MiniExcel better 🙌
</div>

---

### 0.17.3 
- [OPT] Clearer exception message when file is illegal excel zip format. #272

### 0.17.2
- [Bug] Fix v0.16.0-0.17.1 custom format contains specific format (eg:`#,##0.000_);[Red]\(#,##0.000\)`), automatic converter will convert double to datetime #267

### 0.17.1
- [New] Add QueryAsDataTableAsync(this Stream stream..)
- [OPT] More clear strong type conversion error message [#I3X2ZL](https://gitee.com/dotnetchina/MiniExcel/issues/I3X2ZL)


### 0.17.0
- [New] Support Async/Task #52, contributor: [isdaniel ( SHIH,BING-SIOU)](https://github.com/isdaniel)

### 0.16.1
- [New] SaveAsByTemplate support DateTime custom format #255, contributor: [网虫 (landde) - Gitee.com](https://gitee.com/landde)

### 0.16.0
- [New] Query support custom datetime format mapping datetime type #256
- [Bug] Fix Query call convertValueByStyleFormat method repeatedly, cause waste of resources #259

### 0.15.5
- [Bug] Chinese env datetime format InvalidCastException #257

### 0.15.4
- [Breaking Change] Set CSV Reader/Writer default encoding : UTF-8 => UTF-8-BOM
- [Breaking Change] Rename CsvConfiguration GetStreamReaderFunc => StreamReaderFunc
- [New] Csv SaveAs support custom StreamWriter

### 0.15.3
- [New] Csv SaveAs support datareader

### 0.15.2
- [New] Support Custom Datetime format #241
- [Bug] Csv type mapping Query error "cannot be converted to xxx type" #243
- [Bug] No error exception throw when reading xls file #242
- [Breaking Change] Stream cannot know the file type, please specify ExcelType manually

### 0.15.1
- [Bug] Fix Sheetxml writer error, it contains two ">" #240

### 0.15.0
- [New] SaveAs change default style and provide style options enum #132
- [New] Support SaveAs by DataSet #235

### 0.14.8 
- [Bug] Fix csv Query split comma not correct #237 #I3R95M
- [Bug] QueryAsDataTable type check problem, e.g A2=5.5 , A3=0.55/1.1 will case double type check error #233

### 0.14.7
- [New] SaveAs Support Create Multiple Sheets
- [Breaking Change] Change GetSheetNames type IEnumerable<string> -> List<string>

### 0.14.6
- [Bug] Fix SaveAs by datareader error "Invalid attempt to call FieldCount when reader is closed" #230

### 0.14.5
- [Breaking Change] Rename OpenXmlConfiguration FillMergedCells

### 0.14.4
- [New] Query support Fill Merged Cells Down [#122](https://github.com/shps951023/MiniExcel/issues/122)
- [Bug] Fix QueryAsDataTable error "Cannot set Column to be null" #229

### 0.14.3
- [Opt] Support Xlsm AutoCheck #227
- [Bug] Fix SaveAsByTemplate single column demension index error [#226](https://github.com/shps951023/MiniExcel/issues/226)

### 0.14.2
- [Bug] Fix asp.net webform gridview datasource can't use QueryAsDataTable [#223](https://github.com/shps951023/MiniExcel/issues/223)

### 0.14.1
- [Bug] Fix custom m/d format not convert datetime [#222](https://github.com/shps951023/MiniExcel/issues/222)

### 0.14.0
- [New] Query、GetColumns support startCell [#147](https://github.com/shps951023/MiniExcel/issues/147)
- [New] GetColumns support read headers

### 0.13.5
- [New] Support QueryAsDataTable method [#216](https://github.com/shps951023/MiniExcel/issues/216)
- [New] SaveAs support IDataReader value parameter [#211](https://github.com/shps951023/MiniExcel/issues/211)
- [Bug] Fix numeric format string will be cast to numeric type [#I3OSKV](https://gitee.com/dotnetchina/MiniExcel/issues/I3OSKV)
- [Opt] Optimize SaveAs convert value type logic to improve performance

### 0.13.4
- [Changed] DataTable use Caption for column name first, then use columname #217
- [New] Type Query support Enum mapping #89
- [OPT] Optimize stream excel type check #215

### 0.13.3
- [New] Support open with read only mode, avoid error of The process cannot access the file because it is being used by another process [#87](https://github.com/shps951023/MiniExcel/issues/87)
- [Breaking Change] Change CSV SaveAs datetime default format : "yyyy-MM-dd HH:mm:ss"
- [Bug] Fixed SaveAsByTemplate when merge cells will cause collection rendering error [#207](https://github.com/shps951023/MiniExcel/issues/207)
- [Bug] Fixed MiniExcel.SaveAs(path, value,sheetName:"Name"), the actual sheetName is Sheet1

### 0.13.2
- [Bug] Fix Column more than 255 rows cannot be read error [#208](https://github.com/shps951023/MiniExcel/issues/208)

### 0.13.1
- [New] SaveAsByTemplate by template bytes, convenient to cache and support multiple users to read the same template at the same time #189
- [New] SaveAsByTemplate support input `IEnmerable<IDicionary<string,object>> or DapperRows or DataTable` parameters [#201](https://github.com/shps951023/MiniExcel/issues/201)
- [Bug] Fix after stream SaveAs/SaveAsByTemplate, miniexcel will close stream [#200](https://github.com/shps951023/MiniExcel/issues/200)

### 0.13.0
- [New] Support .NET Framework 4.5
- [Bug] Fix template excel that with namespace prefix will cause parsing error #193
- [OPT] Optimize template paresing performance

### 0.12.2
- [Bug] Template concating cell value type problem #179
- [Bug] Template fix non-nullable numeric type cell type is 'str' #180

### 0.12.1
- [OPT] Optimize performance of filling excel
- [OPT] Template IEnumerable generate support type auto mapping (Issue #177)
- [New] Support GetColumns method #174
- [New] Template support $rowindex keyword to get current row index
- [Bug] Dimension without x prefix #175

### 0.12.0-beta
- [New] Support `Filling Excel` by SaveAsByTemplate method to fill data into excel by xlsx template

### 0.11.1
- [New] Support ColumnIndex Attribute [#142](https://github.com/shps951023/MiniExcel/issues/142) & [#I3I3EB](https://gitee.com/dotnetchina/MiniExcel/issues/I3I3EB)
- [Bug] Fix issue #157 : Special conditions will get the wrong worksheet name
- [Update] issue #150 : SaveAs input IEnuerable<valuetype> should throw clear msg exception

### 0.11.0
- [New] Added GetSheetNames method support multi-sheets Query
- [New] Query support by sheet name
- [New] Csv SaveAs support DataTable/Dictionary parameters
- [New] CsvConfiguration support custom newLine & seperator & GetStreamReaderFunc
- [OPT] Optimize SaveAs/Query excel file type auto-check

### 0.10.3(Don't use this version)
- [Bug] Fix Query SharedStrings control character not encoding (issue [Issue #149](https://github.com/shps951023/MiniExcel/issues/149))

### 0.10.2(Don't use this version)
- [Bug] Fix SharedStrings get wrong index (issue [#153](https://github.com/shps951023/MiniExcel/issues/153))
- [Bug] SaveAs support control character encoding (issue [Issue #149](https://github.com/shps951023/MiniExcel/issues/149))

### 0.10.1(Don't use this version)
- [New] SaveAs support POCO excel column name/ignore attribute

### 0.10.0(Don't use this version)
- [New] Query dynamic with first head will ignore blank/whitespace columns
- [New] Query type mapping support Custom POCO excel column name/ignore attribute

### 0.9.1(Don't use this version) 
- [Bug] Solve cannot convert Cell value to Nullable<T> (issue #138)

### 0.9.0(Don't use this version)
- [Bug] Solve System.IO.Compression referencing twice  (issue #97)
- [Bug] StrongTypeMapping Query empty row will be generated repeatedly

### 0.8.0(Don't use this version)
- [New] Add MiniExcel.Query by file path method

### 0.7.0(Don't use this version)
- Optimize SaveAs logic
- [Breaking Change] SaveAs value parameter change type check logic

### 0.6.0(Don't use this version)
- [New] SaveAs support parameter IEnumerable deferred execution
- [Breaking Change] Remove SaveAs by object, now only support Datatable,IEnumerable<T>,ICollection<T>
- [Bug] Fix empty rows generate excel error (issue: #128)

### 0.5.0(Don't use this version)
- [New] Support OpenXml Xlsx SaveAs writer mode that avoids OOM
- [Breaking Change] Remove SaveAs startCell parameter
- [Bug] Fix SaveAs dimension printHeader:true not correct 

### 0.4.0(Don't use this version)
- [New] Support create CSV by file path or stream 
- [New] Support csv custom configuration setting
- [New] Support auto/manual specify excel type (xlsx or csv)
- [Breaking Changes] Remove Query First/FirstOrDefault/Single/SingleOrDefault, user can use LINQ method do it.

### 0.3.0(Don't use this version)
- [New] Support SaveAs by IEnumerable of DapperRow and IDictionary<string,object>
- [New] Support dynamic query timespan style format mapping to timespan type.

### 0.2.3(Don't use this version)
- [Bug] Fix ShMemory leak and static problem.
- [New] Support style datetime format mapping to datetime type.

### 0.2.2(Don't use this version) 
- SavaAs support xl/sheet dimension
- [Breaking Changes] SaveAs value type from object to DataTable & ICollection
- [Bug] Fix ICollection with type but no data error (https://github.com/shps951023/MiniExcel/issues/105)

### 0.2.1(Don't use this version)  
- [Optimize] Optimize type mapping bool and datetime auto check
- [New] Query Support xl/worksheets/Sheet Xml Xml `<c>` without `r` attribute or without `<dimension>` but `<c>` with `r` attribute, but now performance is slow than with dimension ([](https://github.com/shps951023/MiniExcel/issues/2))

### 0.2.0(Don't use this version)  
- Release to nuget.org

### 0.1.0-preview(Don't use this version)
- [New] Add Query strongly typed mapping
- [New] Add QueryFirstOrDefault、QuerySingle、QuerySingleOrDefault

### 0.0.7-beta(Don't use this version)
- [New] Add QueryFirst method

### 0.0.6-beta(Don't use this version)
- [Breaking Changes] Replace Create by SavaAs

### 0.0.5-beta(Don't use this version)
- Release remove `assembly: InternalsVisibleTo`

### 0.0.4-beta(Don't use this version)
- [New] Support SaveAs Stream

### 0.0.3-beta(Don't use this version)
- [New] Support Query dynamic and IEnumrable Deferred Execution to avoid OOM
- [New] MiniExcelHelper.Create value type change to ICollection
- [New] Encode XML Value `&apos; &quot; &gt; &lt; &amp;`
- [New] Check Multiple Sheet Index Order
- [New] Dynamic Query support A,B,C.. column name key
- [New] Support insert empty Rows between rows

### 0.0.2-beta(Don't use this version)
- [New] Add MiniExcelHelper.Read Method
- [Breaking Changes] Remove System.IO.Packaging.Package Dependency, and replaced by System.IO.Compression.ZipArchive

### 0.0.1-beta(Don't use this version)
- [New] Add MiniExcelHelper.Create

### 0.0.0
- Init