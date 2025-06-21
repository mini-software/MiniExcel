## Release  Notes

<div align="center">
<p><a href="https://www.nuget.org/packages/MiniExcel"><img src="https://img.shields.io/nuget/v/MiniExcel.svg" alt="NuGet"></a>  <a href="https://www.nuget.org/packages/MiniExcel"><img src="https://img.shields.io/nuget/dt/MiniExcel.svg" alt=""></a>  
<a href="https://ci.appveyor.com/project/mini-software/miniexcel/branch/master"><img src="https://ci.appveyor.com/api/projects/status/b2vustrwsuqx45f4/branch/master?svg=true" alt="Build status"></a>
<a href="https://gitee.com/dotnetchina/MiniExcel"><img src="https://gitee.com/dotnetchina/MiniExcel/badge/star.svg" alt="star"></a> <a href="https://github.com/mini-software/MiniExcel" rel="nofollow"><img src="https://img.shields.io/github/stars/mini-software/MiniExcel?logo=github" alt="GitHub stars"></a> 
<a href="https://www.nuget.org/packages/MiniExcel"><img src="https://img.shields.io/badge/.NET-%3E%3D%204.5-red.svg" alt="version"></a>
</p>
</div>

---

<div align="center">
<p><strong><a href="README.md">English</a> | <a href="README.zh-CN.md">简体中文</a> | <a href="README.zh-Hant.md">繁體中文</a></strong></p>
</div>

---

<div align="center">
 Your <a href="https://github.com/mini-software/MiniExcel">Star</a> and <a href="https://miniexcel.github.io">Donate</a> can make MiniExcel better 
</div>

---


### 1.41.2
- [New] Fixes enum behaviour and adds support for DescriptionAttribute when saving by template (via @michelebastione )
- [Bug] SaveAsByTemplate - Excel Dimension Xml is null #459 (via @michelebastione )
- [Other] Benchmarks refactoring and adaptation for making them run on Github Actions (#777) (via @michelebastione )
- [Other] add deepwiki link and badge (via @isdaniel )


### 1.41.1
- [New] Added sheet dimensions feature (#774) (via @michelebastione)
- [Bug] Fix Saving by template introduces unintended value replication in each row #771
(via @shps951023)
- [Bug] SaveAsByTemplate not working in > v1.39.0 for templates with columns past Z #773 (via @shps951023)
- [Bug] Fix Excel style.xml duplicate numFmtId, system error: An item with the same key has already been added. #772 (via @shps951023)


### 1.41.0
- [New] MiniExcel AddPicture #770 (via @shps951023)
- [New] Add a configuration item in OpenXmlConfiguration to control writing the file path, enabling the corresponding cell to remain empty when importing images. (via @dotnetfans)


### 1.40.1
- [New] Added support for DateOnly type in query mapping (via @michelebastione)
- [New] Added StyleOptions class to OpenXmlConfiguration to allow more direct control over cell styles (#761) (via @michelebastione)
- [Bug] Fix SaveAsByTemplate empty generated result (via @shps951023)


### 1.40.0
- [New] Added exception to warn for sheet name too long (via @michelebastione)
- [New] Added options for trimming column names and ignoring empty rows (via @michelebastione)
- [New] Added IAsyncDisposable calls to ExcelOpenXmlsheetWriter.Async (via @michelebastione)
- [New] Add conditional formatting support to Excel template (#745) (via @Notallthatevil)
- [New] Add support for .NET 9.0 in project file (#744) (via @Notallthatevil)
- [Bug] Bugfix of invalid cell values being mistakenly parsed as valid (via @michelebastione)
- [Bug] Changing NotImplementedException assert in some tests to NotSupportedException (via @michelebastione)
- [Bug] Moved sheet name check and fixed datetime formatting bug (via @michelebastione)
- [OPT] Optimize max memory usage of SaveAsByTemplate #750 (#752) (via @shps951023)
- [OPT] Optimization of SaveAsTemplate method (#749) (via @michelebastione)
- [OPT] Removing DateTimeHelper.FromOADate and related code (via @michelebastione)
- [OPT] Remove redundant property in csproj file (#748) (via @AZhrZho)
- [Breaking Change] QueryRange doesn't support column name without row number #763 (via @michelebastione)


### 1.39.0
- [New] Added support for Uri mapping (#726) (via @michelebastione)
- [New] Added insert sheet feature about ContentTypesXml processing (#728) (via @izanhzh)
- [New] Supports the TimeSpan type, double.NaN exports invalid values, and when reading, it needs to be determined whether it is a double value. (via @wxn401)
- [Bug] Fixed parsing bug in the conversion to double (#734) (via @michelebastione)
- [Bug] Fixed configuration is not used when writing using IDataReader #735 (via @michelebastione)
- [Bug] Fixed cancellation token not working for some async methods, enhanced export methods by returning number of rows, added active tab functionality, tests and code cleanup (#738) (via @michelebastione)


### 1.38.0
- [New] feat(DynamicExcelColumn): make the CustomFormatter property more powerful (#715) (via @izanhzh)
- [New] feat(ExcelNumberFormat): extend the SectionType (#716) (via @izanhzh)
- [New] feat(ExcelOpenXmlSheetWriter): adjust the location of the CustomFormatter execution (#722) (via @izanhzh)
- [New] WriteEmptyStringAsNull implementation (#725) (via @jiangyi1985)
- [Bug] Fix @group tag not working with certain IEnumerable types (#723) (via @JamesDSource)
- [OPT] Optimized ContainsKey calls to TryGetValue (via @michelebastione) 
- [OPT] Changed Count() method calls to Count property (via @michelebastione) 
- [OPT] Materialized some IEnumerables using ToList (via @michelebastione) 
- [OPT] Added safe IDisposable cast to IEnumerator (via @michelebastione) 
- [OPT] Removed superfluous indentation and parenthesis (via @michelebastione) 

### 1.37.0
- [New] feat: support insert sheet (#709) (via @izanhzh)
- [Bug] fix(MiniExcelDataReader): GetOrdinal always returns 0 (#711) (via @izanhzh)
- [OPT] Generalize excel writing with a common write adapter and implement writing IAsyncEnumerable (#712) (via @Discolai , @izanhzh)

### 1.36.1
- [New] feat(MiniExcelDataReaderBase): add asynchronous support (#706) (via @izanhzh , @ArgoZhang )

### 1.36.0
- [New] Write auto column width (#695) (via @Discolai)
- [New] Enhance IDataReader export with DynamicColumnFirst and Custom Formatting Delegate (#700) (via @IcedMango)
- [Bug] If cellValue is string no matter that it contains only numbers will put the value as text. Preventing of losing leading zeroes (via @psyhlo)


### 1.35.0

- [New] Formula attribute added to support in rows with dto or dynamic attributes (#679) (via @RaZer0k & Co-authored-by: Eulises Vargas )
- [New] Async implementation of freezing top row (#684)  (via @BaatenHannes )
- [New] Upgrade to .NET 8.0 and refactor input value extraction (#681) (via @ramioh )
- [Bug] Yield empty self-closing row tags as empty row during query. (#673) (via @aulickiDnv )



### 1.34.2

- [New] Add MniExcelDataReaderBase class to simplify code (#651) (via @ArgoZhang )
- [OPT] perf csv insert (#653) (via @izanhzh )
- [Bug] Fix dimension writing in FastMode (#659) (via @Discolai )
- [Bug] Fix QueryAsDataTable can't read Excel with only header rows (#647) (via @Discolai )


### 1.34.1
- [Bug] Fix Issue 632, refactor sheet styles (#640) (via @meld-cp)
- [Bug] Fix StartSheetView for multiple selection (#641) (via @jiaguangli)

### 1.34.0
- [New] Add freeze panes (#626) (via @meld-cp)
- [New] Add DateTime Nullale support (via @jiaguangli)
- [OPT] Reduce memory requirements when processing templates + template formulas (#638) (via @meld-cp)
- [Bug] Fix problem with multi-line when using Query (#628) (via @meld-cp)
- [Bug] Fix empty data reader issue. (#629) (via @duszekmestre)
- [Bug] Fix Fields of type long cannot be set to text normally #627 (via @shps951023)

### 1.33.0
- [New] Support Template formulas (#622) (via @meld-cp)
- [New] Update DynamicConfiguration format (#595) (via @wangboshun)
- [OPT] CSV enumeration and code reusage (#600) (via @duszekmestre)
- [OPT] 1900 year DateTime correction #599 (via @duszekmestre)

### 1.32.1
- [New] .NET5^ support image `ReadOnlySpan<byte>`  (via @shps951023)
- [Bug] Remove bug with Portable.System.DateTimeOnly and only support DateOnly .NET6^ #594 (via @shps951023)

### 1.32.0
- [New] Using DynamicConfiguration when writing data using DataTable (via @pszybiak)
- [New] Make System.DateOnly available as date in Excel (#576) (via @ofthelit)
- [New] Allow ampersand in sheet names (via @ofthelit)
- [OPT] Use true async processing for excel writer (#573) (via @duszekmestre)
- [Bug] Fix nullable enumeration conversion failure #567) (via @PurplestViper)
- [Bug] IEnumerable traversed twice #422 (via @Discolai)
- [Bug] Fix Read empty string as null (via @pszybiak)
- [Bug] Fix CSV Reader Query faild on specific xlsx file, throws ExcelColumnNotFoundException (via @AZhrZho)
- [Bug] Fix DynamicExcelColumn configuration while saving IDictionary values (via @johannes-barta)
- [Bug] Fix propInfo.Key missing (via @shps951023)
- [Bug] Fix sharedStrings ref #549 (via @shps951023)

### 1.31.3
- [Bug] DescriptionAttr null check(via @wulaoh)
- [Bug] Throw custom exception when CSV column not found #543 (via @pszybiak)
- [Bug] SaveAsByTemplate rowInfo.IEnumerableMercell.Height null exception #553 (via @shps951023)

### 1.31.2
- [New] Support automatic merge for same vertical cells between @merge and @endmerge tags (via @eynarhaji)
- [New] Limit merge tagged columns with @mergelimit column. First merge limited column and then merge other columns accordingly. (via @eynarhaji)
- [New] Support dynamic columns when generating sheet by IDataReader to change columns names & widths #514 (via @Laxynium)
- [Bug] Fix R1C1 reference to A1 reference bug (via @ivan132)

### 1.31.1
- [OPT] Support property cache #23 (via @RRQM_Home)

### 1.31.0
- [New] Support Fields #490 (via @jsgervais)
- [New] Support skipping null values when writing to Excel #497 (via @0MG-DEN)
- [Bug] Fix calc chain.xml  #491(via @ArgoZhang)
- [Bug] Support some sheet `/xl` location error  #494 (via @ArgoZhang)

### 1.30.3
- [New] support if/else statements inside cell (via @eynarhaji)

### 1.30.2
- [New] support grouped rows (via @eynarhaji)
- [New] support automatic merge for same vertical cells (via @eynarhaji)

### 1.30.1
- [New] support function to custom separator (via @hyzx86)
- [New] support config for get sheet names (via @H4ad)

### 1.30.0
- [New] Remove .NET 5.0 support
- [New] support excel enum description string to enum #289 (via @KaneLeung)


### 1.29.0

- [New] SaveAs support FastMode
- [Bug] Fixed SaveAs OOM problem

### 1.28.2

- [New] Support Assembly Strong Name Signature #450
- [New] Support QueryRange (via @1ras1)

### 1.28.1

- [Optimization] Reduce string memory allocation when template save #439 (via @cupsos)
- [Optimization] Remove dependency System.Memory #441 (via @ping9719)

### 1.28.0

- [New] Support CSV Insert #I4X92G (via @shps951023)

### 1.27.0

- [New] Support DateTimeOffset and ExcelFormat #430 (via @Lightczx , @shps951023 )
- [Optimization] SaveAs by datareader support dimension #231 (via @shps951023)

### 1.26.7

- [OPT] Reduce memory allocation when using MemoryStream #427 (via @cupsos)
- [OPT] Add System.Memory pacakge #427 (via @cupsos)
- [OPT] Reduce memory allocation in GetImageFormat() #427 (via @cupsos)
- [Bug] Fixed MiniExcel.SaveAsByTemplate error when value is List<Dictionary<string, object>> #413 (via @shps951023)

### 1.26.6
- [OPT] Template save performance #425 (via @lileyzhao)

### 1.26.5
- [New] Added DataReader AutoFilter toggle #402 #401 (via @Rollerss)
- [New] SaveAs support empty sharedstring #405

### 1.26.4
- [Bug] Using stream.SaveAs will close the Stream automatically when Specifying excelType
- [OPT] Optimize Query big file  _IntMappingAlphabet.Count hot loading count (#400 via @CollapseNav)

### 1.26.3
- [OPT] Export default buffersize from 1024 bytes -> 1024 * 512 bytes
- [New] Export support custom buffersize 
- [New] SaveAsByTemplate number use InvariantCulture (via @psxbox)

### 1.26.2

- [Bug] Fixed DynamicColumnAttribute Ignore, Index error #377

### 1.26.1
- [New] DynamicColumnAttribute support Dictionary #370
- [Bug] Fixed MiniExcelDataReader SqlBulkCopy error (via @yfl8910)

### 1.26.0
- [New] Support DynamicColumnAttribute (via @y976362357, @shps951023)
- [New] Support ExcelColumnAttribute to merge attributes (#357 via @Weilence)
- [OPT]  Only when necessary system will converts ExpandoObject type  (#366 via @isdaniel)
- [OPT] Optimize startsWith & EndWith performance (#365 via @isdaniel)

### 1.25.2

- [New] Remove overdue ExcelNumberFormat Dependency #271

### 1.25.1
- [Bug] Fixed QueryAsDataTable read big file will throw NotImplementedException #360

### 1.25.0
- [New] Support SharingStrings disk cache (when this file size >= 5 MB), it can reduce reading 2GB SharingStrings only needs 1~13 MB memory #117(#346 via @Weilence) (via @shps951023)
- [New] Async support cancellationToken (#350 via @isdaniel)
- [New] SaveAs support overwriteFile parameter for enable/unable overwriting exist file #307
- [Bug] SaveAs by datareader, sometime it will add one more autoFilter column #352

### 1.24.3
- [Bug] Fixed multiple threads Async error 'The given key N was not present in the dictionary' #344
- [Bug] Fixed when CultureInfo likes`ff-Latn` , datareader field type is datetime that will get error OA Date format #343

### 1.24.2
- [Bug] Fiexd Query multiple same title will cause startcell to get wrong column index #I4YCLQ 
- [OPT] Optimize Query<T> algorithm

### 1.24.1
- [Bug] Fiexd QueryAsync configulation not working #338 
- [Bug] Fixed QueryAsync not return dynamic type

### 1.24.0
- [New] Query support strict open xml, thanks [Weilence (Lowell)](https://github.com/Weilence)  #335 
- [New] SaveAs use the configured CultureInfo to write format cell value, thanks [0xced (Cédric Luthi)](https://github.com/0xced) #333
- [New] SaveAsByTemplate default ignore template missing parameter key exception, OpenXmlConfiguration.IgnoreTemplateParameterMissing can control it. #I4WXFB
- [New] SaveAsByTemplate will clean template string when parameter is IEnumerable and empty collection. #I4WM67


### 1.23.3
- [Bug] SaveAs CSV when value is DataTable, if Key contains `"` then column name will not show `"`。 #I4WDA9

### 1.23.2

- [New] Support System.ComponentModel.DisplayName's `[DisplayName]` as title [#I4TXGT](https://gitee.com/dotnetchina/MiniExcel/issues/I4TXGT)
- [Bug] Fix when CultureInfo like `Czech` will get invalid output with decimal numbers #331

### 1.23.0
- [New] Support `GetReader` method #328 #290 (Thanks [杨福来 Yang](https://github.com/yfl8910) )


### 1.22.0

- [New] SaveAs support to custom CultureInfo #316
- [New] Query support to custom CultureInfo #316
- [New] New efficiency byte array Converter #327
- [Breaking Change] Remove Byte Array to base64 Converter
- [Breaking Change] Replace `ConvertByteArrayToBase64String` by `EnableConvertByteArray`

### 0.21.5
- [Bug] Fix SaveAs multiple sheet value error "Excel completed file level validation and repair. Some parts of this workbook may have been repaired or discarded."  #325

### 0.21.4
- [New] Remove LISENCE_CODE check

### 0.21.1
- [New] Check License Code

### 0.21.0
- [New] ExcelFormat support DateTimeOffset/Decimal/double etc. type format #I49RZH #312 #305
- [New] Support byte file import/export
- [New] SaveAs support to convert byte[] value to base64 string 
- [New] Query support to convert base64 value to byte[]
- [New] OpenXmlConfiguration add `ConvertByteArrayToBase64String` to turn on/off base64 convertor
- [New] Query support ExcelInvalidCastException to store column, row, value data #309


### 0.20.0 
- [New] SaveAs support image #304
- [Opt] Improve SaveAs efficiency 

### 0.19.3-beta
- [Fix]  Excelnumberformat 1.1.0 valid date expired (Valid from: 2018-04-10 08:00:00 to 2021-04-14 20:00:00) [link](https://github.com/andersnm/ExcelNumberFormat/issues/34)

### 0.19.2
- [New] SaveAsByTemplate support datareader [#I4HL54](https://gitee.com/dotnetchina/MiniExcel/issues/I4HL54)

### 0.19.1
- [New] QueryAsDataTable remove empty column keys. #298
- [Bug] Error NU3037: ExcelNumberFormat 1.1.0 #302
- [Bug] Prefix and suffix blank space will lost after SaveAs #294

### 0.19.0
- [New] SaveAs default style with autoFilter mode. #190
- [New] Add ConvertCsvToXlsx、ConvertXlsxToCsv method. #292
- [New] OpenXmlConfiguration add AutoFilter property. 
- [Bug] Fix after CSV Query then SaveAs system will throw "Stream was not readable." exception. #293
- [Bug] Fix SaveAsByTemplate & convert to &amp; [I4DQUN](https://gitee.com/dotnetchina/MiniExcel/issues/I4DQUN)

### 0.18.0
- [New] SaveAs support enum description #I49RYZ
- [New] Query strong type support multiple column names mapping to the same property. [#I40QA5](https://gitee.com/dotnetchina/MiniExcel/issues/I40QA5)
- [Breaking Change] SaveAs by empty IEnumerable<StrongType> will generate header now empty rows now. #133 
- [Bug] SaveAs sheet enum mapping cell number type #286

### 0.17.5
- [Bug] Fix xlsx file header column name with `&,<,>,",'`, the file cannot be opened.

### 0.17.4 
- [Bug] Fix v0.17.3 SaveAs xlsx file will cause "XML error : Catastrophic failure"

### 0.17.3 
- [New] Support set column width #280
- [Bug] Fix csv not support QueryAsDataTable #279
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
- [New] Query support Fill Merged Cells Down [#122](https://github.com/mini-software/MiniExcel/issues/122)
- [Bug] Fix QueryAsDataTable error "Cannot set Column to be null" #229

### 0.14.3
- [Opt] Support Xlsm AutoCheck #227
- [Bug] Fix SaveAsByTemplate single column demension index error [#226](https://github.com/mini-software/MiniExcel/issues/226)

### 0.14.2
- [Bug] Fix asp.net webform gridview datasource can't use QueryAsDataTable [#223](https://github.com/mini-software/MiniExcel/issues/223)

### 0.14.1
- [Bug] Fix custom m/d format not convert datetime [#222](https://github.com/mini-software/MiniExcel/issues/222)

### 0.14.0
- [New] Query、GetColumns support startCell [#147](https://github.com/mini-software/MiniExcel/issues/147)
- [New] GetColumns support read headers

### 0.13.5
- [New] Support QueryAsDataTable method [#216](https://github.com/mini-software/MiniExcel/issues/216)
- [New] SaveAs support IDataReader value parameter [#211](https://github.com/mini-software/MiniExcel/issues/211)
- [Bug] Fix numeric format string will be cast to numeric type [#I3OSKV](https://gitee.com/dotnetchina/MiniExcel/issues/I3OSKV)
- [Opt] Optimize SaveAs convert value type logic to improve performance

### 0.13.4
- [Changed] DataTable use Caption for column name first, then use columname #217
- [New] Type Query support Enum mapping #89
- [OPT] Optimize stream excel type check #215

### 0.13.3
- [New] Support open with read only mode, avoid error of The process cannot access the file because it is being used by another process [#87](https://github.com/mini-software/MiniExcel/issues/87)
- [Breaking Change] Change CSV SaveAs datetime default format : "yyyy-MM-dd HH:mm:ss"
- [Bug] Fixed SaveAsByTemplate when merge cells will cause collection rendering error [#207](https://github.com/mini-software/MiniExcel/issues/207)
- [Bug] Fixed MiniExcel.SaveAs(path, value,sheetName:"Name"), the actual sheetName is Sheet1

### 0.13.2
- [Bug] Fix Column more than 255 rows cannot be read error [#208](https://github.com/mini-software/MiniExcel/issues/208)

### 0.13.1
- [New] SaveAsByTemplate by template bytes, convenient to cache and support multiple users to read the same template at the same time #189
- [New] SaveAsByTemplate support input `IEnmerable<IDicionary<string,object>> or DapperRows or DataTable` parameters [#201](https://github.com/mini-software/MiniExcel/issues/201)
- [Bug] Fix after stream SaveAs/SaveAsByTemplate, miniexcel will close stream [#200](https://github.com/mini-software/MiniExcel/issues/200)

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
- [New] Support ColumnIndex Attribute [#142](https://github.com/mini-software/MiniExcel/issues/142) & [#I3I3EB](https://gitee.com/dotnetchina/MiniExcel/issues/I3I3EB)
- [Bug] Fix issue #157 : Special conditions will get the wrong worksheet name
- [Update] issue #150 : SaveAs input IEnuerable<valuetype> should throw clear msg exception

### 0.11.0
- [New] Added GetSheetNames method support multi-sheets Query
- [New] Query support by sheet name
- [New] Csv SaveAs support DataTable/Dictionary parameters
- [New] CsvConfiguration support custom newLine & seperator & GetStreamReaderFunc
- [OPT] Optimize SaveAs/Query excel file type auto-check

### 0.10.3(Don't use this version)
- [Bug] Fix Query SharedStrings control character not encoding (issue [Issue #149](https://github.com/mini-software/MiniExcel/issues/149))

### 0.10.2(Don't use this version)
- [Bug] Fix SharedStrings get wrong index (issue [#153](https://github.com/mini-software/MiniExcel/issues/153))
- [Bug] SaveAs support control character encoding (issue [Issue #149](https://github.com/mini-software/MiniExcel/issues/149))

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
- [Bug] Fix ICollection with type but no data error (https://github.com/mini-software/MiniExcel/issues/105)

### 0.2.1(Don't use this version)  
- [Optimize] Optimize type mapping bool and datetime auto check
- [New] Query Support xl/worksheets/Sheet Xml Xml `<c>` without `r` attribute or without `<dimension>` but `<c>` with `r` attribute, but now performance is slow than with dimension ([](https://github.com/mini-software/MiniExcel/issues/2))

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