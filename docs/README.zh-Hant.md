## 更新日誌

---

<div align="center">
<p><a href="https://www.nuget.org/packages/MiniExcel"><img src="https://img.shields.io/nuget/v/MiniExcel.svg" alt="NuGet"></a>  <a href="https://www.nuget.org/packages/MiniExcel"><img src="https://img.shields.io/nuget/dt/MiniExcel.svg" alt=""></a>  
<a href="https://ci.appveyor.com/project/shps951023/miniexcel/branch/master"><img src="https://ci.appveyor.com/api/projects/status/b2vustrwsuqx45f4/branch/master?svg=true" alt="Build status"></a>
<a href="https://gitee.com/dotnetchina/MiniExcel"><img src="https://gitee.com/dotnetchina/MiniExcel/badge/star.svg" alt="star"></a> <a href="https://github.com/shps951023/MiniExcel" rel="nofollow"><img src="https://img.shields.io/github/stars/shps951023/MiniExcel?logo=github" alt="GitHub stars"></a> 
<a href="https://www.nuget.org/packages/MiniExcel"><img src="https://img.shields.io/badge/.NET-%3E%3D%204.5-red.svg" alt="version"></a>
</p>
</div>

---

<div align="center">
<p><strong><a href="README.md">English</a> | <a href="README.zh-CN.md">簡體中文</a> | <a href="README.zh-Hant.md">繁體中文</a></strong></p>
</div>

---

<div align="center">
<p> 您的 <a href="https://github.com/shps951023/miniexcel">Star</a> 和 <a href="https://miniexcel.github.io">贊助</a> 能幫助 MiniExcel 成長 </p>
</div>

---



## 1.28.0

- [New] 支持 CSV Insert 方法 #I4X92G (via @shps951023)

### 1.27.0

- [New] 支持 DateTimeOffset and ExcelFormat #430 (via @Lightczx , @shps951023 )

- [Optimization] SaveAs by datareader 支持 dimension #231 (via @shps951023)

### 1.26.7

- [OPT] 減少 memory allocation 使用 MemoryStream #427 (via @cupsos)
- [OPT] 添加 System.Memory pacakge #427 (via @cupsos)
- [OPT] 減少 memory allocation in GetImageFormat() #427 (via @cupsos)
- [Bug] 修正 MiniExcel.SaveAsByTemplate value 為 List<Dictionary<string, object>> 異常錯誤 #413 (via @shps951023)

### 1.26.6
- [OPT] Template save performance #425 (via @lileyzhao)

### 1.26.5
- [New] Added DataReader AutoFilter toggle #402 #401 (via @Rollerss)
- [New] SaveAs 支持空白 sharedstring #405

### 1.26.4
- [Bug] 使用Stream.SaveAs時指定excelType會自動關閉Stream #I57WMM
- [OPT] 減少在讀取大文件時 _IntMappingAlphabet.Count 的調用 (#400 via @CollapseNav)

### 1.26.3
- [OPT] Export 預設 buffersize 從 1024 bytes -> 1024 * 512 bytes
- [New] Export 支持自定義 buffersize 
- [New] SaveAsByTemplate number 改為 InvariantCulture (via @psxbox)

### 1.26.2
- [Bug] 修正 DynamicColumnAttribute Ignore, Index 問題 #377

### 1.26.1
- [New] DynamicColumnAttribute 支持 Dictionary #370
- [Bug] 修正 MiniExcelDataReader SqlBulkCopy 中斷問題 (via @yfl8910)

### 1.26.0
- [New] 支持 DynamicColumnAttribute (via @y976362357, @shps951023)
- [New] 支持 ExcelColumnAttribute 合併現有 attributes (#357 via @Weilence)
- [OPT] ExpandoObject 效能增強，在需要時再轉換 Type. (#366 via @isdaniel)
- [OPT] 優化 startswith & endwith 效能 (#365 via @isdaniel)

### 1.25.2
- [New] 移除過期的 ExcelNumberFormat Dependency #271

### 1.25.1
- [Bug] 修正 QueryAsDataTable 讀取大文件會拋出 NotImplementedException #360

### 1.25.0
- [New] 支持 SharingStrings disk cache (當該文件大小 >= 5 MB)，現在讀取 2GB SharingStrings 只需要使用 1~13MB 記憶體  (#346 via @Weilence) (via @shps951023)
- [New] Async 支持 cancellationToken  (#350 via @isdaniel)
- [New] SaveAs 支持 overwriteFile 參數，方便調整是否要覆蓋已存在文件。 #307
- [Bug] 修正 SaveAs by datareader， 有時會多一個 autoFilter column #352

### 1.24.3
- [Bug] 修正多 threads Async 可能錯誤 'The given key N was not present in the dictionary' #344
- [Bug] 修正當 CultureInfo 像是`ff-Latn` , datareader field 類型是 datetime 系統會生成錯誤 OA Date 格式 #343

### 1.24.2
- [Bug] Query<T> 有多個相同標題會導致StartCell無法正確取得該欄位 #I4YCLQ 
-  [OPT] 優化 Query<T> 的算法

### 1.24.1
- [Bug] 修正 QueryAsync configulation 沒有效果問題 #338 
- [Bug] 修正 QueryAsync 無法使用 dynamic 類別

### 1.24.0

- [New] Query 支持 strict open xml, 感謝 [Weilence (Lowell)](https://github.com/Weilence)  #335 
- [New] SaveAs 以自訂的 CultureInfo 轉換 Format cell 值, 感謝[0xced (Cédric Luthi)](https://github.com/0xced) #333
- [New] SaveAsByTemplate 預設忽略 template 缺少參數 key 錯誤, OpenXmlConfiguration.IgnoreTemplateParameterMissing 可以開關此卡控. #I4WXFB
- [New] SaveAsByTemplate 當參數集合為空時會清空模版字串. #I4WM67

### 1.23.3
- [Bug] SaveAs CSV 當 value 為 DataTable 時，Key包含雙引號Column Name不會顯示`"`。 #I4WDA9

### 1.23.2
- [New] 支持 System.ComponentModel.DisplayName 的 `[DisplayName]` 作為excel標題 [#I4TXGT](https://gitee.com/dotnetchina/MiniExcel/issues/I4TXGT)
- [Bug] 修正  `Czech` 等國家CultureInfo會生成錯誤 decimal 數字提示 #331

### 1.23.0
- [New] 新增 `GetReader` 方法 #328 #290 (感謝 [楊福來 Yang](https://github.com/yfl8910) )

### 1.22.0

- [New] SaveAs 支持自定義 CultureInfo #316
- [New] Query 支持自定義 CultureInfo #316
- [New] 新 byte array 轉換器 #327
- [Breaking Change] 移除 Byte Array 與 base64 轉換器
- [Breaking Change] `EnableConvertByteArray` 取代 `ConvertByteArrayToBase64String` 

### 0.21.5
- [Bug] 修正 SaveAs multiple sheet value error "Excel completed file level validation and repair. Some parts of this workbook may have been repaired or discarded."  #325

### 0.21.4
- [New] Remove LISENCE_CODE check

### 0.21.1
- [New] Check License Code

### 0.21.0
- [New] ExcelFormat 支持 DateTimeOffset/Decimal/double 等類別 format #I49RZH #312 #305
- [New] 支持byte文件導入/導出
- [New] SaveAs 支持預設轉換byte[] 值為 base64 字串
- [New] Query 支持轉換 base64 字串值為 bytep[]
- [New] OpenXmlConfiguration 增加 `ConvertByteArrayToBase64String` 屬性來開關 base64 轉換器
- [New] Query 支持 ExcelInvalidCastException 儲存行、列、值數據 #309

### 0.20.0 
- [New] SaveAs 支持圖片生成 #304
- [Opt] 提高 SaveAs 執行效率

### 0.19.3-beta
- [Fix]  Excelnumberformat 1.1.0 憑證過期 (Valid from: 2018-04-10 08:00:00 to 2021-04-14 20:00:00) [link](https://github.com/andersnm/ExcelNumberFormat/issues/34)

### 0.19.2
- [New] SaveAsByTemplate 支持 datareader [#I4HL54](https://gitee.com/dotnetchina/MiniExcel/issues/I4HL54)

### 0.19.1
- [New] QueryAsDataTable 刪除空白 Column keys. #298
- [Bug] Error NU3037: ExcelNumberFormat 1.1.0 #302
- [Bug] SaveAs 前綴和後綴空格會丟失 #294

### 0.19.0
- [New] SaveAs 預設樣式增加篩選功能. #190
- [New] 新增 ConvertCsvToXlsx、ConvertXlsxToCsv 方法. #292
- [New] OpenXmlConfiguration 新增 AutoFilter 屬性.
- [Bug] 修正 CSV 讀取後 SaveAs 會拋出 "Stream was not readable." 錯誤. #293
- [Bug] 修正 SaveAsByTemplate & 被轉成 &amp; [I4DQUN](https://gitee.com/dotnetchina/MiniExcel/issues/I4DQUN)

### 0.18.0
- [New] SaveAs 支持 enum description #I49RYZ
- [New] Query 強型別支持多列名對應同一屬性 [#I40QA5](https://gitee.com/dotnetchina/MiniExcel/issues/I40QA5)
- [Breaking Change] SaveAs 傳空`IEnumerable<強型別>`現在會生成表頭. #133
- [Bug] SaveAs sheet enum 映射 cell 錯誤 number 型別 #286

### 0.17.5
- [Bug] 修復xlsx文件中標題字段包含`&,<,>,",'`符號，導出後打不開文件

### 0.17.4 
- [Bug] 修復0.17.3版本更新導致 SaveAs 創建 xlsx 文件錯誤 "XML error : Catastrophic failure"

### 0.17.3
- [New] 支持設定列寬 #280
- [Bug] 修復 csv 不支持 QueryAsDataTable #279
- [OPT] 更加清楚的錯誤訊息，當擋案是錯誤 excel zip 格式 #272

### 0.17.2
- [Bug] 修復 v0.16.0-0.17.1 自定義格式含有特定格式(e.g:`#,##0.000_);[Red]\(#,##0.000\)`)，自動轉換器會將 double 被轉成 datetime 異常 #267

### 0.17.1
- [New] 增加 QueryAsDataTableAsync(this Stream stream..)
- [OPT] 強型別 Query 轉型錯誤訊息能知道在哪一行列出錯 [#I3X2ZL](https://gitee.com/dotnetchina/MiniExcel/issues/I3X2ZL)

### 0.17.0
- [New] 支持 Async/Task #52, contributor: [isdaniel ( SHIH,BING-SIOU)](https://github.com/isdaniel)

### 0.16.1
- [New] SaveAsByTemplate 支持 DateTime 自定義格式 #255, contributor: [網蟲 (landde) - Gitee.com](https://gitee.com/landde)

### 0.16.0
- [New] Query 支持自定義日期格式轉成 datetime 型別 #256
- [Bug] 修正 Query 重複呼叫 convertValueByStyleFormat 方法造成資源浪費 #259

### 0.15.5
- [Bug] 特定中文環境日期格式轉換InvalidCastException錯誤 #257

### 0.15.4
- [Breaking Change] CSV Reader/Writer 預設編碼改變 : UTF-8 => UTF-8-BOM
- [Breaking Change] 重新命名 CsvConfiguration GetStreamReaderFunc => StreamReaderFunc
- [New] Csv SaveAs 支持自定義 StreamWriter

### 0.15.3
- [New] Csv SaveAs 支持 datareader

### 0.15.2
- [New] 支持自定義日期時間格式 #241
- [Bug] CSV類型映射查詢錯誤 "cannot be converted to xxx type" #243
- [Bug] Stream 讀取 xls 文件時沒有錯誤異常拋出 #242
- [Breaking Change] 流無法識別文件類型，請手動指定ExcelType

### 0.15.1
- [Bug] 修正 Sheetxml 結尾包含兩個 ">" 導致解析錯誤 #240

### 0.15.0
- [New] SaveAs 更改預設樣式、並提供樣式選擇 #132
- [New] SaveAs 支持 DataSet #235

### 0.14.8 
- [Bug] 修正 csv Query 內文包含逗號造成異常 #237 #I3R95M
- [Bug] 修正 QueryAsDataTable 類別檢查異常，如 A2=5.5 , A3=0.55/1.1 系統會顯示 double type check error #233

### 0.14.7
- [New] SaveAs 支持建立多工作表
- [Breaking Change] 更換 GetSheetNames 返回類型 IEnumerable<string> -> List<string>

### 0.14.6
- [Bug] 修正 SaveAs by datareader 錯誤 "Invalid attempt to call FieldCount when reader is closed" #230

### 0.14.5
- [Breaking Change] 更換 OpenXmlConfiguration FillMergedCells 名稱

### 0.14.4
- [New] Query 支持向下填充合併的單元格 [#122](https://github.com/shps951023/MiniExcel/issues/122)
- [Bug] 修正 QueryAsDataTable 錯誤 "Cannot set Column to be null" #229

### 0.14.3
- [Opt] 支持 Xlsm 自動判斷 #227
- [Bug] 修正 SaveAsByTemplate 單列 demension 索引錯誤 [#226](https://github.com/shps951023/MiniExcel/issues/226)

### 0.14.2
- [Bug] 修正 asp.net webform gridview datasource 不能使用 QueryAsDataTable [#223](https://github.com/shps951023/MiniExcel/issues/223)

### 0.14.1
- [Bug] 修正自定義 m/d 格式沒轉成 datetime [#222](https://github.com/shps951023/MiniExcel/issues/222)

### 0.14.0
- [New] Query、GetColumns 支持 startCell 能指定 cell 開始讀取資料 [#147](https://github.com/shps951023/MiniExcel/issues/147)
- [New] GetColumns 支持讀取表頭

### 0.13.5
- [New] 新增 QueryAsDataTable 方法 [#216](https://github.com/shps951023/MiniExcel/issues/216)
- [New] SaveAs 支持 IDataReader value 參數 [#211](https://github.com/shps951023/MiniExcel/issues/211)
- [Bug] 修正數字格式的字串會被強制轉換為decimal類型 [#I3OSKV](https://gitee.com/dotnetchina/MiniExcel/issues/I3OSKV)
- [Opt] 優化 SaveAs 類別轉換算法，避免效率浪費

### 0.13.4
- [Changed] DataTable 以 Caption 優先當欄位名稱 #217
- [New] Query 支持 Enum mapping #89
- [OPT] 優化 stream excel 類別(xlsx or csv)檢查 #215

### 0.13.3
- [New] 支持 Excel 單純讀取模式，避免同時改模版又運行 MiniExcel 出現錯誤 "The process cannot access the file because it is being used by another process" [#87](https://github.com/shps951023/MiniExcel/issues/87)
- [Breaking Change] CSV SaveAs datetime 預設格式改為 "yyyy-MM-dd HH:mm:ss"
- [Bug] 修正模版模式集合渲染遇到合併列會出現異常問題 [#207](https://github.com/shps951023/MiniExcel/issues/207)
- [Bug] 修正 MiniExcel.SaveAs(path, value,sheetName:"Name"), 實際 sheetName 是 Sheet1

### 0.13.2
- [Bug] 超過 255 列無法讀取錯誤 [#208](https://github.com/shps951023/MiniExcel/issues/208)

### 0.13.1
- [New] SaveAsByTemplate 支持讀取模板 byte[],方便緩存跟支持多用戶同時讀取同一個模板 [#189](https://github.com/shps951023/MiniExcel/issues/189)
- [New] SaveAsByTemplate 支持傳入 `IEnmerable<IDicionary<string,object>> 或 DapperRows 或 DataTable` 參數 [#201](https://github.com/shps951023/MiniExcel/issues/201)
- [Bug] 修正使用 stream SaveAs/SaveAsByTemplate 系統會自動關閉流 stream [#200](https://github.com/shps951023/MiniExcel/issues/200)

### 0.13.0
- [New] 支持 .NET Framework 4.5 以上版本
- [Bug] 修正特殊情況Excel模板含有 namespace prefixFix 會造成模板解析錯誤 #193
- [OPT] 優化模板解析效率

### 0.12.2
- [Bug] 修正模板串接 Cell 類別不是字串問題 #179
- [Bug] 修正模板遇到非數字類別 t 是 str 問題 #180

### 0.12.1
- [OPT] 優化填充 Excel 效率
- [OPT] 模板集合列表支持類別自動判斷 (Issue #177)
- [New] 新增 GetColumns 方法 (Issue #174)
- [New] 模板支持 $rowindex 關鍵字獲取當前列索引
- [Bug] Dimension 沒有 x 字首 (Issue #175)

### 0.12.0-beta
- [New] 支持`填充Excel`模式 ，借由 SaveAsByTemplate 方法以模板填充資料，

### 0.11.1
- [New] 支持 ColumnIndex Attribute [#142](https://github.com/shps951023/MiniExcel/issues/142) & [#I3I3EB](https://gitee.com/dotnetchina/MiniExcel/issues/I3I3EB)
- [Bug] 修正 issue #157 : 特別情況無法使用指定 sheet name 查詢
- [Update] issue #150 : SaveAs 值集合錯誤訊息更明細

### 0.11.0
- [New] 添加 GetSheetNames 方法支持多 sheet 查詢
- [New] Query 指定 sheet 名稱
- [New] Csv SaveAs 支持 DataTable/Dictionary 參數
- [New] CsvConfiguration 支持自訂義 newLine & seperator & GetStreamReaderFunc
- [Optimization] 優化 SaveAs/Query excel 檔案類型自動判斷

### 0.10.3(請勿使用) 
- [Bug] 修正 Query SharedStrings 控制字符沒有 encoding (issue [Issue #149](https://github.com/shps951023/MiniExcel/issues/149))

### 0.10.2
- [Bug] 修正 SharedStrings get wrong index (issue [#153](https://github.com/shps951023/MiniExcel/issues/153))
- [Bug] SaveAs 支持 control character encoding (issue [Issue #149](https://github.com/shps951023/MiniExcel/issues/149))

### 0.10.1(請勿使用) 
- [New] SaveAs 支持 POCO excel 欄位名稱/忽略 attribute

### 0.10.0(請勿使用) 
- [New] Query dynamic 表頭預設自動忽略空白字串欄位
- [New] Query 強型別支持自訂義 excel 欄位名稱/忽略 attribute

### 0.9.1(請勿使用) 
- [Bug] 解決無法 mapping Cell Value 到 Nullable 屬性類別 (issue #138)

### 0.9.0(請勿使用)
- [Bug] 解決 System.IO.Compression 引用兩次問題  (issue #97)
- [Bug] 強型別 Query 空列會重複複製問題

### 0.8.0(請勿使用)
- [New] MiniExcel.Query 支持檔案路徑查詢

### 0.7.0(請勿使用)
- 優化 SaveAs 效率
- [Breaking Change] SaveAs value 參數類別檢查邏輯

### 0.6.0(請勿使用)
- [New] SaveAs 支持 參數 IEnumerable 延遲查詢
- [Breaking Change] 移除 SaveAs by object, 現在只支持 Datatable,IEnumerable<T>,ICollection<T>
- [Bug] 修正空列生成 excel 錯誤 (issue: #128)

### 0.5.0(請勿使用)
- [New] 支持 OpenXml Xlsx SaveAs writer 模式避免 OOM
- [Breaking Change] 移除 SaveAs startCell 參數
- [Bug] 修正 SaveAs dimension printHeader:true 異常

### 0.4.0(請勿使用)
- [New] 支持 create CSV by 檔案路徑或是 stream 
- [New] 支持 csv 自訂義 configuration 
- [New] 支持自動/手動指定 excel 類型 (xlsx or csv)
- [Breaking Changes] 移除 Query First/FirstOrDefault/Single/SingleOrDefault 方法, 使用者使用 LINQ 即可

### 0.3.0(請勿使用)
- [New] 支持 SaveAs by IEnumerable of DapperRow and IDictionary<string,object>
- [New] 支持 dynamic query timespan style 格式 mapping timespan 類別

### 0.2.3(請勿使用)
- [Bug] 修正記憶體洩漏問題
- [New] 支持 style datetime 格式 mapping datetime 類別.

### 0.2.2(請勿使用)
- SavaAs 支持 xl/sheet dimension
- [Breaking Changes] SaveAs value 類別準許 object & DataTable & ICollection
- [Bug] 修正 ICollection with type 沒有資料錯誤 (https://github.com/shps951023/MiniExcel/issues/105)

### 0.2.1(請勿使用)  
- [Optimize] Optimize type mapping bool and datetime auto check
- [New] Query 支持 xl/worksheets/Sheet Xml `<c>` 沒有 `r` 屬性或是沒有 `<dimension>` 但 `<c>` 有 `r` 屬性情況, 但是效率會遠低於有 dimension ([](https://github.com/shps951023/MiniExcel/issues/2))

### 0.2.0(請勿使用)  
- 發布至 nuget.org

### 0.1.0-preview
- [New] 添加 Query 強型別 mapping
- [New] 添加 QueryFirstOrDefault、QuerySingle、QuerySingleOrDefault

### 0.0.7-beta
- [New] 添加 QueryFirst 方法

### 0.0.6-beta
- [Breaking Changes] 替換 Create 名稱為 SavaAs

### 0.0.5-beta
- [Bug] Release 刪除 `assembly: InternalsVisibleTo` 依賴

### 0.0.4-beta
- [New] 支持 SaveAs Stream

### 0.0.3-beta
- [New] 支持 Query dynamic and IEnumrable 延遲查詢避免 OOM
- [New] MiniExcelHelper.Create value 類別換成 ICollection
- [New] Encode XML 值 `&apos; &quot; &gt; &lt; &amp;`
- [New] 檢查多 Sheet Index 排序
- [New] Dynamic Query 支持 A,B,C.. 欄位名稱 key
- [New] 支持列與列之間空列情況

### 0.0.2-beta
- [New] 添加 MiniExcelHelper.Read 方法
- [Breaking Changes] 移除 System.IO.Packaging.Package 依賴, 換成 System.IO.Compression.ZipArchive

### 0.0.1-beta
- [New] 添加 MiniExcelHelper.Create 方法

### 0.0.0
- Init