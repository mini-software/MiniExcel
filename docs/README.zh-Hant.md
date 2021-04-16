
## 更新日誌

---

[English](README.md) / [繁體中文](README.zh-Hant.md) / [简体中文](README.zh-CN.md)

---

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