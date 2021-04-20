
## 更新日志

---

[English](README.md) / [简体中文](README.zh-CN.md) / [繁體中文](README.zh-Hant.md) 

---

### 0.13.3
- [New] 支持 Excel 单纯读取模式，避免同时改模版又运行 MiniExcel 出现错误 "The process cannot access the file because it is being used by another process" [#87](https://github.com/shps951023/MiniExcel/issues/#87)

### 0.13.2
- [Bug] 超过 255 列无法读取错误 [#208](https://github.com/shps951023/MiniExcel/issues/#208)

### 0.13.1
- [New] SaveAsByTemplate 支持读取模板 byte[],方便缓存跟支持多用户同时读取同一个模板 #189
- [New] SaveAsByTemplate 支持传入 `IEnmerable<IDicionary<string,object>> 或 DapperRows 或 DataTable` 参数 [#201](https://github.com/shps951023/MiniExcel/issues/201)
- [Bug] 修正使用 stream SaveAs/SaveAsByTemplate 系统会自动关闭流 stream [#200](https://github.com/shps951023/MiniExcel/issues/200)

### 0.13.0
- [New] 支持 .NET Framework 4.5 以上版本
- [Bug] 修正特殊情况Excel模板含有 namespace prefixFix 会造成模板解析错误 #193
- [OPT] 优化模板解析效率

### 0.12.2
- [Bug] 修正模板串接 Cell 類別不是字串問題 #179
- [Bug] 修正模板遇到非数字类别 t 是 str 問題 #180

### 0.12.1
- [OPT] 优化填充 Excel 效率
- [OPT] 模板集合列表支持类别自动判断 (Issue #177)
- [New] 新增 GetColumns 方法 (Issue #174)
- [New] 模板支持 $rowindex 关键字获取当前列索引
- [Bug] Dimension 没有 x 字首 (Issue #175)


### 0.12.0-beta
- [New] 支持`填充Excel`模式 ，借由 SaveAsByTemplate 方法以模板填充数据，

### 0.11.1
- [New] 支持 ColumnIndex Attribute [#142](https://github.com/shps951023/MiniExcel/issues/142) & [#I3I3EB](https://gitee.com/dotnetchina/MiniExcel/issues/I3I3EB)
- [Bug] 修正 issue #157 : 特别情况无法使用指定 sheet name 查询
- [Update] issue #150 : SaveAs 值集合错误信息更明细

### 0.11.0
- [New] 添加 GetSheetNames 方法支持多 sheet 查询
- [New] Query 指定 sheet 名称
- [New] Csv SaveAs 支持 DataTable/Dictionary 参数
- [New] CsvConfiguration 支持自订义 newLine & seperator & GetStreamReaderFunc
- [Optimization] 优化 SaveAs/Query excel 文件类型自动判断

### 0.10.3 
- [Bug] 修正 Query SharedStrings 控制字符没有 encoding (issue [Issue #149](https://github.com/shps951023/MiniExcel/issues/149))

### 0.10.2(请勿使用) 
- [Bug] 修正 SharedStrings get wrong index (issue [#153](https://github.com/shps951023/MiniExcel/issues/153))
- [Bug] SaveAs 支持 control character encoding (issue [Issue #149](https://github.com/shps951023/MiniExcel/issues/149))

### 0.10.1(请勿使用) 
- [New] SaveAs 支持 POCO excel 栏位名称/忽略 attribute

### 0.10.0(请勿使用) 
- [New] Query dynamic 表头预设自动忽略空白字串栏位
- [New] Query 强型别支持自订义 excel 栏位名称/忽略 attribute

### 0.9.1(请勿使用) 
- [Bug] 解决无法 mapping Cell Value 到 Nullable 属性类别 (issue #138)

### 0.9.0(请勿使用)
- [Bug] 解决 System.IO.Compression 引用两次问题  (issue #97)
- [Bug] 强型别 Query 空列会重複複製问题

### 0.8.0(请勿使用)
- [New] MiniExcel.Query 支持文件路径查询

### 0.7.0(请勿使用)
- 优化 SaveAs 效率
- [Breaking Change] SaveAs value 参数类别检查逻辑

### 0.6.0(请勿使用)
- [New] SaveAs 支持 参数 IEnumerable 延迟查询
- [Breaking Change] 移除 SaveAs by object, 现在只支持 Datatable,IEnumerable<T>,ICollection<T>
- [Bug] 修正空列生成 excel 错误 (issue: #128)

### 0.5.0(请勿使用)
- [New] 支持 OpenXml Xlsx SaveAs writer 模式避免 OOM
- [Breaking Change] 移除 SaveAs startCell 参数
- [Bug] 修正 SaveAs dimension printHeader:true 异常

### 0.4.0(请勿使用)
- [New] 支持 create CSV by 文件路径或是 stream 
- [New] 支持 csv 自订义 configuration 
- [New] 支持自动/手动指定 excel 类型 (xlsx or csv)
- [Breaking Changes] 移除 Query First/FirstOrDefault/Single/SingleOrDefault 方法, 使用者使用 LINQ 即可

### 0.3.0(请勿使用)
- [New] 支持 SaveAs by IEnumerable of DapperRow and IDictionary<string,object>
- [New] 支持 dynamic query timespan style 格式 mapping timespan 类别

### 0.2.3(请勿使用)
- [Bug] 修正內存洩漏问题
- [New] 支持 style datetime 格式 mapping datetime 类别.

### 0.2.2(请勿使用)
- SavaAs 支持 xl/sheet dimension
- [Breaking Changes] SaveAs value 类别准许 object & DataTable & ICollection
- [Bug] 修正 ICollection with type 没有数据错误 (https://github.com/shps951023/MiniExcel/issues/105)

### 0.2.1(请勿使用)  
- [Optimize] Optimize type mapping bool and datetime auto check
- [New] Query 支持 xl/worksheets/Sheet Xml `<c>` 没有 `r` 属性或是没有 `<dimension>` 但 `<c>` 有 `r` 属性情况, 但是效率会远低于有 dimension ([](https://github.com/shps951023/MiniExcel/issues/2))

### 0.2.0(请勿使用)  
- 发布至 nuget.org

### 0.1.0-preview(请勿使用) 
- [New] 添加 Query 强型别 mapping
- [New] 添加 QueryFirstOrDefault、QuerySingle、QuerySingleOrDefault

### 0.0.7-beta(请勿使用) 
- [New] 添加 QueryFirst 方法

### 0.0.6-beta(请勿使用) 
- [Breaking Changes] 替换 Create 名称为 SavaAs

### 0.0.5-beta(请勿使用) 
- [Bug] Release 删除 `assembly: InternalsVisibleTo` 依赖

### 0.0.4-beta(请勿使用) 
- [New] 支持 SaveAs Stream

### 0.0.3-beta(请勿使用) 
- [New] 支持 Query dynamic and IEnumrable 延迟查询避免 OOM
- [New] MiniExcelHelper.Create value 类别换成 ICollection
- [New] Encode XML 值 `&apos; &quot; &gt; &lt; &amp;`
- [New] 检查多 Sheet Index 排序
- [New] Dynamic Query 支持 A,B,C.. 栏位名称 key
- [New] 支持列与列之间空列情况

### 0.0.2-beta(请勿使用) 
- [New] 添加 MiniExcelHelper.Read 方法
- [Breaking Changes] 移除 System.IO.Packaging.Package 依赖, 换成 System.IO.Compression.ZipArchive

### 0.0.1-beta(请勿使用) 
- [New] 添加 MiniExcelHelper.Create 方法

### 0.0.0(请勿使用) 
- Init