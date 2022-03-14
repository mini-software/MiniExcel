
## 更新日志

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
<p><strong><a href="README.md">English</a> | <a href="README.zh-CN.md">简体中文</a> | <a href="README.zh-Hant.md">繁體中文</a></strong></p>
</div>

---

<div align="center">
<p> 您的 <a href="https://github.com/shps951023/miniexcel">Star</a> 和 <a href="https://miniexcel.github.io">赞助</a> 能帮助 MiniExcel 成长 </p>
</div>

---

<div align="center">
<p> 您的 <a href="https://github.com/shps951023/miniexcel">Star</a> 和 <a href="https://miniexcel.github.io">赞助</a> 能帮助 MiniExcel 成长 </p>
</div>


---

### 1.24.1
- [Bug] 修正 QueryAsync configulation 没有效果问题 #338 
- [Bug] 修正 QueryAsync 无法使用 dynamic 类别

### 1.24.0
- [New] Query 支持 strict open xml, 感谢 [Weilence (Lowell)](https://github.com/Weilence)  #335 
- [New] SaveAs 以自订的 CultureInfo 转换 Format cell 值, 感谢[0xced (Cédric Luthi)](https://github.com/0xced) #333
- [New] SaveAsByTemplate 预设忽略 template 缺少参数 key 错误, OpenXmlConfiguration.IgnoreTemplateParameterMissing 可以开关此卡控. #I4WXFB
- [New] SaveAsByTemplate 当参数集合为空时会清空模版字串. #I4WM67

### 1.23.3
- [Bug] SaveAs CSV 当 value 为 DataTable 时，Key包含双引号Column Name不会显示`"`。 #I4WDA9

### 1.23.2
- [New] 支持 System.ComponentModel.DisplayName 的 `[DisplayName]` 作为excel标题 [#I4TXGT](https://gitee.com/dotnetchina/MiniExcel/issues/I4TXGT)
- [Bug] 修正  `Czech` 等国家CultureInfo会生成错误 decimal 数字提示 #331

### 1.23.0

- [New] 新增 `GetReader` 方法 #328 #290  (感谢 [杨福来 Yang](https://github.com/yfl8910) )

### 1.22.0
- [New] SaveAs 支持自定义 CultureInfo #316
- [New] Query 支持自定义 CultureInfo #316
- [New] 新 byte array 转换器 #327
- [Breaking Change] 移除 Byte Array 与 base64 转换器
- [Breaking Change] `EnableConvertByteArray` 取代 `ConvertByteArrayToBase64String` 

### 0.21.5
- [Bug] 修正 SaveAs multiple sheet value error "Excel completed file level validation and repair. Some parts of this workbook may have been repaired or discarded."  #325

### 0.21.4
- [New] Remove LISENCE_CODE check

### 0.21.1
- [New] Check License Code

### 0.21.0
- [New] ExcelFormat 支持 DateTimeOffset/Decimal/double 等类别 format #I49RZH #312 #305
- [New] 支持byte文件导入/导出
- [New] SaveAs 支持预设转换byte[] 值为 base64 字串
- [New] Query 支持转换 base64 字串值为 bytep[]
- [New] OpenXmlConfiguration 增加 `ConvertByteArrayToBase64String` 属性来开关 base64 转换器
- [New] Query 支持 ExcelInvalidCastException 储存行、列、值数据 #309

### 0.20.0 
- [New] SaveAs 支持图片生成 #304
- [Opt] 提高 SaveAs 执行效率

### 0.19.3-beta
- [Fix]  Excelnumberformat 1.1.0 凭证过期 (Valid from: 2018-04-10 08:00:00 to 2021-04-14 20:00:00) [link](https://github.com/andersnm/ExcelNumberFormat/issues/34)

### 0.19.2
- [New] SaveAsByTemplate 支持 datareader [#I4HL54](https://gitee.com/dotnetchina/MiniExcel/issues/I4HL54)

### 0.19.1
- [New] QueryAsDataTable 删除空白 Column keys. #298
- [Bug] Error NU3037: ExcelNumberFormat 1.1.0 #302
- [Bug] SaveAs 前缀和后缀空格会丢失 #294

### 0.19.0
- [New] SaveAs 预设样式增加筛选功能. #190
- [New] 新增 ConvertCsvToXlsx、ConvertXlsxToCsv 方法. #292
- [New] OpenXmlConfiguration 新增 AutoFilter 属性. 
- [Bug] 修正 CSV 读取后 SaveAs 会抛出 "Stream was not readable." 错误. #293
- [Bug] 修正 SaveAsByTemplate & 被转成 &amp; [I4DQUN](https://gitee.com/dotnetchina/MiniExcel/issues/I4DQUN)


### 0.18.0
- [New] SaveAs 支持 enum description #I49RYZ
- [New] Query 强型别支持多列名对应同一属性 [#I40QA5](https://gitee.com/dotnetchina/MiniExcel/issues/I40QA5)
- [Breaking Change] SaveAs 传空`IEnumerable<强型别>`现在会生成表头. #133
- [Bug] SaveAs sheet enum 映射 cell 错误 number 型别 #286

### 0.17.5
- [Bug] 修复xlsx文件中标题字段包含`&,<,>,",'`符号，导出后打不开文件

### 0.17.4 
- [Bug] 修复0.17.3版本更新导致 SaveAs 创建 xlsx 文件错误 "XML error : Catastrophic failure"

### 0.17.3
- [New] 支持设定列宽 #280
- [Bug] 修复 csv 不支持 QueryAsDataTable #279
- [OPT] 更加清楚的错误信息，当文件是错误 excel zip 格式 #272

### 0.17.2
- [Bug] 修复 v0.16.0-0.17.1 自定义格式含有特定格式(e.g:`#,##0.000_);[Red]\(#,##0.000\)`)，自动转换器会将 double 被转成 datetime 异常 #267

### 0.17.1
- [New] 增加 QueryAsDataTableAsync(this Stream stream..)
- [OPT] 强型别 Query 转型错误信息能知道在哪一行列出错 [#I3X2ZL](https://gitee.com/dotnetchina/MiniExcel/issues/I3X2ZL)

### 0.17.0
- [New] 支持 Async/Task #52, contributor: [isdaniel ( SHIH,BING-SIOU)](https://github.com/isdaniel)

### 0.16.1
- [New] SaveAsByTemplate 支持 DateTime 自定义格式 #255, contributor: [网虫 (landde) - Gitee.com](https://gitee.com/landde)

### 0.16.0
- [New] Query 支持自定义日期格式转成 datetime 型别 #256
- [Bug] 修正 Query 重复呼叫 convertValueByStyleFormat 方法造成资源浪费 #259

### 0.15.5
- [Bug] 特定中文环境日期格式转换InvalidCastException错误 #257

### 0.15.4
- [Breaking Change] CSV Reader/Writer 预设编码改变 : UTF-8 => UTF-8-BOM
- [Breaking Change] 重新命名 CsvConfiguration GetStreamReaderFunc => StreamReaderFunc
- [New] Csv SaveAs 支持自定义 StreamWriter

### 0.15.3
- [New] Csv SaveAs 支持 datareader

### 0.15.2
- [New] 支持自定义日期时间格式 #241
- [Bug] CSV类型映射查询错误 "cannot be converted to xxx type" #243
- [Bug] Stream 读取 xls 文件时没有错误异常抛出 #242
- [Breaking Change] 流无法识别文件类型，请手动指定ExcelType

### 0.15.1
- [Bug] 修正 Sheetxml 结尾包含两个 ">" 导致解析错误 #240

### 0.15.0
- [New] SaveAs 更改预设样式、并提供样式选择 #132
- [New] SaveAs 支持 DataSet #235

### 0.14.8 
- [Bug] 修正 csv Query 内文包含逗号造成异常 #237 #I3R95M
- [Bug] 修正 QueryAsDataTable 类别检查异常，如 A2=5.5 , A3=0.55/1.1 系统会显示 double type check error #233

### 0.14.7
- [New] SaveAs 支持建立多工作表
- [Breaking Change] 更换 GetSheetNames 返回类型 IEnumerable<string> -> List<string>


### 0.14.6
- [Bug] 修正 SaveAs by datareader 错误 "Invalid attempt to call FieldCount when reader is closed" #230

### 0.14.5
- [Breaking Change] 更换 OpenXmlConfiguration FillMergedCells 名称

### 0.14.4
- [New] Query 支持向下填充合并的单元格 [#122](https://github.com/shps951023/MiniExcel/issues/122)
- [Bug] 修正 QueryAsDataTable 错误 "Cannot set Column to be null" #229

### 0.14.3
- [Opt] 支持 Xlsm 自动判断 #227
- [Bug] 修正 SaveAsByTemplate 单列 demension 索引错误 [#226](https://github.com/shps951023/MiniExcel/issues/226)

### 0.14.2
- [Bug] 修正 asp.net webform gridview datasource 不能使用 QueryAsDataTable [#223](https://github.com/shps951023/MiniExcel/issues/223)

### 0.14.1
- [Bug] 修正自定义 m/d 格式没转成 datetime [#222](https://github.com/shps951023/MiniExcel/issues/222)

### 0.14.0
- [New] Query、GetColumns 支持 startCell 能指定 cell 开始读取数据 [#147](https://github.com/shps951023/MiniExcel/issues/147)
- [New] GetColumns 支持读取表头

### 0.13.5
- [New] 新增 QueryAsDataTable 方法 [#216](https://github.com/shps951023/MiniExcel/issues/216)
- [New] SaveAs 支持 IDataReader value 参数 [#211](https://github.com/shps951023/MiniExcel/issues/211)
- [Bug] 修正数字格式的字串会被强制转换为decimal类型 [#I3OSKV](https://gitee.com/dotnetchina/MiniExcel/issues/I3OSKV)
- [Opt] 优化 SaveAs 类别转换算法，避免效率浪费

### 0.13.4
- [Changed] DataTable 以 Caption 优先当栏位名称 #217
- [New] Query 支持 Enum mapping #89
- [Opt] 优化 stream excel 类别(xlsx or csv)检查 #215

### 0.13.3
- [New] 支持 Excel 单纯读取模式，避免同时改模版又运行 MiniExcel 出现错误 "The process cannot access the file because it is being used by another process" [#87](https://github.com/shps951023/MiniExcel/issues/87)
- [Breaking Change] CSV SaveAs datetime 预设格式改为 "yyyy-MM-dd HH:mm:ss"
- [Bug] 修正模版模式集合渲染遇到合并列会出现异常问题 [#207](https://github.com/shps951023/MiniExcel/issues/207)
- [Bug] 修正 MiniExcel.SaveAs(path, value,sheetName:"Name"), 实际 sheetName 是 Sheet1

### 0.13.2
- [Bug] 超过 255 列无法读取错误 [#208](https://github.com/shps951023/MiniExcel/issues/208)

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