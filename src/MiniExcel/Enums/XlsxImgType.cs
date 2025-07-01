using System;

namespace MiniExcelLibs.Enums;

/// <summary>
/// Excel 图片展示方式（是否随单元格对齐/缩放）。
/// </summary>
public enum XlsxImgType
{
	/// <summary>
	/// 图片随单元格移动但不缩放（OneCellAnchor）。
	/// 通常用于图片只绑定一个起点单元格。
	/// </summary>
	OneCellAnchor,
	/// <summary>
	/// 图片浮动在表格上，固定位置不随单元格变化（AbsoluteAnchor）。
	/// </summary>
	AbsoluteAnchor,
	/// <summary>
	/// 图片嵌入单元格中，随单元格移动和缩放（TwoCellAnchor）。
	/// </summary>
	TwoCellAnchor,
	
}
