// ReSharper disable once CheckNamespace
namespace MiniExcelLibs.OpenXml;

public sealed class OpenXmlConfiguration : MiniExcelLib.OpenXml.OpenXmlConfiguration, IConfiguration
{
    public OpenXmlConfiguration()
    {
        // Needs assignment inside the constructor so that it also assigns the base property through the forwarding setter,
        // so that the facade object and the original object are kept in sync from the start. 
        StyleOptions = new OpenXmlStyleOptions();
    }

    public new OpenXmlStyleOptions StyleOptions
    {
        get;
        set
        {
            field = value;
            base.StyleOptions = value;
        }
    }

    public new TableStyles TableStyles
    {
        get => base.TableStyles switch
        {
            MiniExcelLib.OpenXml.TableStyles.Default => TableStyles.Default,
            _ => TableStyles.None
        };
        set => base.TableStyles = value switch
        {
            TableStyles.Default => MiniExcelLib.OpenXml.TableStyles.Default,
            _ => MiniExcelLib.OpenXml.TableStyles.None
        };
    }
}

public sealed class OpenXmlStyleOptions : MiniExcelLib.OpenXml.Styles.OpenXmlStyleOptions
{
    public new OpenXmlHeaderStyle? HeaderStyle
    {
        get;
        set
        {
            field = value;
            base.HeaderStyle = value;
        }
    }

    public new HorizontalCellAlignment HorizontalAlignment
    {
        get => base.HorizontalAlignment switch
        {
            MiniExcelLib.Core.Enums.HorizontalCellAlignment.Right => HorizontalCellAlignment.Right,
            MiniExcelLib.Core.Enums.HorizontalCellAlignment.Center => HorizontalCellAlignment.Center,
            _ => HorizontalCellAlignment.Left
        };
        set => base.HorizontalAlignment = value switch
        {
            HorizontalCellAlignment.Right => MiniExcelLib.Core.Enums.HorizontalCellAlignment.Right,
            HorizontalCellAlignment.Center => MiniExcelLib.Core.Enums.HorizontalCellAlignment.Center,
            _ => MiniExcelLib.Core.Enums.HorizontalCellAlignment.Left
        };
    }

    public new VerticalCellAlignment VerticalAlignment
    {
        get => base.VerticalAlignment switch
        {
            MiniExcelLib.Core.Enums.VerticalCellAlignment.Top => VerticalCellAlignment.Top,
            MiniExcelLib.Core.Enums.VerticalCellAlignment.Center => VerticalCellAlignment.Center,
            _ => VerticalCellAlignment.Bottom
        };
        set => base.VerticalAlignment = value switch
        {
            VerticalCellAlignment.Top => MiniExcelLib.Core.Enums.VerticalCellAlignment.Top,
            VerticalCellAlignment.Center => MiniExcelLib.Core.Enums.VerticalCellAlignment.Center,
            _ => MiniExcelLib.Core.Enums.VerticalCellAlignment.Bottom
        };
    }
}

public sealed class OpenXmlHeaderStyle : MiniExcelLib.OpenXml.Styles.OpenXmlHeaderStyle
{
    public new HorizontalCellAlignment HorizontalAlignment
    {
        get => base.HorizontalAlignment switch
        {
            MiniExcelLib.Core.Enums.HorizontalCellAlignment.Right => HorizontalCellAlignment.Right,
            MiniExcelLib.Core.Enums.HorizontalCellAlignment.Center => HorizontalCellAlignment.Center,
            _ => HorizontalCellAlignment.Left
        };
        set => base.HorizontalAlignment = value switch
        {
            HorizontalCellAlignment.Right => MiniExcelLib.Core.Enums.HorizontalCellAlignment.Right,
            HorizontalCellAlignment.Center => MiniExcelLib.Core.Enums.HorizontalCellAlignment.Center,
            _ => MiniExcelLib.Core.Enums.HorizontalCellAlignment.Left
        };
    }

    public new VerticalCellAlignment VerticalAlignment
    {
        get => base.VerticalAlignment switch
        {
            MiniExcelLib.Core.Enums.VerticalCellAlignment.Top => VerticalCellAlignment.Top,
            MiniExcelLib.Core.Enums.VerticalCellAlignment.Center => VerticalCellAlignment.Center,
            _ => VerticalCellAlignment.Bottom
        };
        set => base.VerticalAlignment = value switch
        {
            VerticalCellAlignment.Top => MiniExcelLib.Core.Enums.VerticalCellAlignment.Top,
            VerticalCellAlignment.Center => MiniExcelLib.Core.Enums.VerticalCellAlignment.Center,
            _ => MiniExcelLib.Core.Enums.VerticalCellAlignment.Bottom
        };
    }
}

public enum TableStyles { None, Default }
public enum HorizontalCellAlignment { Left, Center, Right }
public enum VerticalCellAlignment { Bottom, Center, Top }
