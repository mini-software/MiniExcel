using MiniExcelLib.Core.Helpers;
using MiniExcelLib.Core.OpenXml.Models;
using MiniExcelLib.Core.OpenXml.Utils;

namespace MiniExcelLib.Core.OpenXml.Constants;

internal static class ExcelXml
{
    static ExcelXml()
    {
        DefaultRels = XmlHelper.MinifyXml(DefaultRels);
        DefaultWorkbookXml = XmlHelper.MinifyXml(DefaultWorkbookXml);
        DefaultWorkbookXmlRels = XmlHelper.MinifyXml(DefaultWorkbookXmlRels);
        DefaultSheetRelXml = XmlHelper.MinifyXml(DefaultSheetRelXml);
        DefaultDrawing = XmlHelper.MinifyXml(DefaultDrawing);
    }

    internal const string EmptySheetXml = """<?xml version="1.0" encoding="utf-8"?><x:worksheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><x:dimension ref="A1"/><x:sheetData></x:sheetData></x:worksheet>""";

    internal static readonly string DefaultRels = 
        """
       <?xml version="1.0" encoding="utf-8"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
            <Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml" Id="Rfc2254092b6248a9" />
        </Relationships>
       """;

    internal static readonly string DefaultWorkbookXmlRels = 
        """
        <?xml version="1.0" encoding="utf-8"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
            {{sheets}}
            <Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="/xl/styles.xml" Id="R3db9602ace774fdb" />
            <Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="/xl/sharedStrings.xml" Id="R3db9602ace778fdb" />
        </Relationships>
        """;

    internal static readonly string DefaultWorkbookXml = 
        """
        <?xml version="1.0" encoding="utf-8"?>
        <x:workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <x:sheets>
                {{sheets}}
            </x:sheets>
        </x:workbook>
        """;

    internal static readonly string DefaultSheetRelXml = 
        """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
            {{format}}
        </Relationships>
        """;
        
    internal static readonly string DefaultDrawing = 
        """
        <?xml version="1.0" encoding="utf-8" standalone="yes"?>
        <xdr:wsDr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing">
            {{format}}
        </xdr:wsDr>
        """;
        
    internal const string DefaultDrawingXmlRels = 
        """
        <?xml version="1.0" encoding="utf-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
            {{format}}
        </Relationships>
        """;

    internal const string DefaultSharedString = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\" ?><sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"0\" uniqueCount=\"0\"></sst>";

    internal const string StartTypes = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings" Extension="bin"/><Default ContentType="application/xml" Extension="xml"/><Default ContentType="image/jpeg" Extension="jpg"/><Default ContentType="image/png" Extension="png"/><Default ContentType="image/gif" Extension="gif"/><Default ContentType="application/vnd.openxmlformats-package.relationships+xml" Extension="rels"/>""";
    internal static string ContentType(string contentType, string partName) => $"<Override ContentType=\"{contentType}\" PartName=\"/{partName}\" />";
    internal const string EndTypes = "</Types>";

    internal static string WorksheetRelationship(SheetDto sheetDto)
        => $"""<Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="/{sheetDto.Path}" Id="{sheetDto.ID}" />""";

    internal static string ImageRelationship(FileDto image)
        => $"""<Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="{image.Path2}" Id="{image.ID}" />""";

    internal static string DrawingRelationship(int sheetId)
        => $"""<Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing{sheetId}.xml" Id="drawing{sheetId}" />""";

    internal static string DrawingXml(FileDto file, int fileIndex)
        => $"""
            <xdr:oneCellAnchor>
                    <xdr:from>
                        <xdr:col>{file.CellIndex - 1 /* why -1 : https://user-images.githubusercontent.com/12729184/150460189-f08ed939-44d4-44e1-be6e-9c533ece6be8.png*/}</xdr:col>
                        <xdr:colOff>0</xdr:colOff>
                        <xdr:row>{file.RowIndex - 1}</xdr:row>
                        <xdr:rowOff>0</xdr:rowOff>
                    </xdr:from>
                    <xdr:ext cx="609600" cy="190500" />
                    <xdr:pic>
                        <xdr:nvPicPr>
                            <xdr:cNvPr id="{fileIndex + 1}" descr="" name="2a3f9147-58ea-4a79-87da-7d6114c4877b" />
                            <xdr:cNvPicPr>
                                <a:picLocks noChangeAspect="1" />
                            </xdr:cNvPicPr>
                        </xdr:nvPicPr>
                        <xdr:blipFill>
                            <a:blip r:embed="{file.ID}" cstate="print" />
                            <a:stretch>
                                <a:fillRect />
                            </a:stretch>
                        </xdr:blipFill>
                        <xdr:spPr>
                            <a:xfrm>
                                <a:off x="0" y="0" />
                                <a:ext cx="0" cy="0" />
                            </a:xfrm>
                            <a:prstGeom prst="rect">
                                <a:avLst />
                            </a:prstGeom>
                        </xdr:spPr>
                    </xdr:pic>
                    <xdr:clientData />
                </xdr:oneCellAnchor>
            """;

    internal static string Sheet(SheetDto sheetDto, int sheetId)
        => $"""<x:sheet name="{XmlHelper.EncodeXml(sheetDto.Name)}" sheetId="{sheetId}"{(string.IsNullOrWhiteSpace(sheetDto.State) ? string.Empty : $" state=\"{sheetDto.State}\"")} r:id="{sheetDto.ID}" />""";

}