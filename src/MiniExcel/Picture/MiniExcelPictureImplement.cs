using MiniExcelLibs.OpenXml;
using MiniExcelLibs.Zip;
using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;

namespace MiniExcelLibs.Picture
{
    internal static partial class MiniExcelPictureImplement
    {
        [Zomp.SyncMethodGenerator.CreateSyncVersion]
        public static async Task AddPictureAsync(Stream excelStream, CancellationToken cancellationToken = default, params MiniExcelPicture[] images)
        {
            // get sheets
            var excelArchive = new ExcelOpenXmlZip(excelStream);
            var reader = await ExcelOpenXmlSheetReader.CreateAsync(excelStream, null, cancellationToken: cancellationToken).ConfigureAwait(false);
            var sheetEntries = await reader.GetWorkbookRelsAsync(excelArchive.entries, cancellationToken).ConfigureAwait(false);

            var drawingRelId = $"rId{Guid.NewGuid():N}";
            var drawingId = Guid.NewGuid().ToString("N");
            var imageId = 2;
            using (var archive = new ZipArchive(excelStream, ZipArchiveMode.Update, true))
            {
                foreach (var image in images)
                {
                    var imageBytes = image.ImageBytes;
                    var sheetEnt = image?.SheetName == null 
                        ? sheetEntries[0] 
                        : sheetEntries.FirstOrDefault(x => x.Name == image.SheetName) ?? sheetEntries.First();

                    var sheetName = sheetEnt.Path.Split('/').Last().Split('.')[0];
                    var col = image.ColumnNumber;
                    var row = image.RowNumber;
                    var widthPx = image?.WidthPx;
                    var heightPx = image?.HeightPx;

                    // Step 1: Add image to /xl/media/
                    var imageName = $"image{Guid.NewGuid():N}.png";
                    var imagePath = $"xl/media/{imageName}";
                    var imageEntry = archive.CreateEntry(imagePath);
                    using (var entryStream = imageEntry.Open())
                    {
                        entryStream.Write(imageBytes, 0, imageBytes.Length);
                    }

                    // Step 2: Update [Content_Types].xml
                    var contentTypesEntry = archive.GetEntry("[Content_Types].xml");
                    var contentTypesDoc = LoadXml(contentTypesEntry);
                    if (!contentTypesDoc.DocumentElement.InnerXml.Contains("image/png"))
                    {
                        var defaultNode = contentTypesDoc.CreateElement("Default", contentTypesDoc.DocumentElement.NamespaceURI);
                        defaultNode.SetAttribute("Extension", "png");
                        defaultNode.SetAttribute("ContentType", "image/png");
                        contentTypesDoc.DocumentElement.AppendChild(defaultNode);
                    }

                    var overrideDrawingFileExists = contentTypesDoc.DocumentElement.ChildNodes
                        .Cast<XmlNode>()
                        .Any(node => node.Name == "Override" && node.Attributes?["PartName"].Value == $"/xl/drawings/drawing{drawingId}.xml");

                    if (!overrideDrawingFileExists)
                    {
                        var overrideNode = contentTypesDoc.CreateElement("Override", contentTypesDoc.DocumentElement.NamespaceURI);
                        overrideNode.SetAttribute("PartName", $"/xl/drawings/drawing{drawingId}.xml");
                        overrideNode.SetAttribute("ContentType", "application/vnd.openxmlformats-officedocument.drawing+xml");
                        contentTypesDoc.DocumentElement.AppendChild(overrideNode);
                    }
                    SaveXml(contentTypesDoc, contentTypesEntry);

                    // Step 3: Update xl/worksheets/sheetX.xml
                    var sheetPath = $"xl/worksheets/{sheetName}.xml";
                    var sheetEntry = archive.GetEntry(sheetPath);
                    var sheetDoc = LoadXml(sheetEntry);
                    var relId = $"rId{Guid.NewGuid():N}";
                    // unique relId for drawing

                    // existMiniExcelUniqueDrawingNode = check sheetDoc exist <drawing r:id="rId51b2a752f2454acfba519a539186a413"/> and check its attribut r:id = drawingRelId
                    var uniqueDrawingNode = sheetDoc.SelectSingleNode(
                        $"/x:worksheet/x:drawing[@r:id='{drawingRelId}']",
                        GetRNamespaceManager(sheetDoc));

                    if (uniqueDrawingNode != null)
                    {
                        var drawingNode = sheetDoc.CreateElement("drawing", sheetDoc.DocumentElement?.NamespaceURI);
                        drawingNode.Attributes
                            .Append(sheetDoc.CreateAttribute("r", "id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships"))
                            .Value = drawingRelId;
                        sheetDoc.DocumentElement.AppendChild(drawingNode);
                    }
                    SaveXml(sheetDoc, sheetEntry);

                    {
                        var drawingPath = $"xl/worksheets/_rels/{sheetName}.xml.rels";
                        var isExistEntry = false;
                        var sheetRelsEntry = archive.GetEntry(drawingPath);
                        if (sheetRelsEntry != null)
                        {
                            isExistEntry = true;
                        }
                        else
                        {
                            sheetRelsEntry = archive.CreateEntry(drawingPath);
                        }

                        if (isExistEntry)
                        {
                            var sheetRelsDoc = LoadXml(sheetRelsEntry);
                            var exists = CheckRelationshipExists(
                                sheetRelsDoc,
                                drawingRelId,
                                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing",
                                $"../drawings/drawing{drawingId}.xml"
                            );
                            if (!exists)
                            {
                                var relNode = sheetRelsDoc.CreateElement("Relationship", sheetRelsDoc.DocumentElement.NamespaceURI);
                                relNode.SetAttribute("Id", drawingRelId);
                                relNode.SetAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing");
                                relNode.SetAttribute("Target", $"../drawings/drawing{drawingId}.xml");
                                sheetRelsDoc.DocumentElement.AppendChild(relNode);
                            }
                            SaveXml(sheetRelsDoc, sheetRelsEntry);
                        }
                        else
                        {
                            var sheetRelsDoc = new XmlDocument();
                            sheetRelsDoc.LoadXml(@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships""/>");
                            var relNode = sheetRelsDoc.CreateElement("Relationship", sheetRelsDoc.DocumentElement.NamespaceURI);
                            relNode.SetAttribute("Id", drawingRelId);
                            relNode.SetAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing");
                            relNode.SetAttribute("Target", $"../drawings/drawing{drawingId}.xml");
                            sheetRelsDoc.DocumentElement.AppendChild(relNode);
                            SaveXml(sheetRelsDoc, sheetRelsEntry);
                        }
                        
                    }

                    // Step 4: Update exist xl/drawings/drawingX if not create one
                    {
                        XmlDocument drawingDoc;
                        var drawingPath = $"xl/drawings/drawing{drawingId}.xml";
                        var drawingEntry = archive.GetEntry(drawingPath);
                        if (drawingEntry != null)
                        {
                            drawingDoc = LoadXml(drawingEntry);
                            if (drawingDoc.DocumentElement != null)
                            {
                                // Create the new <xdr:twoCellAnchor> node
                                var newTwoCellAnchor = drawingDoc.CreateElement("xdr", "twoCellAnchor", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing");
                                newTwoCellAnchor.SetAttribute("editAs", "oneCell");

                                // Add the <xdr:from> node
                                var fromNode = drawingDoc.CreateElement("xdr", "from", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing");
                                fromNode.InnerXml = $@"
        <xdr:col xmlns:xdr=""http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"">{col}</xdr:col>
        <xdr:colOff xmlns:xdr=""http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"">0</xdr:colOff>
        <xdr:row xmlns:xdr=""http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"">{row}</xdr:row>
        <xdr:rowOff xmlns:xdr=""http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"">0</xdr:rowOff>";
                                newTwoCellAnchor.AppendChild(fromNode);

                                // Add the <xdr:to> node
                                var toNode = drawingDoc.CreateElement("xdr", "to", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing");
                                toNode.InnerXml = $@"
        <xdr:col xmlns:xdr=""http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"">{col + 1}</xdr:col>
        <xdr:colOff xmlns:xdr=""http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"">{widthPx * 9525}</xdr:colOff>
        <xdr:row xmlns:xdr=""http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"">{row + 1}</xdr:row>
        <xdr:rowOff xmlns:xdr=""http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"">{heightPx * 9525}</xdr:rowOff>";
                                newTwoCellAnchor.AppendChild(toNode);

                                // Add the <xdr:pic> node
                                var picNode = drawingDoc.CreateElement("xdr", "pic", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing");
                                picNode.InnerXml = $@"
        <xdr:nvPicPr xmlns:xdr=""http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"">
            <xdr:cNvPr id=""{++imageId}"" name=""Picture {relId}"" xmlns:xdr=""http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing""/>
            <xdr:cNvPicPr xmlns:xdr=""http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"">
                <a:picLocks noChangeAspect=""1"" xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main""/>
            </xdr:cNvPicPr>
        </xdr:nvPicPr>
        <xdr:blipFill xmlns:xdr=""http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"">
            <a:blip r:embed=""{relId}"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"" xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main""/>
            <a:stretch xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main"">
                <a:fillRect xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main""/>
            </a:stretch>
        </xdr:blipFill>
        <xdr:spPr xmlns:xdr=""http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"">
            <a:xfrm xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main"">
                <a:off x=""0"" y=""0"" xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main""/>
                <a:ext {(widthPx==null?"":$@"cx = ""{widthPx * 9525}""")} {(heightPx == null ? "" : $@"cy=""{heightPx * 9525}""")} xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main""/>
            </a:xfrm>
            <a:prstGeom prst=""rect"" xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main"">
                <a:avLst xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main""/>
            </a:prstGeom>
        </xdr:spPr>";
                                newTwoCellAnchor.AppendChild(picNode);

                                // Add the <xdr:clientData> node
                                var clientDataNode = drawingDoc.CreateElement("xdr", "clientData", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing");
                                newTwoCellAnchor.AppendChild(clientDataNode);

                                // Append the new node after the last <xdr:twoCellAnchor>
                                drawingDoc.DocumentElement.AppendChild(newTwoCellAnchor);
                            }

                        }
                        else
                        {
                            drawingEntry = archive.CreateEntry(drawingPath);
                            drawingDoc = CreateDrawingXml(col, row, widthPx, heightPx, relId);
                        }

                        SaveXml(drawingDoc, drawingEntry);
                    }

                    // Step 5: Create or update xl/drawings/_rels/drawingX.xml.rels
                    {
                        var drawingRelsPath = $"xl/drawings/_rels/drawing{drawingId}.xml.rels";
                        var drawingRelsEntry = archive.GetEntry(drawingRelsPath) ?? archive.CreateEntry(drawingRelsPath);
                        var drawingRelsDoc = LoadXml(drawingRelsEntry);

                        // Check if the relationship already exists
                        var existingRel = drawingRelsDoc.SelectSingleNode($"/x:Relationships/x:Relationship[@Id='{relId}']", GetNamespaceManager(drawingRelsDoc)); //todo: why never used?
                        var relNode = drawingRelsDoc.CreateElement("Relationship", drawingRelsDoc.DocumentElement.NamespaceURI);
                        relNode.SetAttribute("Id", relId);
                        relNode.SetAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image");
                        relNode.SetAttribute("Target", $"../media/{imageName}");
                        drawingRelsDoc.DocumentElement.AppendChild(relNode);

                        SaveXml(drawingRelsDoc, drawingRelsEntry);
                    }
                }
            }
        }

        private static XmlDocument LoadXml(ZipArchiveEntry entry)
        {
            var doc = new XmlDocument();
            if (entry == null)
            {
                doc.LoadXml(@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships""/>");
                return doc;
            }
            
            using (var stream = entry.Open())
            using (var reader = new StreamReader(stream))
            {
                var streamString = reader.ReadToEnd();
                if (string.IsNullOrEmpty(streamString))
                {
                    doc.LoadXml(@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships""/>");
                    return doc;
                }
                stream.Position = 0;
                doc.Load(stream);
            }
            
            return doc;
        }

        private static void SaveXml(XmlDocument doc, ZipArchiveEntry entry)
        {
            using (var stream = entry.Open())
            {
                stream.SetLength(0);
                doc.Save(stream);
            }
        }

        private static XmlNamespaceManager GetNamespaceManager(XmlDocument doc)
        {
            var nsmgr = new XmlNamespaceManager(doc.NameTable);
            nsmgr.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
            return nsmgr;
        }

        private static XmlDocument CreateDrawingXml(int col, int row, int? widthPx, int? heightPx, string relId)
        {
            var doc = new XmlDocument();
            doc.LoadXml($@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<xdr:wsDr xmlns:xdr=""http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"" xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main"">
    <xdr:twoCellAnchor editAs=""oneCell"">
        <xdr:from><xdr:col>{col}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{row}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
        <xdr:to><xdr:col>{col + 1}</xdr:col><xdr:colOff>{widthPx * 9525}</xdr:colOff><xdr:row>{row + 1}</xdr:row><xdr:rowOff>{heightPx * 9525}</xdr:rowOff></xdr:to>
        <xdr:pic>
            <xdr:nvPicPr>
                <xdr:cNvPr id=""2"" name=""Picture {relId}""/>
                <xdr:cNvPicPr><a:picLocks noChangeAspect=""1""/></xdr:cNvPicPr>
            </xdr:nvPicPr>
            <xdr:blipFill>
                <a:blip r:embed=""{relId}"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships""/>
                <a:stretch><a:fillRect/></a:stretch>
            </xdr:blipFill>
            <xdr:spPr><a:xfrm><a:off x=""0"" y=""0""/><a:ext {(widthPx == null ? "" : $@"cx = ""{widthPx * 9525}""")} {(heightPx == null ? "" : $@"cy=""{heightPx * 9525}""")}/></a:xfrm><a:prstGeom prst=""rect""><a:avLst/></a:prstGeom></xdr:spPr>
        </xdr:pic>
        <xdr:clientData/>
    </xdr:twoCellAnchor>
</xdr:wsDr>");
            return doc;
        }

        private static XmlNamespaceManager GetRNamespaceManager(XmlDocument doc)
        {
            var nsmgr = new XmlNamespaceManager(doc.NameTable);
            nsmgr.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
            nsmgr.AddNamespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            return nsmgr;
        }

        private static bool CheckRelationshipExists(XmlDocument doc, string id, string type, string target)
        {
            var namespaceManager = new XmlNamespaceManager(doc.NameTable);
            namespaceManager.AddNamespace("x", "http://schemas.openxmlformats.org/package/2006/relationships");

            var xpath = $"/x:Relationships/x:Relationship[@Id='{id}' and @Type='{type}' and @Target='{target}']";
            var node = doc.SelectSingleNode(xpath, namespaceManager);

            return node != null;
        }
    }
}