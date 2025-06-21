﻿namespace MiniExcelLibs.Picture
{
    using MiniExcelLibs.OpenXml;
    using MiniExcelLibs.Zip;
    using System;
    using System.IO;
    using System.IO.Compression;
    using System.Linq;
    using System.Xml;

    internal static class MiniExcelPictureImplement
    {
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
        [Zomp.SyncMethodGenerator.CreateSyncVersion]
        public static async Task AddPictureAsync(Stream excelStream, CancellationToken cancellationToken = default, params MiniExcelPicture[] images)
        {
            // get sheets
            using var excelArchive = new ExcelOpenXmlZip(excelStream);
            using var reader = await ExcelOpenXmlSheetReader.CreateAsync(excelStream, null, cancellationToken: cancellationToken).ConfigureAwait(false);
            var sheetEntries = await reader.GetWorkbookRelsAsync(excelArchive.entries, cancellationToken).ConfigureAwait(false);

            using (var archive = new ZipArchive(excelStream, ZipArchiveMode.Update, true))
            {
                // Group images by sheet
                var imagesBySheet = images.GroupBy(img => img.SheetName ?? sheetEntries.First().Name);
                foreach (var sheetGroup in imagesBySheet)
                {
                    var sheetName = sheetGroup.Key;
                    var sheetEnt = sheetEntries.FirstOrDefault(x => x.Name == sheetName) ?? sheetEntries.First();
                    var sheetXmlName = sheetEnt.Path.Split('/').Last().Split('.')[0];
                    string sheetPath = $"xl/worksheets/{sheetXmlName}.xml";

                    var sheetEntry = archive.GetEntry(sheetPath);
                    var sheetDoc = LoadXml(sheetEntry);

                    // Check for existing <drawing> node
                    var nsmgr = GetRNamespaceManager(sheetDoc);
                    var drawingNode = sheetDoc.SelectSingleNode("/x:worksheet/x:drawing", nsmgr) as XmlElement;
                    string drawingRelId;
                    string drawingId;
                    if (drawingNode != null)
                    {
                        // Drawing exists, get r:id
                        drawingRelId = drawingNode.GetAttribute("id", nsmgr.LookupNamespace("r"));
                        // Find the drawing target from .rels
                        string relsPath = $"xl/worksheets/_rels/{sheetXmlName}.xml.rels";
                        var relsEntry = archive.GetEntry(relsPath);
                        var relsDoc = LoadXml(relsEntry);
                        var namespaceManager = new XmlNamespaceManager(relsDoc.NameTable);
                        namespaceManager.AddNamespace("x", "http://schemas.openxmlformats.org/package/2006/relationships");
                        var xpath = $"/x:Relationships/x:Relationship[@Id='{drawingRelId}']";
                        var relNode = relsDoc.SelectSingleNode(xpath, namespaceManager);
                        string drawingTarget = relNode?.Attributes["Target"]?.Value;
                        drawingId = drawingTarget != null ? drawingTarget.Split('/').Last().Replace("drawing", "").Replace(".xml", "") : Guid.NewGuid().ToString("N");
                    }
                    else
                    {
                        // No drawing, create new
                        drawingRelId = $"rId{Guid.NewGuid().ToString("N")}";
                        drawingId = Guid.NewGuid().ToString("N");
                        // Add <drawing> node
                        var newDrawingNode = sheetDoc.CreateElement("drawing", sheetDoc.DocumentElement.NamespaceURI);
                        newDrawingNode.SetAttribute("id", nsmgr.LookupNamespace("r"), drawingRelId);
                        sheetDoc.DocumentElement.AppendChild(newDrawingNode);
                        SaveXml(sheetDoc, sheetEntry);
                        // Add relationship
                        string relsPath = $"xl/worksheets/_rels/{sheetXmlName}.xml.rels";
                        var relsEntry = archive.GetEntry(relsPath) ?? archive.CreateEntry(relsPath);
                        var relsDoc = LoadXml(relsEntry);
                        var relNode = relsDoc.CreateElement("Relationship", relsDoc.DocumentElement.NamespaceURI);
                        relNode.SetAttribute("Id", drawingRelId);
                        relNode.SetAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing");
                        relNode.SetAttribute("Target", $"../drawings/drawing{drawingId}.xml");
                        relsDoc.DocumentElement.AppendChild(relNode);
                        SaveXml(relsDoc, relsEntry);
                        // Update [Content_Types].xml for drawing
                        var contentTypesEntry = archive.GetEntry("[Content_Types].xml");
                        var contentTypesDoc = LoadXml(contentTypesEntry);
                        var overrideDrawingFileExists = false;
                        foreach (XmlNode node in contentTypesDoc.DocumentElement.ChildNodes)
                        {
                            if (node.Name == "Override" && node.Attributes["PartName"].Value == $"/xl/drawings/drawing{drawingId}.xml")
                            {
                                overrideDrawingFileExists = true;
                                break;
                            }
                        }
                        if (!overrideDrawingFileExists)
                        {
                            var overrideNode = contentTypesDoc.CreateElement("Override", contentTypesDoc.DocumentElement.NamespaceURI);
                            overrideNode.SetAttribute("PartName", $"/xl/drawings/drawing{drawingId}.xml");
                            overrideNode.SetAttribute("ContentType", "application/vnd.openxmlformats-officedocument.drawing+xml");
                            contentTypesDoc.DocumentElement.AppendChild(overrideNode);
                        }
                        SaveXml(contentTypesDoc, contentTypesEntry);
                    }


                    // Load or create drawing XML
                    string drawingPath = $"xl/drawings/drawing{drawingId}.xml";
                    var drawingEntry = archive.GetEntry(drawingPath) ?? archive.CreateEntry(drawingPath);
                    XmlDocument drawingDoc = LoadXml(drawingEntry);

                    // Load or create drawing rels
                    string drawingRelsPath = $"xl/drawings/_rels/drawing{drawingId}.xml.rels";
                    var drawingRelsEntry = archive.GetEntry(drawingRelsPath) ?? archive.CreateEntry(drawingRelsPath);
                    var drawingRelsDoc = LoadXml(drawingRelsEntry);

                    // Add each image to drawing and rels
                    foreach (var image in sheetGroup)
                    {
                        var imageBytes = image.ImageBytes;
                        var col = image.ColumnNumber;
                        var row = image.RowNumber;
                        var widthPx = image.WidthPx;
                        var heightPx = image.HeightPx;
                        // Step 1: Add image to /xl/media/
                        string imageName = $"image{Guid.NewGuid().ToString("N")}.png";
                        string imagePath = $"xl/media/{imageName}";
                        var imageEntry = archive.CreateEntry(imagePath);
                        using (var entryStream = imageEntry.Open())
                            entryStream.Write(imageBytes, 0, imageBytes.Length);
                        // Step 2: Update [Content_Types].xml for image
                        var contentTypesEntry = archive.GetEntry("[Content_Types].xml");
                        var contentTypesDoc = LoadXml(contentTypesEntry);
                        if (!contentTypesDoc.DocumentElement.InnerXml.Contains("image/png"))
                        {
                            var defaultNode = contentTypesDoc.CreateElement("Default", contentTypesDoc.DocumentElement.NamespaceURI);
                            defaultNode.SetAttribute("Extension", "png");
                            defaultNode.SetAttribute("ContentType", "image/png");
                            contentTypesDoc.DocumentElement.AppendChild(defaultNode);
                            SaveXml(contentTypesDoc, contentTypesEntry);
                        }
                        // Step 3: Add anchor to drawing XML
                        var relId = $"rId{Guid.NewGuid().ToString("N")}";
                        drawingDoc = CreateDrawingXml(drawingDoc, col, row, widthPx, heightPx, relId);
                        // Step 4: Add image relationship to drawing rels
                        var relNode = drawingRelsDoc.CreateElement("Relationship", drawingRelsDoc.DocumentElement.NamespaceURI);
                        relNode.SetAttribute("Id", relId);
                        relNode.SetAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image");
                        relNode.SetAttribute("Target", $"../media/{imageName}");
                        drawingRelsDoc.DocumentElement.AppendChild(relNode);
                    }
                    SaveXml(drawingDoc, drawingEntry);
                    SaveXml(drawingRelsDoc, drawingRelsEntry);
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

        private static XmlDocument CreateDrawingXml(XmlDocument existingDoc, int col, int row, int widthPx, int heightPx, string relId)
        {
            return DrawingXmlHelper.CreateOrUpdateDrawingXml(existingDoc, col, row, widthPx, heightPx, relId);
        }


        public class DrawingXmlHelper
        {
            private const string XdrNamespace = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
            private const string ANamespace = "http://schemas.openxmlformats.org/drawingml/2006/main";

            private static long PixelsToEMU(int pixels) => (long)(pixels * 9525);

            private static string GetColumnName(int colIndex)
            {
                string columnName = "";
                int dividend = colIndex + 1;
                while (dividend > 0)
                {
                    int modulo = (dividend - 1) % 26;
                    columnName = Convert.ToChar('A' + modulo).ToString() + columnName;
                    dividend = (dividend - modulo) / 26;
                }
                return columnName;
            }

            public static XmlDocument CreateOrUpdateDrawingXml(
    XmlDocument existingDoc,
    int col, int row,
    int widthPx, int heightPx,
    string relId)
            {
                var doc = existingDoc ?? new XmlDocument();
                var ns = new XmlNamespaceManager(doc.NameTable);
                ns.AddNamespace("xdr", XdrNamespace);
                ns.AddNamespace("a", ANamespace);
                ns.AddNamespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");


                // check or create <xdr:wsDr>
                XmlElement wsDr;
                if (existingDoc == null)
                {
                    wsDr = doc.CreateElement("xdr", "wsDr", XdrNamespace);
                    wsDr.SetAttribute("xmlns:xdr", XdrNamespace);
                    wsDr.SetAttribute("xmlns:a", ANamespace);
                    doc.AppendChild(wsDr);
                }
                else
                {
                    wsDr = doc.DocumentElement;
                }

                // get current oneCellAnchor to get id 和 name
                XmlNodeList anchors = wsDr.SelectNodes("//xdr:oneCellAnchor", ns);
                int imageCount = anchors?.Count ?? 0;

                // next ID（2）
                int nextId = imageCount + 2;

                // create oneCellAnchor
                var oneCellAnchor = doc.CreateElement("xdr", "oneCellAnchor", XdrNamespace);

                // <xdr:from>
                var from = doc.CreateElement("xdr", "from", XdrNamespace);
                AppendXmlElement(doc, from, "xdr", "col", col.ToString());
                AppendXmlElement(doc, from, "xdr", "colOff", "0");
                AppendXmlElement(doc, from, "xdr", "row", row.ToString());
                AppendXmlElement(doc, from, "xdr", "rowOff", "0");

                // <xdr:ext>
                var ext = doc.CreateElement("xdr", "ext", XdrNamespace);
                ext.SetAttribute("cx", PixelsToEMU(widthPx).ToString());
                ext.SetAttribute("cy", PixelsToEMU(heightPx).ToString());

                // <xdr:pic>
                var pic = doc.CreateElement("xdr", "pic", XdrNamespace);

                // <xdr:nvPicPr>
                var nvPicPr = doc.CreateElement("xdr", "nvPicPr", XdrNamespace);
                var cNvPr = doc.CreateElement("xdr", "cNvPr", XdrNamespace);
                cNvPr.SetAttribute("id", nextId.ToString());
                cNvPr.SetAttribute("name", $"ImageAt{GetColumnName(col)}{row + 1}");

                // <a:extLst>...<a16:creationId ... />
                var extLst = doc.CreateElement("a", "extLst", ANamespace);
                var extNode = doc.CreateElement("a", "ext", ANamespace);
                extNode.SetAttribute("uri", "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}");

                var creationId = doc.CreateElement("a16", "creationId", "http://schemas.microsoft.com/office/drawing/2014/main");
                creationId.SetAttribute("id", "http://schemas.microsoft.com/office/drawing/2014/main", $"{{00000000-0008-0000-0000-0000{nextId:D6}000000}}");

                extNode.AppendChild(creationId);
                extLst.AppendChild(extNode);
                cNvPr.AppendChild(extLst);

                // <xdr:cNvPicPr><a:picLocks noChangeAspect="1" /></xdr:cNvPicPr>
                var cNvPicPr = doc.CreateElement("xdr", "cNvPicPr", XdrNamespace);
                var picLocks = doc.CreateElement("a", "picLocks", ANamespace);
                picLocks.SetAttribute("noChangeAspect", "1");
                cNvPicPr.AppendChild(picLocks);

                nvPicPr.AppendChild(cNvPr);
                nvPicPr.AppendChild(cNvPicPr);
                pic.AppendChild(nvPicPr);

                // <xdr:blipFill>

                var blipFill = doc.CreateElement("xdr", "blipFill", XdrNamespace);
                var blip = doc.CreateElement("a", "blip", ANamespace);

                blip.SetAttribute("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
                blip.SetAttribute("embed", ns.LookupNamespace("r"), relId);
                blip.SetAttribute("cstate", "print");


                var stretch = doc.CreateElement("a", "stretch", ANamespace);
                var fillRect = doc.CreateElement("a", "fillRect", ANamespace);
                stretch.AppendChild(fillRect);

                blipFill.AppendChild(blip);
                blipFill.AppendChild(stretch);
                pic.AppendChild(blipFill);

                // <xdr:spPr>
                var spPr = doc.CreateElement("xdr", "spPr", XdrNamespace);
                var xfrm = doc.CreateElement("a", "xfrm", ANamespace);

                var off = doc.CreateElement("a", "off", ANamespace);
                off.SetAttribute("x", "0");
                off.SetAttribute("y", "0");

                var spExt = doc.CreateElement("a", "ext", ANamespace);
                spExt.SetAttribute("cx", "0");
                spExt.SetAttribute("cy", "0");

                xfrm.AppendChild(off);
                xfrm.AppendChild(spExt);

                var prstGeom = doc.CreateElement("a", "prstGeom", ANamespace);
                prstGeom.SetAttribute("prst", "rect");

                var avLst = doc.CreateElement("a", "avLst", ANamespace);
                prstGeom.AppendChild(avLst);

                spPr.AppendChild(xfrm);
                spPr.AppendChild(prstGeom);

                pic.AppendChild(spPr);

                // <xdr:clientData />
                var clientData = doc.CreateElement("xdr", "clientData", XdrNamespace);

                oneCellAnchor.AppendChild(from);
                oneCellAnchor.AppendChild(ext);
                oneCellAnchor.AppendChild(pic);
                oneCellAnchor.AppendChild(clientData);

                wsDr.AppendChild(oneCellAnchor);

                return doc;
            }

            private static void AppendXmlElement(XmlDocument doc, XmlElement parent, string prefix, string localName, string value)
            {
                var el = doc.CreateElement(prefix, localName, prefix == "xdr" ? XdrNamespace : ANamespace);
                el.InnerText = value;
                parent.AppendChild(el);
            }
        }
    }
}