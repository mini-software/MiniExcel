﻿using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using System.Xml.XPath;

namespace MiniExcelLibs.Tests.Utils;

internal static class Helpers
{
    private const int GENERAL_COLUMN_INDEX = 255;
    private const int MAX_COLUMN_INDEX = 16383;
    private static Dictionary<int, string>? _IntMappingAlphabet;
    private static Dictionary<string, int>? _AlphabetMappingInt;
    
    static Helpers()
    {
        if (_IntMappingAlphabet != null || _AlphabetMappingInt != null)
            return;
        
        _IntMappingAlphabet = new Dictionary<int, string>();
        _AlphabetMappingInt = new Dictionary<string, int>();
        for (int i = 0; i <= GENERAL_COLUMN_INDEX; i++)
        {
            _IntMappingAlphabet.Add(i, IntToLetters(i));
            _AlphabetMappingInt.Add(IntToLetters(i), i);
        }
    }

    public static string GetAlphabetColumnName(int columnIndex)
    {
        CheckAndSetMaxColumnIndex(columnIndex);
        return _IntMappingAlphabet[columnIndex];
    }

    public static int GetColumnIndex(string columnName)
    {
        var columnIndex = _AlphabetMappingInt[columnName];
        CheckAndSetMaxColumnIndex(columnIndex);
        return columnIndex;
    }

    private static void CheckAndSetMaxColumnIndex(int columnIndex)
    {
        if (columnIndex < _IntMappingAlphabet.Count)
            return;
        
        if (columnIndex > MAX_COLUMN_INDEX)
            throw new InvalidDataException($"ColumnIndex {columnIndex} is over Excel vaild max index.");
        
        for (int i = _IntMappingAlphabet.Count; i <= columnIndex; i++)
        {
            _IntMappingAlphabet.Add(i, IntToLetters(i));
            _AlphabetMappingInt.Add(IntToLetters(i), i);
        }
    }

    internal static string IntToLetters(int value)
    {
        value++;
        var result = string.Empty;
        
        while (--value >= 0)
        {
            result = (char)('A' + value % 26) + result;
            value /= 26;
        }
        
        return result;
    }
    
    internal static string GetZipFileContent(string zipPath, string filePath)
    {
        var ns = new XmlNamespaceManager(new NameTable());
        ns.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");

        using var stream = File.OpenRead(zipPath);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read, false, Encoding.UTF8);
        var sheet = archive.Entries.Single(w => w.FullName == filePath);

        using var sheetStream = sheet.Open();
        var doc = XDocument.Load(sheetStream);
        return doc.ToString();
    }

    internal static string GetFirstSheetDimensionRefValue(string path)
    {
        var ns = new XmlNamespaceManager(new NameTable());
        ns.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");

        using var stream = File.OpenRead(path);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read, false, Encoding.UTF8);
        var sheet = archive.Entries
            .Single(w => w.FullName.StartsWith("xl/worksheets/sheet1", StringComparison.OrdinalIgnoreCase) ||
                         w.FullName.StartsWith("/xl/worksheets/sheet1", StringComparison.OrdinalIgnoreCase));
            
        using var sheetStream = sheet.Open();
        var doc = XDocument.Load(sheetStream);
        var dimension = doc.XPathSelectElement("/x:worksheet/x:dimension", ns);
        var refV = dimension.Attribute("ref").Value;

        return refV;
    }

    internal static Dictionary<int, string> GetFirstSheetMergedCells(string path)
    {
        var ns = new XmlNamespaceManager(new NameTable());
        ns.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
        var mergeCellsDict = new Dictionary<int, string>();
        
        using var stream = File.OpenRead(path);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read, false, Encoding.UTF8);
        var sheet = archive.Entries
            .Single(w => w.FullName.StartsWith("xl/worksheets/sheet1", StringComparison.OrdinalIgnoreCase) ||
                         w.FullName.StartsWith("/xl/worksheets/sheet1", StringComparison.OrdinalIgnoreCase));

        using var sheetStream = sheet.Open();
        var doc = new XmlDocument();
        doc.Load(sheetStream);
        var mergeCells = doc.SelectSingleNode($"/x:worksheet/x:mergeCells", ns)?.Cast<XmlElement>().ToList();

        if (mergeCells is { Count: > 0 })
        {
            for (int i = 0; i < mergeCells.Count; i++)
            {
                mergeCellsDict.Add(i, mergeCells[i].GetAttribute("ref"));
            }
        }

        return mergeCellsDict;
    }
}