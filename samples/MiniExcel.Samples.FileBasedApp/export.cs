#! /bin/bash

#:package MiniExcel@2.0.0-preview.4

using MiniExcelLib;
using MiniExcelLib.OpenXml;

object[] data = 
[ 
	new { Name = "John", Surname = "Smith", Age = 25 }, 
	new { Name = "Jane", Surname = "Doe", Age = 21 }
];

MiniExcel.Exporters.GetOpenXmlExporter().Export("test.xlsx", data);
Console.WriteLine("Document saved succesfully");