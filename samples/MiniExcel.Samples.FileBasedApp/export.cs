#! /bin/bash

#:package MiniExcel@1.41.4

using MiniExcelLibs;

object[] data = 
[ 
	new { Name = "John", Surname = "Smith", Age = 25 }, 
	new { Name = "Jane", Surname = "Doe", Age = 21 }
];

MiniExcel.SaveAs("test.xlsx", data);
Console.WriteLine("Document saved succesfully");