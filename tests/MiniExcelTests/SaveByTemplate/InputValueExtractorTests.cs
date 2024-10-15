using System.Collections.Generic;
using System.Data;
using FluentAssertions;
using MiniExcelLibs.OpenXml.SaveByTemplate;
using Xunit;

namespace MiniExcelTests.SaveByTemplate;

public class InputValueExtractorTests
{
    [Fact]
    public void ToValueDictionary_Given_InputIsDictionaryWithoutDataReader_Then_Output_IsAnEquivalentDictionary()
    {
        var valueDictionary = new Dictionary<string, object>
        {
            ["Name"] = "John",
            ["Age"] = 18,
            ["Fruits"] = new List<string> { "Apples, Oranges" },
        };

        var sut = new InputValueExtractor();
        var result = sut.ToValueDictionary(valueDictionary);

        result.Should().BeEquivalentTo(valueDictionary);
    }

    [Fact]
    public void ToValueDictionary_Given_InputIsDictionaryWithDataReader_Then_DataReaderIsConvertedToListOfDictionaries()
    {
        var dataTable = new DataTable();

        dataTable.Columns.Add("id", typeof(int));
        dataTable.Columns.Add("name", typeof(string));
        dataTable.Rows.Add(1, "Jack");
        dataTable.Rows.Add(2, "Mike");

        var expectedOutput = new List<Dictionary<string, object>>
        {
            new() { ["id"] = 1, ["name"] = "Jack" },
            new() { ["id"] = 2, ["name"] = "Mike" }
        };

        var valueDictionary = new Dictionary<string, object>
        {
            ["DataReader"] = dataTable.CreateDataReader()
        };

        var sut = new InputValueExtractor();
        var result = sut.ToValueDictionary(valueDictionary);

        result["DataReader"].Should().BeEquivalentTo(expectedOutput);
    }

    [Fact]
    public void ToValueDictionary_Given_InputIsPocoRecord_Then_Output_IsAnEquivalentDictionary()
    {
        var valueObject = new PocoRecord("John", 18, new List<string> { "Apples, Oranges" });

        var expectedOutput = new Dictionary<string, object>()
        {
            ["Name"] = "John",
            ["Age"] = 18,
            ["Fruits"] = new List<string> { "Apples, Oranges" }
        };

        var sut = new InputValueExtractor();
        var result = sut.ToValueDictionary(valueObject);

        result.Should().BeEquivalentTo(expectedOutput);
    }

    [Fact]
    public void ToValueDictionary_Given_InputIsPocoClass_Then_Output_IsAnEquivalentDictionary()
    {
        var valueObject = new PocoClass
        {
            Name = "John",
            Age = 18,
            Fruits = new List<string> { "Apples, Oranges" }
        };

        var expectedOutput = new Dictionary<string, object>()
        {
            ["Name"] = "John",
            ["Age"] = 18,
            ["Fruits"] = new List<string> { "Apples, Oranges" }
        };

        var sut = new InputValueExtractor();
        var result = sut.ToValueDictionary(valueObject);

        result.Should().BeEquivalentTo(expectedOutput);
    }


    private record PocoRecord(string Name, int Age, IEnumerable<string> Fruits);

    private class PocoClass
    {
        public string Name { get; set; }
        public int Age { get; set; }
        public IEnumerable<string> Fruits; // Field
    };
}