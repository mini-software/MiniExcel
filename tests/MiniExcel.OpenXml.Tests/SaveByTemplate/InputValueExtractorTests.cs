using MiniExcelLib.OpenXml.Templates;

namespace MiniExcelLib.OpenXml.Tests.SaveByTemplate;

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

        var sut = new OpenXmlValueExtractor();
        var result = sut.ToValueDictionary(valueDictionary);
        
        Assert.Equal(result.Count, valueDictionary.Count);
        foreach (var (key, value) in result)
        {
            Assert.Equal(value, valueDictionary[key]);
        }
    }

    [Fact]
    public void ToValueDictionary_Given_InputIsDictionaryWithDataReader_Then_DataReaderIsConvertedToListOfDictionaries()
    {
        using var dataTable = new DataTable();

        dataTable.Columns.Add("id", typeof(int));
        dataTable.Columns.Add("name", typeof(string));
        dataTable.Rows.Add(1, "Jack");
        dataTable.Rows.Add(2, "Mike");

        List<Dictionary<string, object>> expectedOutput =
        [
            new() { ["id"] = 1, ["name"] = "Jack" },
            new() { ["id"] = 2, ["name"] = "Mike" }
        ];

        var valueDictionary = new Dictionary<string, object>
        {
            ["DataReader"] = dataTable.CreateDataReader()
        };

        var sut = new OpenXmlValueExtractor();
        var extracted = sut.ToValueDictionary(valueDictionary);
        var result = (List<IDictionary<string, object>>)extracted["DataReader"];

        Assert.Equal(result.Count, expectedOutput.Count);
        for (int i = 0; i < result.Count; i++)
        {
            var row = result[i];
            var expected = expectedOutput[i];
            
            Assert.Equal(row.Count, expected.Count);
            foreach (var (key, value) in row)
            {
                Assert.Equal(value, expected[key]);
            }
        }
    }

    [Fact]
    public void ToValueDictionary_Given_InputIsPocoRecord_Then_Output_IsAnEquivalentDictionary()
    {
        var valueObject = new PocoRecord("John", 18, new List<string> { "Apples, Oranges" });
        var expectedOutput = new Dictionary<string, object>
        {
            ["Name"] = "John",
            ["Age"] = 18,
            ["Fruits"] = new List<string> { "Apples, Oranges" }
        };

        var sut = new OpenXmlValueExtractor();
        var result = sut.ToValueDictionary(valueObject);

        Assert.Equal(result.Count, expectedOutput.Count);
        foreach (var (key, value) in result)
        {
            Assert.Equal(value, expectedOutput[key]);
        }
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

        var expectedOutput = new Dictionary<string, object>
        {
            ["Name"] = "John",
            ["Age"] = 18,
            ["Fruits"] = new List<string> { "Apples, Oranges" }
        };

        var sut = new OpenXmlValueExtractor();
        var result = sut.ToValueDictionary(valueObject);
     
        Assert.Equal(result.Count, expectedOutput.Count);
        foreach (var (key, value) in result)
        {
            Assert.Equal(value, expectedOutput[key]);
        }
    }

    private record PocoRecord(string Name, int Age, IEnumerable<string> Fruits);
    private class PocoClass
    {
        public string Name { get; set; }
        public int Age { get; set; }
        public IEnumerable<string> Fruits; // Field
    };
}