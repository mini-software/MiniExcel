using MiniExcelLib.Core.FluentMapping;
using MiniExcelLib.Tests.Common.Utils;

namespace MiniExcelLib.Tests.FluentMapping;

public class MiniExcelMappingTemplateTests
{
    private readonly OpenXmlImporter _importer = MiniExcel.Importers.GetOpenXmlImporter();
    private readonly OpenXmlExporter _exporter = MiniExcel.Exporters.GetOpenXmlExporter();
    
    private static DateTime ParseDateValue(object? value)
    {
        return value switch
        {
            double serialDate => DateTime.FromOADate(serialDate),
            DateTime dt => dt,
            _ => DateTime.Parse(value?.ToString() ?? "")
        };
    }

    private class TestEntity
    {
        public string Name { get; set; } = "";
        public DateTime CreateDate { get; set; }
        public bool VIP { get; set; }
        public int Points { get; set; }
    }

    private class Department
    {
        public string Title { get; set; } = "";
        public List<Person> Managers { get; set; } = [];
        public List<Person> Employees { get; set; } = [];
    }

    private class Person
    {
        public string Name { get; set; } = "";
        public string Department { get; set; } = "";
    }
    
    [Fact]
    public async Task BasicTemplateTest()
    {
        using var templatePath = AutoDeletingPath.Create();
        
        var templateData = new[]
        {
            new { A = "Name", B = "Date", C = "VIP", D = "Points" },
            new { A = "", B = "", C = "", D = "" } // Empty row for data
        };
        await _exporter.ExportAsync(templatePath.ToString(), templateData);
        
        var registry = new MappingRegistry();
        registry.Configure<TestEntity>(config =>
        {
            config.Property(x => x.Name).ToCell("A3");
            config.Property(x => x.CreateDate).ToCell("B3");
            config.Property(x => x.VIP).ToCell("C3");
            config.Property(x => x.Points).ToCell("D3");
        });
        
        var data = new TestEntity
        {
            Name = "Jack",
            CreateDate = new DateTime(2021, 01, 01),
            VIP = true,
            Points = 123
        };
        
        using var outputPath = AutoDeletingPath.Create();
        var templater = MiniExcel.Templaters.GetMappingTemplater(registry);
        await templater.ApplyTemplateAsync(outputPath.ToString(), templatePath.ToString(), [data]);
        
        var rows = _importer.Query(outputPath.ToString(), useHeaderRow: false).ToList();
        
        Assert.Equal(3, rows.Count);
        
        // Row 0 is column headers
        Assert.Equal("A", rows[0].A);
        Assert.Equal("B", rows[0].B);
        Assert.Equal("C", rows[0].C);
        Assert.Equal("D", rows[0].D);
        
        // Row 1 is our custom headers
        Assert.Equal("Name", rows[1].A);
        Assert.Equal("Date", rows[1].B);
        Assert.Equal("VIP", rows[1].C);
        Assert.Equal("Points", rows[1].D);
        
        // Row 2 is the data
        Assert.Equal("Jack", rows[2].A);
        Assert.Equal(new DateTime(2021, 01, 01), ParseDateValue(rows[2].B));
        Assert.Equal(true, rows[2].C);
        Assert.Equal(123, rows[2].D);
    }
    
    [Fact]
    public async Task StreamOverloadTest()
    {
        using var templatePath = AutoDeletingPath.Create();
        var templateData = new[]
        {
            new { A = "Name", B = "Date", C = "VIP", D = "Points" },
            new { A = "", B = "", C = "", D = "" }
        };
        await _exporter.ExportAsync(templatePath.ToString(), templateData);
        
        var registry = new MappingRegistry();
        registry.Configure<TestEntity>(config =>
        {
            config.Property(x => x.Name).ToCell("A3");
            config.Property(x => x.CreateDate).ToCell("B3");
            config.Property(x => x.VIP).ToCell("C3");
            config.Property(x => x.Points).ToCell("D3");
        });
        
        var data = new TestEntity
        {
            Name = "Jack",
            CreateDate = new DateTime(2021, 01, 01),
            VIP = true,
            Points = 123
        };
        
        // Test stream overload
        using var outputPath = AutoDeletingPath.Create();
        using (var outputStream = File.Create(outputPath.ToString()))
        using (var templateStream = File.OpenRead(templatePath.ToString()))
        {
            var templater = MiniExcel.Templaters.GetMappingTemplater(registry);
            await templater.ApplyTemplateAsync(outputStream, templateStream, [data]);
        }
        
        var rows = _importer.Query(outputPath.ToString(), useHeaderRow: false).ToList();
        Assert.Equal("Jack", rows[2].A);
    }
    
    [Fact]
    public async Task ByteArrayOverloadTest()
    {
        using var templatePath = AutoDeletingPath.Create();
        var templateData = new[]
        {
            new { A = "Name", B = "Date", C = "VIP", D = "Points" },
            new { A = "", B = "", C = "", D = "" }
        };
        await _exporter.ExportAsync(templatePath.ToString(), templateData);
        
        var templateBytes = await File.ReadAllBytesAsync(templatePath.ToString());
        
        var registry = new MappingRegistry();
        registry.Configure<TestEntity>(config =>
        {
            config.Property(x => x.Name).ToCell("A3");
            config.Property(x => x.CreateDate).ToCell("B3");
            config.Property(x => x.VIP).ToCell("C3");
            config.Property(x => x.Points).ToCell("D3");
        });
        
        var data = new TestEntity
        {
            Name = "Jack",
            CreateDate = new DateTime(2021, 01, 01),
            VIP = true,
            Points = 123
        };
        
        using var outputPath = AutoDeletingPath.Create();
        using (var outputStream = File.Create(outputPath.ToString()))
        {
            var templater = MiniExcel.Templaters.GetMappingTemplater(registry);
            await templater.ApplyTemplateAsync(outputStream, templateBytes, [data]);
        }
        
        var rows = _importer.Query(outputPath.ToString(), useHeaderRow: false).ToList();
        Assert.Equal("Jack", rows[2].A);
    }
    
    [Fact]
    public async Task CollectionTemplateTest()
    {
        using var templatePath = AutoDeletingPath.Create();
        var templateData = new List<dynamic>
        {
            new { A = "Company", B = "", C = "" },
            new { A = "", B = "", C = "" }, // Row 2
            new { A = "Managers", B = "Department", C = "" } // Row 3
        };

        for (int i = 0; i < 3; i++)
        {
            templateData.Add(new { A = "", B = "", C = "" });
        }
        
        templateData.Add(new { A = "Employees", B = "Department", C = "" }); // Row 7
        
        for (int i = 0; i < 3; i++)
        {
            templateData.Add(new { A = "", B = "", C = "" });
        }
        
        // Saving our actual template first
        await _exporter.ExportAsync(templatePath.ToString(), templateData);
        
        var registry = new MappingRegistry();
        
        registry.Configure<Person>(config =>
        {
            config.Property(x => x.Name).ToCell("A1");
            config.Property(x => x.Department).ToCell("B1");
        });
        
        registry.Configure<Department>(config =>
        {
            config.Property(x => x.Title).ToCell("A2");
            config.Collection(x => x.Managers).StartAt("A5");
            config.Collection(x => x.Employees).StartAt("A9");
        });
        
        var dept = new Department
        {
            Title = "FooCompany",
            Managers =
            [
                new Person { Name = "Jack", Department = "HR" },
                new Person { Name = "Jane", Department = "IT" }
            ],
            Employees =
            [
                new Person { Name = "Wade", Department = "HR" },
                new Person { Name = "John", Department = "Sales" }
            ]
        };
        
        using var outputPath = AutoDeletingPath.Create();
        var templater = MiniExcel.Templaters.GetMappingTemplater(registry);
        await templater.ApplyTemplateAsync(outputPath.ToString(), templatePath.ToString(), [dept]);
        
        var rows = _importer.Query(outputPath.ToString(), useHeaderRow: false).ToList();
        
        
        Assert.Equal(11, rows.Count); // We expect 11 rows total
        
        Assert.Equal("FooCompany", rows[1].A);
        
        Assert.Equal("Managers", rows[3].A);
        Assert.Equal("Department", rows[3].B);
        
        Assert.Equal("Jack", rows[4].A);
        Assert.Equal("HR", rows[4].B);
        Assert.Equal("Jane", rows[5].A);
        Assert.Equal("IT", rows[5].B);
        
        Assert.Equal("Employees", rows[7].A);
        Assert.Equal("Department", rows[7].B);
        
        Assert.Equal("Wade", rows[8].A);
        Assert.Equal("HR", rows[8].B);
        Assert.Equal("John", rows[9].A);
        Assert.Equal("Sales", rows[9].B);
    }
    
    [Fact]
    public async Task EmptyDataTest()
    {
        using var templatePath = AutoDeletingPath.Create();
        var templateData = new[]
        {
            new { A = "Name", B = "Date", C = "VIP", D = "Points" },
            new { A = "", B = "", C = "", D = "" }
        };
        await _exporter.ExportAsync(templatePath.ToString(), templateData);
        
        var registry = new MappingRegistry();
        registry.Configure<TestEntity>(config =>
        {
            config.Property(x => x.Name).ToCell("A3");
            config.Property(x => x.CreateDate).ToCell("B3");
            config.Property(x => x.VIP).ToCell("C3");
            config.Property(x => x.Points).ToCell("D3");
        });
        
        using var outputPath = AutoDeletingPath.Create();
        var templater = MiniExcel.Templaters.GetMappingTemplater(registry);
        await templater.ApplyTemplateAsync(outputPath.ToString(), templatePath.ToString(), Array.Empty<TestEntity>());
        
        var rows = _importer.Query(outputPath.ToString(), useHeaderRow: false).ToList();
        Assert.Equal(3, rows.Count); // Column headers + our headers + empty data row
        Assert.Equal("Name", rows[1].A);
        Assert.Equal("Date", rows[1].B);
        
        // Third row should be empty
        Assert.True(string.IsNullOrEmpty(rows[2].A?.ToString()));
    }
    
    [Fact]
    public async Task NullValuesTest()
    {
        // Create template
        using var templatePath = AutoDeletingPath.Create();
        var templateData = new[]
        {
            new { A = "Name", B = "Date", C = "VIP", D = "Points" },
            new { A = "Default", B = "2020-01-01", C = "false", D = "0" }
        };
        await _exporter.ExportAsync(templatePath.ToString(), templateData);
        
        // Setup mapping
        var registry = new MappingRegistry();
        registry.Configure<TestEntity>(config =>
        {
            config.Property(x => x.Name).ToCell("A3");
            config.Property(x => x.CreateDate).ToCell("B3");
            config.Property(x => x.VIP).ToCell("C3");
            config.Property(x => x.Points).ToCell("D3");
        });
        
        var data = new TestEntity
        {
            Name = null!, // Null value
            CreateDate = new DateTime(2021, 01, 01),
            VIP = false,
            Points = 0
        };
        
        // Apply template
        using var outputPath = AutoDeletingPath.Create();
        var templater = MiniExcel.Templaters.GetMappingTemplater(registry);
        await templater.ApplyTemplateAsync(outputPath.ToString(), templatePath.ToString(), [data]);
        
        // Verify null handling
        // Verify - use useHeaderRow=false since we want to see all rows
        var rows = _importer.Query(outputPath.ToString(), useHeaderRow: false).ToList();
        Assert.True(string.IsNullOrEmpty(rows[2].A?.ToString())); // Null replaced the default
        Assert.Equal(new DateTime(2021, 01, 01), ParseDateValue(rows[2].B));
        Assert.Equal(false, rows[2].C);
        Assert.Equal(0, rows[2].D);
    }
    
    [Fact]
    public async Task MultipleItemsTest()
    {
        // Create template with space for multiple items
        using var templatePath = AutoDeletingPath.Create();
        var templateData = new[]
        {
            new { A = "Name", B = "Date", C = "VIP", D = "Points" },
            new { A = "", B = "", C = "", D = "" },
            new { A = "", B = "", C = "", D = "" },
            new { A = "", B = "", C = "", D = "" }
        };
        await _exporter.ExportAsync(templatePath.ToString(), templateData);
        
        // Setup mapping for multiple rows
        var registry = new MappingRegistry();
        registry.Configure<TestEntity>(config =>
        {
            config.Property(x => x.Name).ToCell("A3");
            config.Property(x => x.CreateDate).ToCell("B3");
            config.Property(x => x.VIP).ToCell("C3");
            config.Property(x => x.Points).ToCell("D3");
        });
        
        var data = new[]
        {
            new TestEntity { Name = "Jack", CreateDate = new DateTime(2021, 01, 01), VIP = true, Points = 123 },
            new TestEntity { Name = "Jane", CreateDate = new DateTime(2021, 01, 02), VIP = false, Points = 456 },
            new TestEntity { Name = "John", CreateDate = new DateTime(2021, 01, 03), VIP = true, Points = 789 }
        };
        
        // Apply template
        using var outputPath = AutoDeletingPath.Create();
        var templater = MiniExcel.Templaters.GetMappingTemplater(registry);
        await templater.ApplyTemplateAsync(outputPath.ToString(), templatePath.ToString(), data);
        
        // Verify - should only update first item since mapping is for specific cells
        // Verify - use useHeaderRow=false since we want to see all rows
        var rows = _importer.Query(outputPath.ToString(), useHeaderRow: false).ToList();
        Assert.Equal("Jack", rows[2].A);
        Assert.Equal(123, rows[2].D);
        
        // Other rows should remain empty as we mapped to specific cells (A3, B3, etc.)
        Assert.True(string.IsNullOrEmpty(rows[3].A?.ToString()));
        Assert.True(string.IsNullOrEmpty(rows[4].A?.ToString()));
    }
}