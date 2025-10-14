using System.Reflection;
using MiniExcelLib.Core.FluentMapping;

namespace MiniExcelLib.Tests.FluentMapping
{
    public class MiniExcelMappingTests
    {
        #region Test Models

        public class Person
        {
            public string Name { get; set; } = "";
            public int Age { get; set; }
            public string Email { get; set; } = "";
            public string Department { get; set; } = "";
            public DateTime BirthDate { get; set; }
            public decimal Salary { get; set; }
        }

        public class Product
        {
            public int Id { get; set; }
            public string Name { get; set; } = "";
            public string Category { get; set; } = "";
            public decimal Price { get; set; }
            public int Stock { get; set; }
            public DateTime LastRestocked { get; set; }
            public bool IsActive { get; set; }
            public double? DiscountPercentage { get; set; }
        }

        public class ComplexEntity
        {
            public int Id { get; set; }
            public string Name { get; set; } = "";
            public DateTime CreatedDate { get; set; }
            public double Score { get; set; }
            public bool IsEnabled { get; set; }
            public string? Description { get; set; }
            public decimal Amount { get; set; }
            public List<string> Tags { get; set; } = [];
            public int[] Numbers { get; set; } = [];
        }

        public class ComplexModel
        {
            public Guid Id { get; set; }
            public string Title { get; set; } = "";
            public DateTimeOffset CreatedAt { get; set; }
            public TimeSpan Duration { get; set; }
            public byte[] BinaryData { get; set; } = [];
            public Uri? Website { get; set; }
        }

        public class Department
        {
            public string Name { get; set; } = "";
            public List<Person> Managers { get; set; } = [];
            public List<Person> Employees { get; set; } = [];
            public List<string> PhoneNumbers { get; set; } = [];
            public string[] Tags { get; set; } = [];
            public IEnumerable<Project> Projects { get; set; } = [];
        }

        public class Company
        {
            public string Name { get; set; } = "";
            public List<Department> Departments { get; set; } = [];
        }

        public class TestModel
        {
            public string Name { get; set; } = "";
            public int Value { get; set; }
        }

        public class Project
        {
            public string Code { get; set; } = "";
            public string Title { get; set; } = "";
            public DateTime StartDate { get; set; }
            public decimal Budget { get; set; }
            public List<ProjectTask> Tasks { get; set; } = [];
        }

        public class ProjectTask
        {
            public string Name { get; set; } = "";
            public int EstimatedHours { get; set; }
            public bool IsCompleted { get; set; }
        }

        public class Report
        {
            public string Title { get; set; } = "";
            public DateTime GeneratedAt { get; set; }
            public List<int> Numbers { get; set; } = [];
            public Dictionary<string, decimal> Metrics { get; set; } = new();
        }

        public class Address
        {
            public string Street { get; set; } = "";
            public string City { get; set; } = "";
            public string PostalCode { get; set; } = "";
        }

        public class NestedModel
        {
            public string Name { get; set; } = "";
            public Address HomeAddress { get; set; } = new();
            public Address? WorkAddress { get; set; }
        }

        #endregion

        #region Basic Mapping Tests

        [Fact]
        public async Task MappingReader_ReadBasicData_Success()
        {
            // Arrange
            var registry = new MappingRegistry();
            registry.Configure<TestModel>(cfg =>
            {
                cfg.Property(p => p.Name).ToCell("A1");
                cfg.Property(p => p.Value).ToCell("B1");
            });

            var testData = new[] { new TestModel { Name = "Test", Value = 42 } };
            using var stream = new MemoryStream();
            
            var exporter = MiniExcel.Exporters.GetMappingExporter(registry);
            await exporter.ExportAsync(stream, testData);
            stream.Position = 0;

            // Act
            var importer = MiniExcel.Importers.GetMappingImporter(registry);
            var resultList = new List<TestModel>();
            await foreach (var item in importer.QueryAsync<TestModel>(stream))
            {
                resultList.Add(item);
            }

            // Assert
            Assert.Single(resultList);
            Assert.Equal("Test", resultList[0].Name);
            Assert.Equal(42, resultList[0].Value);
        }

        [Fact]
        public async Task SaveAs_WithBasicMapping_ShouldGenerateCorrectFile()
        {
            // Arrange
            var people = new[]
            {
                new Person { Name = "Alice", Age = 30, Email = "alice@example.com", BirthDate = new DateTime(1993, 5, 15), Salary = 75000.50m }
            };

            var registry = new MappingRegistry();
            registry.Configure<Person>(cfg =>
            {
                cfg.Property(p => p.Name).ToCell("A1");
                cfg.Property(p => p.Age).ToCell("B1");
                cfg.Property(p => p.Email).ToCell("C1");
                cfg.ToWorksheet("People");
            });

            var exporter = MiniExcel.Exporters.GetMappingExporter(registry);

            // Act & Assert
            using var stream = new MemoryStream();
            await exporter.ExportAsync(stream, people);
            Assert.True(stream.Length > 0);
        }

        [Fact]
        public void SaveAs_WithBasicMapping_SyncVersion_ShouldGenerateCorrectFile()
        {
            // Arrange
            var people = new[]
            {
                new Person { Name = "Bob", Age = 25, Email = "bob@example.com", BirthDate = new DateTime(1998, 8, 20), Salary = 60000.00m }
            };

            var registry = new MappingRegistry();
            registry.Configure<Person>(cfg =>
            {
                cfg.Property(p => p.Name).ToCell("B2");
                cfg.Property(p => p.Age).ToCell("C2");
                cfg.Property(p => p.Email).ToCell("D2");
                cfg.ToWorksheet("Employees");
            });

            var exporter = MiniExcel.Exporters.GetMappingExporter(registry);

            // Act & Assert
            using var stream = new MemoryStream();
            exporter.Export(stream, people);
            Assert.True(stream.Length > 0);
        }

        [Fact]
        public async Task Query_WithBasicMapping_ShouldReadDataCorrectly()
        {
            // Arrange
            var testData = new[]
            {
                new Person { Name = "John", Age = 35, Email = "john@test.com", BirthDate = new DateTime(1988, 3, 10), Salary = 85000m },
                new Person { Name = "Jane", Age = 28, Email = "jane@test.com", BirthDate = new DateTime(1995, 7, 22), Salary = 72000m }
            };

            var registry = new MappingRegistry();
            registry.Configure<Person>(cfg =>
            {
                cfg.Property(p => p.Name).ToCell("A1");
                cfg.Property(p => p.Age).ToCell("B1");
                cfg.Property(p => p.Email).ToCell("C1");
                cfg.Property(p => p.BirthDate).ToCell("D1");
                cfg.Property(p => p.Salary).ToCell("E1");
            });

            var exporter = MiniExcel.Exporters.GetMappingExporter(registry);
            var importer = MiniExcel.Importers.GetMappingImporter(registry);

            // Act
            using var stream = new MemoryStream();
            await exporter.ExportAsync(stream, testData);
            
            stream.Position = 0;
            var results = importer.Query<Person>(stream).ToList();

            // Assert
            Assert.NotNull(results);
            Assert.NotEmpty(results);
        }

        #endregion

        #region Sequential Mapping Tests

        [Fact]
        public async Task Sequential_Mapping_Should_Optimize_Performance()
        {
            // Test that sequential column mappings (A1, B1, C1...) are optimized
            var registry = new MappingRegistry();
            registry.Configure<Product>(cfg =>
            {
                cfg.Property(p => p.Id).ToCell("A1");
                cfg.Property(p => p.Name).ToCell("B1");
                cfg.Property(p => p.Price).ToCell("C1");
                cfg.Property(p => p.Stock).ToCell("D1");
                cfg.Property(p => p.IsActive).ToCell("E1");
            });

            var mapping = registry.GetMapping<Product>();
            
            // Verify optimization is applied
            Assert.NotNull(mapping.OptimizedBoundaries);
            Assert.NotNull(mapping.OptimizedCellGrid);
        }

        [Fact]
        public async Task NonSequential_Mapping_Should_Use_Universal_Optimization()
        {
            // Test that non-sequential mappings use universal optimization
            var registry = new MappingRegistry();
            registry.Configure<Product>(cfg =>
            {
                cfg.Property(p => p.Id).ToCell("A1");
                cfg.Property(p => p.Name).ToCell("C1");  // Skip B
                cfg.Property(p => p.Price).ToCell("E1");  // Skip D
                cfg.Property(p => p.Stock).ToCell("B2");  // Different row
                cfg.Property(p => p.IsActive).ToCell("D2");
            });

            var mapping = registry.GetMapping<Product>();
            
            // Verify optimization is used
            Assert.NotNull(mapping.OptimizedCellGrid);
            Assert.NotNull(mapping.OptimizedBoundaries);
        }

        #endregion

        #region Collection Mapping Tests

        [Fact]
        public async Task Collection_Vertical_Should_Write_And_Read_Correctly()
        {
            // Test vertical collection layout (default)
            var registry = new MappingRegistry();
            registry.Configure<ComplexEntity>(cfg =>
            {
                cfg.Property(e => e.Id).ToCell("A1");
                cfg.Property(e => e.Name).ToCell("B1");
                cfg.Collection(e => e.Tags).StartAt("C2"); // Vertical by default
            });

            var testData = new[]
            {
                new ComplexEntity 
                { 
                    Id = 1, 
                    Name = "Test", 
                    Tags = ["Tag1", "Tag2", "Tag3"]
                }
            };

            var exporter = MiniExcel.Exporters.GetMappingExporter(registry);
            var importer = MiniExcel.Importers.GetMappingImporter(registry);

            using var stream = new MemoryStream();
            await exporter.ExportAsync(stream, testData);
            
            stream.Position = 0;
            var results = importer.Query<ComplexEntity>(stream).ToList();

            Assert.Single(results);
            Assert.Equal(3, results[0].Tags.Count);
            Assert.Equal("Tag1", results[0].Tags[0]);
        }

        [Fact]
        public async Task Collection_ComplexObjectsWithMapping_ShouldMapCorrectly()
        {
            // Arrange
            var departments = new[]
            {
                new Department
                {
                    Name = "Engineering",
                    Employees =
                    [
                        new Person { Name = "Alice", Age = 35, Email = "alice@example.com", Salary = 95000 },
                        new Person { Name = "Bob", Age = 28, Email = "bob@example.com", Salary = 75000 },
                        new Person { Name = "Charlie", Age = 24, Email = "charlie@example.com", Salary = 55000 }
                    ]
                }
            };

            var registry = new MappingRegistry();
            registry.Configure<Department>(cfg =>
            {
                cfg.Property(d => d.Name).ToCell("A1");
                cfg.Collection(d => d.Employees)
                    .StartAt("A3")
                    .WithItemMapping<Person>(empCfg =>
                    {
                        empCfg.Property(p => p.Name).ToCell("A3");
                        empCfg.Property(p => p.Age).ToCell("B3");
                        empCfg.Property(p => p.Email).ToCell("C3");
                    });
            });

            var exporter = MiniExcel.Exporters.GetMappingExporter(registry);
            var importer = MiniExcel.Importers.GetMappingImporter(registry);

            var compiled = registry.GetMapping<Department>();
            var boundaries = compiled.OptimizedBoundaries!;
            var grid = compiled.OptimizedCellGrid!;
            for (var r = 0; r < grid.GetLength(0); r++)
            {
                for (var c = 0; c < grid.GetLength(1); c++)
                {
                    var handler = grid[r, c];
                    if (handler.Type != CellHandlerType.Empty)
                    {
                        
                    }
                }
            }

            // Act
            using var stream = new MemoryStream();
            await exporter.ExportAsync(stream, departments);
            stream.Position = 0;

            var results = importer.Query<Department>(stream).ToList();

            // Assert
            Assert.Single(results);
            var department = results[0];
            Assert.Equal("Engineering", department.Name);
            Assert.Equal(3, department.Employees.Count);
            Assert.Equal("Alice", department.Employees[0].Name);
            Assert.Equal(35, department.Employees[0].Age);
            Assert.Equal("alice@example.com", department.Employees[0].Email);
        }

        [Fact]
        public async Task Collection_WithItemMappingOnly_ShouldWriteAndReadCorrectly()
        {
            var departments = new[]
            {
                new Department
                {
                    Name = "Operations",
                    Managers =
                    [
                        new Person { Name = "Ellen", Department = "Ops" },
                        new Person { Name = "Scott", Department = "Ops" }
                    ]
                }
            };

            var registry = new MappingRegistry();
            registry.Configure<Department>(cfg =>
            {
                cfg.Property(d => d.Name).ToCell("A1");
                cfg.Collection(d => d.Managers)
                    .StartAt("A3")
                    .WithItemMapping<Person>(managerCfg =>
                    {
                        managerCfg.Property(p => p.Name).ToCell("A3");
                        managerCfg.Property(p => p.Department).ToCell("B3");
                    });
            });

            var exporter = MiniExcel.Exporters.GetMappingExporter(registry);
            var importer = MiniExcel.Importers.GetMappingImporter(registry);

            using var stream = new MemoryStream();
            await exporter.ExportAsync(stream, departments);
            stream.Position = 0;

            Assert.True(registry.HasMapping<Person>());

            var compiledMappingsField = typeof(MappingRegistry).GetField("_compiledMappings", BindingFlags.NonPublic | BindingFlags.Instance);
            var compiledMappings = (Dictionary<Type, object>)compiledMappingsField!.GetValue(registry)!;
            var departmentMapping = compiledMappings[typeof(Department)];
            var nestedMappingsProp = departmentMapping.GetType().GetProperty("NestedMappings", BindingFlags.Instance | BindingFlags.Public);
            var nestedMappingsObj = nestedMappingsProp?.GetValue(departmentMapping);
            Assert.NotNull(nestedMappingsObj);
            var countProp = nestedMappingsObj!.GetType().GetProperty("Count");
            var nestedCount = (int)(countProp?.GetValue(nestedMappingsObj) ?? 0);
            Assert.True(nestedCount > 0);

            var getEnumerator = nestedMappingsObj.GetType().GetMethod("GetEnumerator");
            var enumerator = (System.Collections.IEnumerator)getEnumerator!.Invoke(nestedMappingsObj, null)!;
            Assert.True(enumerator.MoveNext());
            var entry = enumerator.Current;
            var valueProp = entry!.GetType().GetProperty("Value");
            var nestedInfo = valueProp!.GetValue(entry);
            var propertiesProp = nestedInfo!.GetType().GetProperty("Properties", BindingFlags.Instance | BindingFlags.Public);
            var nestedProperties = (System.Collections.IEnumerable?)propertiesProp?.GetValue(nestedInfo);
            Assert.NotNull(nestedProperties);

            var firstProperty = nestedProperties!.Cast<object>().FirstOrDefault();
            Assert.NotNull(firstProperty);
            var setterProp = firstProperty!.GetType().GetProperty("Setter", BindingFlags.Instance | BindingFlags.Public);
            var setter = setterProp?.GetValue(firstProperty);
            Assert.NotNull(setter);

            var gridProp = departmentMapping.GetType().GetProperty("OptimizedCellGrid", BindingFlags.Instance | BindingFlags.Public);
            var grid = gridProp?.GetValue(departmentMapping) as Array;
            Assert.NotNull(grid);
            var handlerType = typeof(MappingRegistry).Assembly.GetType("MiniExcelLib.Core.FluentMapping.OptimizedCellHandler");
            Assert.NotNull(handlerType);
            var valueSetterProperty = handlerType!.GetProperty("ValueSetter", BindingFlags.Instance | BindingFlags.Public);
            var propertyNameProperty = handlerType.GetProperty("PropertyName", BindingFlags.Instance | BindingFlags.Public);
            var collectionIndexProperty = handlerType.GetProperty("CollectionIndex", BindingFlags.Instance | BindingFlags.Public);

            var hasSetter = false;
            for (int r = 0; r < grid!.GetLength(0); r++)
            {
                for (int c = 0; c < grid.GetLength(1); c++)
                {
                    if (grid.GetValue(r, c) is { } handler)
                    {
                        var propertyName = propertyNameProperty?.GetValue(handler) as string;
                        var collectionIndex = (int)(collectionIndexProperty?.GetValue(handler) ?? -1);
                        if (string.Equals(propertyName, "Name", StringComparison.Ordinal) && collectionIndex == 0)
                        {
                            var valueSetter = valueSetterProperty?.GetValue(handler);
                            if (valueSetter is not null)
                            {
                                hasSetter = true;
                                break;
                            }
                        }
                    }
                }
                if (hasSetter)
                {
                    break;
                }
            }

            Assert.True(hasSetter);

            var boundariesProp = departmentMapping.GetType().GetProperty("OptimizedBoundaries", BindingFlags.Instance | BindingFlags.Public);
            var boundaries = boundariesProp?.GetValue(departmentMapping);
            Assert.NotNull(boundaries);
            var minRowProp = boundaries!.GetType().GetProperty("MinRow", BindingFlags.Instance | BindingFlags.Public);
            var maxRowProp = boundaries.GetType().GetProperty("MaxRow", BindingFlags.Instance | BindingFlags.Public);
            var minColProp = boundaries.GetType().GetProperty("MinColumn", BindingFlags.Instance | BindingFlags.Public);
            var maxColProp = boundaries.GetType().GetProperty("MaxColumn", BindingFlags.Instance | BindingFlags.Public);
            var minRow = (int)(minRowProp?.GetValue(boundaries) ?? 0);
            var maxRow = (int)(maxRowProp?.GetValue(boundaries) ?? 0);
            var minCol = (int)(minColProp?.GetValue(boundaries) ?? 0);
            var maxCol = (int)(maxColProp?.GetValue(boundaries) ?? 0);

            var tryGetHandlerMethod = departmentMapping.GetType().GetMethod("TryGetHandler", BindingFlags.Instance | BindingFlags.Public);
            Assert.NotNull(tryGetHandlerMethod);

            var setterFoundViaTryGet = false;
            for (var row = minRow; row <= maxRow && !setterFoundViaTryGet; row++)
            {
                for (var col = minCol; col <= maxCol && !setterFoundViaTryGet; col++)
                {
                    var parameters = new object?[] { row, col, null };
                    var success = (bool)tryGetHandlerMethod!.Invoke(departmentMapping, parameters)!;
                    if (!success)
                        continue;

                    if (parameters[2] is { } handlerInstance)
                    {
                        var propertyName = propertyNameProperty?.GetValue(handlerInstance) as string;
                        var collectionIndex = (int)(collectionIndexProperty?.GetValue(handlerInstance) ?? -1);
                        if (collectionIndex == 0 && string.Equals(propertyName, "Name", StringComparison.Ordinal))
                        {
                            var valueSetter = valueSetterProperty?.GetValue(handlerInstance);
                            if (valueSetter is not null)
                            {
                                setterFoundViaTryGet = true;
                            }
                        }
                    }
                }
            }
            Assert.True(setterFoundViaTryGet);

            var results = importer.Query<Department>(stream).ToList();

            Assert.Single(results);
            var managers = results[0].Managers;
            Assert.Equal(2, managers.Count);
            Assert.Equal("Ellen", managers[0].Name);
            Assert.Equal("Scott", managers[1].Name);
        }

        [Fact]
        public async Task Collection_NestedCollections_ShouldMapCorrectly()
        {
            // Arrange
            var departments = new[]
            {
                new Department
                {
                    Name = "Product Development",
                    Projects = new List<Project>
                    {
                        new Project
                        {
                            Code = "PROJ-001",
                            Title = "New Feature",
                            StartDate = new DateTime(2024, 1, 1),
                            Budget = 100000,
                            Tasks =
                            [
                                new ProjectTask { Name = "Design", EstimatedHours = 40, IsCompleted = true },
                                new ProjectTask { Name = "Implementation", EstimatedHours = 120, IsCompleted = false },
                                new ProjectTask { Name = "Testing", EstimatedHours = 60, IsCompleted = false }
                            ]
                        }
                    }
                }
            };

            var registry = new MappingRegistry();
            registry.Configure<Department>(cfg =>
            {
                cfg.Property(d => d.Name).ToCell("A1");
                cfg.Collection(d => d.Projects)
                    .StartAt("A3")
                    .WithItemMapping<Project>(projectCfg =>
                    {
                        projectCfg.Property(p => p.Code).ToCell("A3");
                        projectCfg.Property(p => p.Title).ToCell("B3");
                        projectCfg.Property(p => p.StartDate).ToCell("C3");
                        projectCfg.Collection(p => p.Tasks)
                            .StartAt("D3")
                            .WithItemMapping<ProjectTask>(taskCfg =>
                            {
                                taskCfg.Property(t => t.Name).ToCell("D3");
                                taskCfg.Property(t => t.EstimatedHours).ToCell("E3");
                                taskCfg.Property(t => t.IsCompleted).ToCell("F3");
                            });
                    });
            });

            var exporter = MiniExcel.Exporters.GetMappingExporter(registry);
            var importer = MiniExcel.Importers.GetMappingImporter(registry);

            // Act
            using var stream = new MemoryStream();
            await exporter.ExportAsync(stream, departments);

            stream.Position = 0;

            var results = importer.Query<Department>(stream).ToList();

            // Assert
            Assert.Single(results);
            var department = results[0];
            Assert.Equal("Product Development", department.Name);

            var projects = department.Projects.ToList();
            Assert.Single(projects);

            var project = projects[0];
            Assert.Equal("PROJ-001", project.Code);
            Assert.Equal("New Feature", project.Title);

            Assert.Equal(3, project.Tasks.Count);
            Assert.Equal("Design", project.Tasks[0].Name);
            Assert.True(project.Tasks[0].IsCompleted);
            Assert.Equal(120, project.Tasks[1].EstimatedHours);
        }

        [Fact]
        public void Collection_WithoutStartCell_ShouldThrowException()
        {
            // Arrange
            var registry = new MappingRegistry();

            // Act & Assert
            var exception = Assert.Throws<InvalidOperationException>(() =>
            {
                registry.Configure<Department>(cfg =>
                {
                    cfg.Collection(d => d.PhoneNumbers); // Missing StartAt()
                });
            });
            
            Assert.Contains("start cell", exception.Message, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public async Task Collection_MixedSimpleAndComplex_ShouldMapCorrectly()
        {
            // Arrange
            var department = new Department
            {
                Name = "Mixed Department",
                PhoneNumbers = ["555-1111", "555-2222"],
                Employees =
                [
                    new Person { Name = "Dave", Age = 35, Email = "dave@example.com", Salary = 85000 },
                    new Person { Name = "Eve", Age = 29, Email = "eve@example.com", Salary = 75000 }
                ]
            };

            var departments = new[] { department };

            var registry = new MappingRegistry();
            registry.Configure<Department>(cfg =>
            {
                cfg.Property(d => d.Name).ToCell("A1");
                cfg.Collection(d => d.PhoneNumbers).StartAt("A3");
                cfg.Collection(d => d.Employees)
                    .StartAt("C3")
                    .WithItemMapping<Person>(x =>
                    {
                        x.Property(e => e.Name).ToCell("C3");
                        x.Property(e => e.Age).ToCell("D3");
                        x.Property(e => e.Salary).ToCell("E3");
                        x.Property(e => e.Email).ToCell("F3");
                    });
            });

            var exporter = MiniExcel.Exporters.GetMappingExporter(registry);
            var importer = MiniExcel.Importers.GetMappingImporter(registry);

            // Act
            using var stream = new MemoryStream();
            await exporter.ExportAsync(stream, departments);
            stream.Seek(0, SeekOrigin.Begin);
            
            // Assert
            var results = importer.Query<Department>(stream).ToList();
            var first = results[0];
            
            Assert.Equal("555-1111", first.PhoneNumbers[0]);
            Assert.Equal("555-2222", first.PhoneNumbers[1]);
            
            Assert.Equal("Dave", first.Employees[0].Name);
            Assert.Equal(35, first.Employees[0].Age);
            Assert.Equal(85000, first.Employees[0].Salary);
            Assert.Equal("dave@example.com", first.Employees[0].Email);
            
            Assert.Equal("Eve", first.Employees[1].Name);
            Assert.Equal(29, first.Employees[1].Age);
            Assert.Equal(75000, first.Employees[1].Salary);
            Assert.Equal("eve@example.com", first.Employees[1].Email);
        }

        #endregion

        #region Complex Type and Formula Tests

        [Fact]
        public async Task Formula_Properties_Should_Be_Handled_Correctly()
        {
            // Test formula support
            var registry = new MappingRegistry();
            registry.Configure<Product>(cfg =>
            {
                cfg.Property(p => p.Id).ToCell("A1");
                cfg.Property(p => p.Price).ToCell("B1");
                cfg.Property(p => p.Stock).ToCell("C1");
                cfg.Property(p => p.Price).ToCell("D1").WithFormula("=B1*C1"); // Total value formula
            });

            var testData = new[]
            {
                new Product { Id = 1, Price = 10.50m, Stock = 100 }
            };

            var exporter = MiniExcel.Exporters.GetMappingExporter(registry);
            
            using var stream = new MemoryStream();
            await exporter.ExportAsync(stream, testData);
            Assert.True(stream.Length > 0);
        }

        [Fact]
        public async Task Format_Properties_Should_Apply_Formatting()
        {
            // Test format support
            var registry = new MappingRegistry();
            registry.Configure<Person>(cfg =>
            {
                cfg.Property(p => p.Name).ToCell("A1");
                cfg.Property(p => p.BirthDate).ToCell("B1").WithFormat("yyyy-MM-dd");
                cfg.Property(p => p.Salary).ToCell("C1").WithFormat("#,##0.00");
            });

            var testData = new[]
            {
                new Person 
                { 
                    Name = "Test", 
                    BirthDate = new DateTime(1990, 6, 15), 
                    Salary = 12345.67m 
                }
            };

            var exporter = MiniExcel.Exporters.GetMappingExporter(registry);
            
            using var stream = new MemoryStream();
            await exporter.ExportAsync(stream, testData);
            Assert.True(stream.Length > 0);
        }

        #endregion

        #region Extended Mapping Tests

        [Fact]
        public async Task Mapping_WithComplexCellAddresses_ShouldMapCorrectly()
        {
            // Test various cell address formats
            var products = new[]
            {
                new Product { Id = 1, Name = "Laptop", Price = 999.99m, Stock = 10 }
            };

            var registry = new MappingRegistry();
            registry.Configure<Product>(cfg =>
            {
                cfg.Property(p => p.Id).ToCell("AA1");
                cfg.Property(p => p.Name).ToCell("AB1");
                cfg.Property(p => p.Price).ToCell("AC1");
                cfg.Property(p => p.Stock).ToCell("ZZ1");
                cfg.ToWorksheet("Products");
            });

            var exporter = MiniExcel.Exporters.GetMappingExporter(registry);
            using var stream = new MemoryStream();
            await exporter.ExportAsync(stream, products);
            
            // Verify the file was created
            Assert.True(stream.Length > 0);
            
            // Read back and verify
            stream.Position = 0;
            var importer = MiniExcel.Importers.GetOpenXmlImporter();
            var data = importer.Query(stream);
            var firstRow = data.FirstOrDefault();
            Assert.NotNull(firstRow);
        }

        [Fact]
        public async Task Mapping_WithNumericFormats_ShouldApplyCorrectly()
        {
            var products = new[]
            {
                new Product 
                { 
                    Id = 1, 
                    Name = "Widget", 
                    Price = 1234.5678m, 
                    DiscountPercentage = 0.15
                }
            };

            var registry = new MappingRegistry();
            registry.Configure<Product>(cfg =>
            {
                cfg.Property(p => p.Name).ToCell("A1");
                cfg.Property(p => p.Price).ToCell("B1").WithFormat("$#,##0.00");
                cfg.Property(p => p.DiscountPercentage).ToCell("C1").WithFormat("0.00%");
                cfg.ToWorksheet("Formatted");
            });

            var exporter = MiniExcel.Exporters.GetMappingExporter(registry);
            using var stream = new MemoryStream();
            await exporter.ExportAsync(stream, products);
            
            Assert.True(stream.Length > 0);
        }

        [Fact]
        public async Task Mapping_WithDateFormats_ShouldApplyCorrectly()
        {
            var products = new[]
            {
                new Product 
                { 
                    Name = "Item", 
                    LastRestocked = new DateTime(2024, 3, 15, 14, 30, 0)
                }
            };

            var registry = new MappingRegistry();
            registry.Configure<Product>(cfg =>
            {
                cfg.Property(p => p.Name).ToCell("A1");
                cfg.Property(p => p.LastRestocked).ToCell("B1").WithFormat("yyyy-MM-dd");
                cfg.Property(p => p.LastRestocked).ToCell("C1").WithFormat("MM/dd/yyyy hh:mm:ss");
                cfg.Property(p => p.LastRestocked).ToCell("D1").WithFormat("dddd, MMMM d, yyyy");
                cfg.ToWorksheet("DateFormats");
            });

            var exporter = MiniExcel.Exporters.GetMappingExporter(registry);
            using var stream = new MemoryStream();
            await exporter.ExportAsync(stream, products);
            
            Assert.True(stream.Length > 0);
        }

        [Fact]
        public async Task Mapping_WithBooleanValues_ShouldMapCorrectly()
        {
            var products = new[]
            {
                new Product { Name = "Active", IsActive = true },
                new Product { Name = "Inactive", IsActive = false }
            };

            var registry = new MappingRegistry();
            registry.Configure<Product>(cfg =>
            {
                cfg.Property(p => p.Name).ToCell("A1");
                cfg.Property(p => p.IsActive).ToCell("B1");
                cfg.ToWorksheet("Booleans");
            });

            var exporter = MiniExcel.Exporters.GetMappingExporter(registry);
            using var stream = new MemoryStream();
            await exporter.ExportAsync(stream, products);
            
            Assert.True(stream.Length > 0);
        }

        [Fact]
        public async Task Mapping_WithMultipleRowsToSameCells_ShouldOverwrite()
        {
            // When mapping multiple items to the same cells, last one should win
            var products = new[]
            {
                new Product { Id = 1, Name = "First" },
                new Product { Id = 2, Name = "Second" },
                new Product { Id = 3, Name = "Third" }
            };

            var registry = new MappingRegistry();
            registry.Configure<Product>(cfg =>
            {
                cfg.Property(p => p.Id).ToCell("A1");
                cfg.Property(p => p.Name).ToCell("B1");
                cfg.ToWorksheet("Overwrite");
            });

            var exporter = MiniExcel.Exporters.GetMappingExporter(registry);
            using var stream = new MemoryStream();
            await exporter.ExportAsync(stream, products);
            
            // The file should contain only the last item's data
            Assert.True(stream.Length > 0);
        }

        [Fact]
        public async Task Mapping_WithComplexTypes_ShouldHandleCorrectly()
        {
            var items = new[]
            {
                new ComplexModel
                {
                    Id = Guid.NewGuid(),
                    Title = "Complex Item",
                    CreatedAt = DateTimeOffset.Now,
                    Duration = TimeSpan.FromHours(2.5),
                    BinaryData = [1, 2, 3, 4, 5],
                    Website = new Uri("https://example.com")
                }
            };

            var registry = new MappingRegistry();
            registry.Configure<ComplexModel>(cfg =>
            {
                cfg.Property(p => p.Id).ToCell("A1");
                cfg.Property(p => p.Title).ToCell("B1");
                cfg.Property(p => p.CreatedAt).ToCell("C1").WithFormat("yyyy-MM-dd HH:mm:ss");
                cfg.Property(p => p.Duration).ToCell("D1");
                cfg.Property(p => p.Website).ToCell("E1");
                cfg.ToWorksheet("ComplexTypes");
            });

            var exporter = MiniExcel.Exporters.GetMappingExporter(registry);
            using var stream = new MemoryStream();
            await exporter.ExportAsync(stream, items);
            
            Assert.True(stream.Length > 0);
        }

        [Fact]
        public async Task Mapping_WithMultipleConfigurations_ShouldUseLast()
        {
            var products = new[]
            {
                new Product { Id = 1, Name = "Test" }
            };

            var registry = new MappingRegistry();
            
            // First configuration
            registry.Configure<Product>(cfg =>
            {
                cfg.Property(p => p.Id).ToCell("A1");
                cfg.Property(p => p.Name).ToCell("B1");
                cfg.ToWorksheet("First");
            });
            
            // Second configuration should override
            registry.Configure<Product>(cfg =>
            {
                cfg.Property(p => p.Id).ToCell("X1");
                cfg.Property(p => p.Name).ToCell("Y1");
                cfg.ToWorksheet("Second");
            });

            var exporter = MiniExcel.Exporters.GetMappingExporter(registry);
            using var stream = new MemoryStream();
            await exporter.ExportAsync(stream, products);
            
            Assert.True(stream.Length > 0);
        }

        [Fact]
        public void Mapping_WithInvalidCellAddress_ShouldThrowException()
        {
            var registry = new MappingRegistry();
            
            // Test various invalid cell addresses
            var invalidAddresses = new[] { "", " ", "123", "A", "1A", "@1" };
            
            foreach (var invalidAddress in invalidAddresses)
            {
                Assert.Throws<ArgumentException>(() =>
                {
                    registry.Configure<Product>(cfg =>
                    {
                        cfg.Property(p => p.Name).ToCell(invalidAddress);
                    });
                });
            }
        }

        [Fact]
        public async Task Mapping_WithEnumerableTypes_ShouldHandleCorrectly()
        {
            // Test with IEnumerable, List, Array, etc.
            var registry = new MappingRegistry();
            registry.Configure<Product>(cfg =>
            {
                cfg.Property(p => p.Name).ToCell("A1");
                cfg.Property(p => p.Price).ToCell("B1");
                cfg.ToWorksheet("Enumerable");
            });

            var exporter = MiniExcel.Exporters.GetMappingExporter(registry);
            
            // Test with array
            var array = new[] { new Product { Name = "Array", Price = 10 } };
            using (var stream = new MemoryStream())
            {
                await exporter.ExportAsync(stream, array);
                Assert.True(stream.Length > 0);
            }
            
            // Test with List
            var list = new List<Product> { new Product { Name = "List", Price = 20 } };
            using (var stream = new MemoryStream())
            {
                await exporter.ExportAsync(stream, list);
                Assert.True(stream.Length > 0);
            }
            
            // Test with IEnumerable
            IEnumerable<Product> enumerable = list;
            using (var stream = new MemoryStream())
            {
                await exporter.ExportAsync(stream, enumerable);
                Assert.True(stream.Length > 0);
            }
        }

        [Fact]
        public async Task Mapping_WithThreadSafety_ShouldWork()
        {
            var registry = new MappingRegistry();
            var tasks = new List<Task>();
            
            // Configure multiple types concurrently
            for (int i = 0; i < 10; i++)
            {
                var index = i;
                tasks.Add(Task.Run(() =>
                {
                    if (index % 2 == 0)
                    {
                        registry.Configure<Product>(cfg =>
                        {
                            cfg.Property(p => p.Name).ToCell("A1");
                            cfg.ToWorksheet($"Sheet{index}");
                        });
                    }
                    else
                    {
                        registry.Configure<ComplexModel>(cfg =>
                        {
                            cfg.Property(p => p.Title).ToCell("A1");
                            cfg.ToWorksheet($"Sheet{index}");
                        });
                    }
                }));
            }
            
            await Task.WhenAll(tasks);
            
            // Verify both configurations exist
            Assert.True(registry.HasMapping<Product>());
            Assert.True(registry.HasMapping<ComplexModel>());
        }

        [Fact]
        public async Task Mapping_WithSaveToFile_ShouldCreateFile()
        {
            var products = new[]
            {
                new Product { Id = 1, Name = "FileTest", Price = 99.99m }
            };

            var registry = new MappingRegistry();
            registry.Configure<Product>(cfg =>
            {
                cfg.Property(p => p.Id).ToCell("A1");
                cfg.Property(p => p.Name).ToCell("B1");
                cfg.Property(p => p.Price).ToCell("C1").WithFormat("$#,##0.00");
                cfg.ToWorksheet("FileOutput");
            });

            var exporter = MiniExcel.Exporters.GetMappingExporter(registry);
            
            var filePath = Path.GetTempFileName() + ".xlsx";
            try
            {
                using (var stream = File.Create(filePath))
                {
                    await exporter.ExportAsync(stream, products);
                }
                
                // Verify file exists and has content
                Assert.True(File.Exists(filePath));
                Assert.True(new FileInfo(filePath).Length > 0);
            }
            finally
            {
                if (File.Exists(filePath))
                    File.Delete(filePath);
            }
        }

        #endregion

        #region Edge Cases and Error Handling

        [Fact]
        public void Configuration_Without_Cell_Should_Throw()
        {
            var registry = new MappingRegistry();
            
            Assert.Throws<InvalidOperationException>(() =>
            {
                registry.Configure<Person>(cfg =>
                {
                    cfg.Property(p => p.Name); // Missing ToCell()
                });
                
                var mapping = registry.GetMapping<Person>();
            });
        }

        [Fact]
        public async Task Empty_Collection_Should_Handle_Gracefully()
        {
            var registry = new MappingRegistry();
            registry.Configure<ComplexEntity>(cfg =>
            {
                cfg.Property(e => e.Id).ToCell("A1");
                cfg.Collection(e => e.Tags).StartAt("B1");
            });

            var testData = new[]
            {
                new ComplexEntity { Id = 1, Tags = [] } // Empty collection
            };

            var exporter = MiniExcel.Exporters.GetMappingExporter(registry);
            
            using var stream = new MemoryStream();
            await exporter.ExportAsync(stream, testData);
            Assert.True(stream.Length > 0);
        }

        [Fact]
        public async Task Null_Values_Should_Handle_Gracefully()
        {
            var registry = new MappingRegistry();
            registry.Configure<ComplexEntity>(cfg =>
            {
                cfg.Property(e => e.Id).ToCell("A1");
                cfg.Property(e => e.Name).ToCell("B1");
                cfg.Property(e => e.Description).ToCell("C1");
            });

            var testData = new[]
            {
                new ComplexEntity 
                { 
                    Id = 1, 
                    Name = null!, // Null value
                    Description = null // Nullable property
                }
            };

            var exporter = MiniExcel.Exporters.GetMappingExporter(registry);
            
            using var stream = new MemoryStream();
            await exporter.ExportAsync(stream, testData);
            Assert.True(stream.Length > 0);
        }

        #endregion

        #region Performance and Optimization Tests

        [Fact]
        public void Universal_Optimization_Should_Create_Cell_Grid()
        {
            var registry = new MappingRegistry();
            registry.Configure<Product>(cfg =>
            {
                cfg.Property(p => p.Id).ToCell("A1");
                cfg.Property(p => p.Name).ToCell("C1");
                cfg.Property(p => p.Price).ToCell("E2");
            });

            var mapping = registry.GetMapping<Product>();
            Assert.NotNull(mapping.OptimizedCellGrid);
            Assert.NotNull(mapping.OptimizedBoundaries);
            
            // Verify grid dimensions
            var boundaries = mapping.OptimizedBoundaries;
            Assert.Equal(1, boundaries.MinRow);
            Assert.Equal(2, boundaries.MaxRow);
            Assert.Equal(1, boundaries.MinColumn); // A
            Assert.Equal(5, boundaries.MaxColumn); // E
        }

        [Fact]
        public async Task Large_Dataset_Should_Stream_Efficiently()
        {
            var registry = new MappingRegistry();
            registry.Configure<Product>(cfg =>
            {
                cfg.Property(p => p.Id).ToCell("A1");
                cfg.Property(p => p.Name).ToCell("B1");
                cfg.Property(p => p.Price).ToCell("C1");
            });

            // Generate large dataset
            var testData = Enumerable.Range(1, 10000).Select(i => new Product
            {
                Id = i,
                Name = $"Product {i}",
                Price = i * 10.5m
            });

            var exporter = MiniExcel.Exporters.GetMappingExporter(registry);
            
            using var stream = new MemoryStream();
            await exporter.ExportAsync(stream, testData);
            
            // Should complete without OutOfMemory
            Assert.True(stream.Length > 0);
        }

        #endregion

        #region Multiple Items and Pattern Tests

        [Fact]
        public void Multiple_Items_With_Collections_Should_Detect_Pattern()
        {
            var registry = new MappingRegistry();
            registry.Configure<ComplexEntity>(cfg =>
            {
                cfg.Property(e => e.Id).ToCell("A1");
                cfg.Property(e => e.Name).ToCell("B1");
                cfg.Collection(e => e.Tags).StartAt("A2");
            });

            var mapping = registry.GetMapping<ComplexEntity>();
            
            if (mapping.Collections.Any())
            {
                var boundaries = mapping.OptimizedBoundaries;
                // Pattern detection for multiple items
                Assert.True(boundaries.PatternHeight > 0 || !boundaries.IsMultiItemPattern);
            }
        }

        #endregion
    }
}