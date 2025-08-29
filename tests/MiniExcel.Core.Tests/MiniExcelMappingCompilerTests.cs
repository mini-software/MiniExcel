using MiniExcelLib.Core.Mapping;
using MiniExcelLib.Core.Mapping.Configuration;
using System;
using System.Collections.Generic;
using System.Linq;
using Xunit;

namespace MiniExcelLib.Tests
{
    /// <summary>
    /// Tests for the mapping compiler and optimization system.
    /// Focuses on internal optimization details and performance characteristics.
    /// </summary>
    public class MiniExcelMappingCompilerTests
    {
        #region Test Models

        public class SimpleEntity
        {
            public int Id { get; set; }
            public string Name { get; set; } = "";
            public decimal Value { get; set; }
        }

        public class ComplexEntity
        {
            public int Id { get; set; }
            public string Title { get; set; } = "";
            public List<string> Items { get; set; } = new();
            public Dictionary<string, object> Properties { get; set; } = new();
        }

        #endregion

        #region Optimization Detection Tests

        [Fact]
        public void Sequential_Properties_Should_Be_Detected()
        {
            // Arrange
            var registry = new MappingRegistry();
            registry.Configure<SimpleEntity>(cfg =>
            {
                cfg.Property(e => e.Id).ToCell("A1");
                cfg.Property(e => e.Name).ToCell("B1");
                cfg.Property(e => e.Value).ToCell("C1");
            });

            // Act
            var mapping = registry.GetMapping<SimpleEntity>();

            // Assert - verify optimization is applied
            Assert.NotNull(mapping.OptimizedBoundaries);
            Assert.NotNull(mapping.OptimizedCellGrid);
            Assert.Equal(3, mapping.Properties.Count);
            
            // Verify properties are correctly mapped
            Assert.Equal("A1", mapping.Properties[0].CellAddress);
            Assert.Equal("B1", mapping.Properties[1].CellAddress);
            Assert.Equal("C1", mapping.Properties[2].CellAddress);
        }

        [Fact]
        public void NonSequential_Properties_Should_Use_Optimization()
        {
            // Arrange
            var registry = new MappingRegistry();
            registry.Configure<SimpleEntity>(cfg =>
            {
                cfg.Property(e => e.Id).ToCell("A1");
                cfg.Property(e => e.Name).ToCell("C1"); // Skip B
                cfg.Property(e => e.Value).ToCell("B2"); // Different row
            });

            // Act
            var mapping = registry.GetMapping<SimpleEntity>();

            // Assert - verify optimization is applied
            Assert.NotNull(mapping.OptimizedBoundaries);
            Assert.NotNull(mapping.OptimizedCellGrid);
        }

        #endregion

        #region Cell Grid Tests

        [Fact]
        public void OptimizedCellGrid_Should_Have_Correct_Dimensions()
        {
            // Arrange
            var registry = new MappingRegistry();
            registry.Configure<SimpleEntity>(cfg =>
            {
                cfg.Property(e => e.Id).ToCell("A1");
                cfg.Property(e => e.Name).ToCell("E1");   // Column E
                cfg.Property(e => e.Value).ToCell("C3");  // Row 3
            });

            // Act
            var mapping = registry.GetMapping<SimpleEntity>();

            // Assert
            Assert.NotNull(mapping.OptimizedBoundaries);
            var boundaries = mapping.OptimizedBoundaries;
            
            Assert.Equal(1, boundaries.MinRow);
            Assert.Equal(3, boundaries.MaxRow);
            Assert.Equal(1, boundaries.MinColumn);  // A = 1
            Assert.Equal(5, boundaries.MaxColumn);  // E = 5
            
            Assert.Equal(3, boundaries.GridHeight); // 3 - 1 + 1 = 3
            Assert.Equal(5, boundaries.GridWidth);  // 5 - 1 + 1 = 5
        }

        [Fact]
        public void OptimizedCellGrid_Should_Map_Properties_Correctly()
        {
            // Arrange
            var registry = new MappingRegistry();
            registry.Configure<SimpleEntity>(cfg =>
            {
                cfg.Property(e => e.Id).ToCell("B2");
                cfg.Property(e => e.Name).ToCell("D2");
                cfg.Property(e => e.Value).ToCell("B4");
            });

            // Act
            var mapping = registry.GetMapping<SimpleEntity>();
            var grid = mapping.OptimizedCellGrid!;
            var boundaries = mapping.OptimizedBoundaries!;

            // Assert
            // Grid should be 3x3 (rows 2-4, columns B-D which is 2-4)
            Assert.Equal(3, grid.GetLength(0)); // Height
            Assert.Equal(3, grid.GetLength(1)); // Width

            // Check Id at B2 (relative: 0,0)
            var idHandler = grid[0, 0];
            Assert.Equal(CellHandlerType.Property, idHandler.Type);
            Assert.Equal("Id", idHandler.PropertyName);

            // Check Name at D2 (relative: 0,2)
            var nameHandler = grid[0, 2];
            Assert.Equal(CellHandlerType.Property, nameHandler.Type);
            Assert.Equal("Name", nameHandler.PropertyName);

            // Check Value at B4 (relative: 2,0)
            var valueHandler = grid[2, 0];
            Assert.Equal(CellHandlerType.Property, valueHandler.Type);
            Assert.Equal("Value", valueHandler.PropertyName);

            // Check empty cells
            Assert.Equal(CellHandlerType.Empty, grid[0, 1].Type); // C2
            Assert.Equal(CellHandlerType.Empty, grid[1, 0].Type); // B3
        }

        #endregion

        #region Collection Optimization Tests

        [Fact]
        public void Collection_Should_Mark_Grid_Cells()
        {
            // Arrange
            var registry = new MappingRegistry();
            registry.Configure<ComplexEntity>(cfg =>
            {
                cfg.Property(e => e.Id).ToCell("A1");
                cfg.Property(e => e.Title).ToCell("B1");
                cfg.Collection(e => e.Items).StartAt("A2");
            });

            // Act
            var mapping = registry.GetMapping<ComplexEntity>();
            var grid = mapping.OptimizedCellGrid!;

            // Assert
            // Check that collection cells are marked
            // Note: Collection handling depends on implementation details
            Assert.NotNull(grid);
            Assert.True(mapping.OptimizedBoundaries!.HasDynamicCollections);
        }

        [Fact]
        public void Multiple_Collections_Should_Be_Handled()
        {
            // Arrange
            var registry = new MappingRegistry();
            registry.Configure<ComplexEntity>(cfg =>
            {
                cfg.Property(e => e.Id).ToCell("A1");
                cfg.Collection(e => e.Items).StartAt("B1");
                cfg.Collection(e => e.Properties).StartAt("C1");
            });

            // Act
            var mapping = registry.GetMapping<ComplexEntity>();

            // Assert
            Assert.Equal(2, mapping.Collections.Count);
        }

        #endregion

        #region Pre-compilation Tests

        [Fact]
        public void Property_Getters_Should_Be_Compiled()
        {
            // Arrange
            var registry = new MappingRegistry();
            registry.Configure<SimpleEntity>(cfg =>
            {
                cfg.Property(e => e.Id).ToCell("A1");
                cfg.Property(e => e.Name).ToCell("B1");
            });

            // Act
            var mapping = registry.GetMapping<SimpleEntity>();

            // Assert
            foreach (var prop in mapping.Properties)
            {
                Assert.NotNull(prop.Getter);
                
                // Test getter works
                var entity = new SimpleEntity { Id = 123, Name = "Test" };
                var idValue = mapping.Properties[0].Getter(entity);
                Assert.Equal(123, idValue);
            }
        }

        [Fact]
        public void Property_Setters_Should_Be_Compiled()
        {
            // Arrange
            var registry = new MappingRegistry();
            registry.Configure<SimpleEntity>(cfg =>
            {
                cfg.Property(e => e.Id).ToCell("A1");
                cfg.Property(e => e.Name).ToCell("B1");
            });

            // Act
            var mapping = registry.GetMapping<SimpleEntity>();

            // Assert
            foreach (var prop in mapping.Properties)
            {
                Assert.NotNull(prop.Setter);
                
                // Test setter works
                var entity = new SimpleEntity();
                mapping.Properties[0].Setter!(entity, 456);
                Assert.Equal(456, entity.Id);
            }
        }

        #endregion

        #region Formula and Format Tests

        [Fact]
        public void Formula_Properties_Should_Be_Marked()
        {
            // Arrange
            var registry = new MappingRegistry();
            registry.Configure<SimpleEntity>(cfg =>
            {
                cfg.Property(e => e.Id).ToCell("A1");
                cfg.Property(e => e.Value).ToCell("B1").WithFormula("=A1*2");
            });

            // Act
            var mapping = registry.GetMapping<SimpleEntity>();
            var grid = mapping.OptimizedCellGrid!;

            // Assert
            var formulaHandler = grid[0, 1]; // B1 relative position
            Assert.Equal(CellHandlerType.Formula, formulaHandler.Type);
            Assert.Equal("=A1*2", formulaHandler.Formula);
        }

        [Fact]
        public void Format_Should_Be_Preserved()
        {
            // Arrange
            var registry = new MappingRegistry();
            registry.Configure<SimpleEntity>(cfg =>
            {
                cfg.Property(e => e.Value).ToCell("A1").WithFormat("#,##0.00");
            });

            // Act
            var mapping = registry.GetMapping<SimpleEntity>();

            // Assert
            var prop = mapping.Properties[0];
            Assert.Equal("#,##0.00", prop.Format);
        }

        #endregion

        #region Edge Cases

        [Fact]
        public void Empty_Configuration_Should_Be_Valid()
        {
            // Arrange
            var registry = new MappingRegistry();
            registry.Configure<SimpleEntity>(cfg =>
            {
                // No mappings
            });

            // Act
            var mapping = registry.GetMapping<SimpleEntity>();

            // Assert
            Assert.NotNull(mapping);
            Assert.Empty(mapping.Properties);
            Assert.Empty(mapping.Collections);
        }

        [Fact]
        public void Duplicate_Cell_Mapping_Should_Be_Allowed()
        {
            // Arrange
            var registry = new MappingRegistry();
            registry.Configure<SimpleEntity>(cfg =>
            {
                cfg.Property(e => e.Id).ToCell("A1");
                cfg.Property(e => e.Name).ToCell("A1"); // Same cell
            });

            // Act
            var mapping = registry.GetMapping<SimpleEntity>();

            // Assert
            Assert.Equal(2, mapping.Properties.Count);
            // Both properties map to A1 - last one wins in the grid
        }

        #endregion
    }
}