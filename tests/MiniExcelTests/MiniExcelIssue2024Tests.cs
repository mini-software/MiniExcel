using MiniExcelLibs.Attributes;
using MiniExcelLibs.OpenXml;
using MiniExcelLibs.Tests.Utils;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Xunit;
using Xunit.Abstractions;

namespace MiniExcelLibs.Tests
{
    public partial class MiniExcelIssue2024_2025_Tests
    {
        private readonly ITestOutputHelper output;

        public MiniExcelIssue2024_2025_Tests(ITestOutputHelper output)
        {
            this.output = output;
        }

        /// <summary>
        /// https://github.com/mini-software/MiniExcel/issues/750
        /// </summary>
        [Fact]
        public void TestIssue20250403_SaveAsByTemplate_OPT()
        {
            long memoryBefore = GC.GetTotalMemory(true);
            {
                var path = PathHelper.GetTempFilePath();
                var templatePath = PathHelper.GetFile("xlsx/TestIssue20250403_SaveAsByTemplate_OPT.xlsx");
                var data = new Dictionary<string, object>
                {
                    ["list"] = Enumerable.Range(0, 1000000).Select(s => new { value1 = Guid.NewGuid(), value2 = Guid.NewGuid(), })
                };
                MiniExcel.SaveAsByTemplate(path, templatePath, data);
            }
            long memoryAfter = GC.GetTotalMemory(true);
            long memoryIncrease = memoryAfter - memoryBefore;
            Assert.True(memoryIncrease < 5318168);
        }

        /// <summary>
        /// https://github.com/mini-software/MiniExcel/issues/627
        /// </summary>
        [Fact]
        public void TestIssue627()
        {
            var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.xlsx");
            var config = new OpenXmlConfiguration
            {
                AutoFilter = false,
                DynamicColumns =
                [
                    new DynamicExcelColumn("long2") { Format = "@", Width = 25 }
                ]
            };
            var value = new[] { new { long2 = "1550432695793487872" } };
            var rowsWritten = MiniExcel.SaveAs(path, value, configuration: config);
            Assert.Single(rowsWritten);
            Assert.Equal(1, rowsWritten[0]);
        }
    }
}