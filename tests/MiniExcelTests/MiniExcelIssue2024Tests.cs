using MiniExcelLibs.Attributes;
using MiniExcelLibs.OpenXml;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.IO;
using System.Linq;
using Xunit;
using Xunit.Abstractions;

namespace MiniExcelLibs.Tests
{
    public partial class MiniExcelIssue2024Tests
    {
        private readonly ITestOutputHelper output;

        public MiniExcelIssue2024Tests(ITestOutputHelper output)
        {
            this.output = output;
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
                    new[] {
                        new DynamicExcelColumn("long2") { Format = "@", Width = 25 }
                    }
            };
            var value = new[] { new { long2 = "1550432695793487872" } };
            MiniExcel.SaveAs(path, value, configuration: config);
        }
    }
}