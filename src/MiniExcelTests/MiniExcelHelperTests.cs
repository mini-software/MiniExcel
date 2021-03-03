using Xunit;
using MiniExcel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace MiniExcel.Tests
{
    public class MiniExcelHelperTests
    {
        [Fact()]
        public void CreateTest()
        {
            var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.xlsx");
            MiniExcelHelper.Create(path, new[] {
                  new { a = @"""<>+-*//}{\\n", b = 1234567890,c = true,d=DateTime.Now },
                  new { a = "<test>Hello World</test>", b = -1234567890,c=false,d=DateTime.Now.Date}
             });
            var info = new FileInfo(path);
            
            Assert.True(info.FullName == path);

            File.Delete(path);
        }
    }
}