using MiniExcelLibs;
using System;
using System.IO;
using Xunit;

namespace MiniExcelTests
{
    public class MiniExcelInsertSheetTests
    {
        [Fact]
        public async void InsertSheetTestAsync()
        {
            //    if (File.Exists(@"C:\Users\huangzhenhua\Desktop\1.xlsx"))
            //    {
            //        File.Delete(@"C:\Users\huangzhenhua\Desktop\1.xlsx");
            //    }
            //    File.Copy(@"C:\Users\huangzhenhua\Desktop\3 - 副本.xlsx", @"C:\Users\huangzhenhua\Desktop\1.xlsx");
            //    var dt = await MiniExcelLibs.MiniExcel.QueryAsDataTableAsync(@"C:\Users\huangzhenhua\Desktop\1.xlsx", true);

            //    MiniExcelLibs.MiniExcel.InsertSheet(@"C:\Users\huangzhenhua\Desktop\1.xlsx", dt, sheetName: "Sheet3", overwriteSheet: true);

            string filePath = Path.Combine(@"C:\Users\huangzhenhua\Desktop\1.csv");
            var objList = new[] {
                      new { ID=1,Name ="Frank",InDate=new DateTime(2021,06,07)},
                      new { ID=2,Name ="Gloria",InDate=new DateTime(2022,05,03)},
                };
            using (var stream = File.OpenWrite(filePath))
            {
                MiniExcelLibs.MiniExcel.SaveAs(stream, objList, true, "data", ExcelType.CSV);
            }
            objList = new[] {
                  new { ID=3,Name ="Frank",InDate=new DateTime(2021,06,07)},
                  new { ID=4,Name ="Gloria",InDate=new DateTime(2022,05,03)},
            };
            using (var stream = File.OpenWrite(filePath))
            {
                MiniExcelLibs.MiniExcel.Insert(stream, objList, "data", ExcelType.CSV);
            }
        }
    }
}
