using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MiniExcelLibs.Attributes;
using MiniExcelLibs.OpenXml;
using MiniExcelLibs.Tests.Utils;
using Xunit;

namespace MiniExcelLibs.Tests
{
    public class MiniExcelOpenXmlConfigurationTest
    {
        [Fact]
        public void EnableWriteFilePathTest()
        {
            var img = new HttpClient().GetByteArrayAsync("https://user-images.githubusercontent.com/12729184/150462383-ad9931b3-ed8d-4221-a1d6-66f799743433.png").Result;
            var value = new[] {
                new ImgExportTestDto{ Name="github",Img=File.ReadAllBytes(PathHelper.GetFile("images/github_logo.png"))},
                new ImgExportTestDto{ Name="google",Img=File.ReadAllBytes(PathHelper.GetFile("images/google_logo.png"))},
                new ImgExportTestDto{ Name="microsoft",Img=File.ReadAllBytes(PathHelper.GetFile("images/microsoft_logo.png"))},
                new ImgExportTestDto{ Name="reddit",Img=File.ReadAllBytes(PathHelper.GetFile("images/reddit_logo.png"))},
                new ImgExportTestDto{ Name="statck_overflow",Img=File.ReadAllBytes(PathHelper.GetFile("images/statck_overflow_logo.png"))},
                new ImgExportTestDto{ Name="statck_over",Img=img},
            };
            var path=PathHelper.GetFile("Test_EnableWriteFilePath.xlsx");
            MiniExcel.SaveAs(path, value, configuration: new OpenXmlConfiguration() { EnableWriteFilePath=false},overwriteFile:true);
            Assert.True(File.Exists(path));

            var rows = MiniExcel.Query<ImgExportTestDto>(path).ToList();
            Assert.True( rows.All(x => x.Img is null || x.Img.Length < 1));
        }


    }

    class ImgExportTestDto
    {
        public string Name { get; set; }
        [ExcelColumn(Name = "图片",Width = 100)]
        public byte[] Img { get; set; }
    }
}
