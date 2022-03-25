using System;
using System.Collections.Generic;
using MiniExcelLibs;
using Xunit;

namespace MiniExcelTests
{
    public class DbListTests
    {
        [Fact]
        public void TestDbListCount()
        {
            var dbList = new DbList(Guid.NewGuid().ToString());
            Assert.Equal(0, dbList.Count);
            dbList.Add("test");
            Assert.Equal(1, dbList.Count);
            dbList.AddRange(new List<string>() { "test1", "test2" });
            Assert.Equal(3, dbList.Count);
            dbList.Remove("test");
            Assert.Equal(2, dbList.Count);
            dbList.Insert(0, "test");
            Assert.Equal(3, dbList.Count);
            dbList.RemoveAt(0);
            Assert.Equal(2, dbList.Count);
            dbList.Clear();
            Assert.Equal(0, dbList.Count);
            dbList.Dispose();
        }
    }
}