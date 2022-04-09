using CsvHelper;
using MiniExcelLibs.Tests.Utils;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;

namespace MiniExcelLibs.Tests
{
    public class PerformanceEnhancementTest
    {
        [Fact]
        public void StartsWithPerformanceCompare()
        {
            string stringVal = "@@@fileid@@@,tteffadasdas ", findVal = "@@@fileid@@@,";
            const int LOOP = 5000000;
            int runTimes = 0;
            long normalMilliseconds = 0;
            long ordinalMilliseconds = 0;

            Stopwatch watch = Stopwatch.StartNew();
            for (int i = 0; i < LOOP; i++)
            {
                if (stringVal.StartsWith(findVal)) 
                    runTimes++;
            }
            watch.Stop();
            normalMilliseconds = watch.ElapsedMilliseconds;
            Assert.Equal(LOOP, runTimes);

            runTimes = 0;
            watch = Stopwatch.StartNew();
            for (int i = 0; i < LOOP; i++)
            {
                if (stringVal.StartsWith(findVal,StringComparison.Ordinal)) 
                    runTimes++;
            }
            watch.Stop();
            ordinalMilliseconds = watch.ElapsedMilliseconds;

            Assert.Equal(LOOP, runTimes);
            Assert.True(normalMilliseconds > ordinalMilliseconds);
        }

        [Fact]
        public void EndWithPerformanceCompare()
        {
            string stringVal = "@@@fileid@@@,tteffadasdas ", findVal = " ";
            const int LOOP = 5000000;
            int runTimes = 0;
            long normalMilliseconds = 0;
            long ordinalMilliseconds = 0;

            Stopwatch watch = Stopwatch.StartNew();
            for (int i = 0; i < LOOP; i++)
            {
                if (stringVal.EndsWith(findVal)) 
                    runTimes++;
            }
            watch.Stop();
            normalMilliseconds = watch.ElapsedMilliseconds;
            Assert.Equal(LOOP, runTimes);

            runTimes = 0;
            watch = Stopwatch.StartNew();
            for (int i = 0; i < LOOP; i++)
            {
                if (stringVal.EndsWith(findVal,StringComparison.Ordinal)) 
                    runTimes++;
            }
            watch.Stop();
            ordinalMilliseconds = watch.ElapsedMilliseconds;

            Assert.Equal(LOOP, runTimes);
            Assert.True(normalMilliseconds > ordinalMilliseconds);
        }
    }
}