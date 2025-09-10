using MiniExcelLibs.Attributes;
using System.Globalization;

namespace MiniExcelLibs
{
    public interface IConfiguration { }
    public abstract class Configuration : IConfiguration
    {
        public CultureInfo Culture { get; set; } = CultureInfo.InvariantCulture;
        public DynamicExcelColumn[] DynamicColumns { get; set; }
        public int BufferSize { get; set; } = 1024 * 512;
        public bool FastMode { get; set; } = false;
        
        /// <summary>
        ///     When exporting using DataReader, the data not in DynamicColumn will be filtered.
        /// </summary>
        public bool DynamicColumnFirst { get; set; } = false;
        
        /// <summary>
        /// Specifies when and how DateTime values are converted to DateOnly values.
        /// </summary>
        public DateOnlyConversionMode DateOnlyConversionMode { get; set; }
    }
    
    public enum DateOnlyConversionMode
    {
        /// <summary>
        /// No conversion is applied and DateOnly values are not transformed.
        /// </summary>
        None,

        /// <summary>
        /// Allows conversion from DateTime to DateOnly only if the time component is exactly midnight (00:00:00).
        /// </summary>
        RequireMidnight,

        /// <summary>
        /// Converts DateTime to DateOnly by ignoring the time part completely, assuming the time component is not critical.
        /// </summary>
        IgnoreTimePart
    }
}
