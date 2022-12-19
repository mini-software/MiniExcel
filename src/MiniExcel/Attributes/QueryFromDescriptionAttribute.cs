using System;

namespace MiniExcelLibs.Attributes
{
    [AttributeUsage(AttributeTargets.Property | AttributeTargets.Field, AllowMultiple = false)]
    public class QueryFromDescriptionAttribute : Attribute
    {
    }
}