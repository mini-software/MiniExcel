using System;

namespace MiniExcelNumberFormat
{
    internal static class CompatibleConvert
    {
        /// <summary>
        /// A backward-compatible version of <see cref="Convert.ToString(object, IFormatProvider)"/>.
        /// Starting from .net Core 3.0 the default precision used for formatting floating point number has changed.
        /// To always format numbers the same way, no matter what version of runtime is used, we specify the precision explicitly.
        /// </summary>
        public static string ToString(object value, IFormatProvider provider)
        {
            switch (value)
            {
                case double d:
                    return d.ToString("G15", provider);
                case float f:
                    return f.ToString("G7", provider);
                default:
                    return Convert.ToString(value, provider);
            }
        }
    }
}