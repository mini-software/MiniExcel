namespace MiniExcelLibs.Utils
{
    using System.Reflection;

    internal static class ReflactionExtension {
        internal static bool IsSupportSetMethod(this PropertyInfo propertyInfo) {
            return propertyInfo.GetSetMethod() != null;
        }
    }

}
