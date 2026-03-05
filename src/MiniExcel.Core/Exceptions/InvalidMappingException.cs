namespace MiniExcelLib.Core.Exceptions;

public class InvalidMappingException(string message, Type type, MemberInfo? member = null)
    : InvalidOperationException(message)
{
    public Type CausingType { get; } = type;
    public MemberInfo? CausingProperty { get; } = member;
}