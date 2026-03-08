namespace MiniExcelLib.Core.Exceptions;

public class InvalidMappingException(string message, Type? type, MemberInfo? member = null)
    : InvalidOperationException(message)
{
    public Type? InvalidType { get; } = type;
    public MemberInfo? InvalidProperty { get; } = member;
}