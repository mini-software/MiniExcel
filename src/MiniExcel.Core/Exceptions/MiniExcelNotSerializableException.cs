namespace MiniExcelLib.Core.Exceptions;

public class MiniExcelNotSerializableException(string message, MemberInfo member) 
    : InvalidOperationException(message)
{
    public MemberInfo Member { get; } = member;
}