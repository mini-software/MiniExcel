namespace MiniExcelLib.Core.Exceptions;

public class MemberNotSerializableException(string message, MemberInfo member) 
    : InvalidOperationException(message)
{
    public MemberInfo Member { get; } = member;
}