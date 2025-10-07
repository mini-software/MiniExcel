using System;
using System.Reflection;

namespace MiniExcelLibs.Exceptions
{
    public class MiniExcelNotSerializableException : InvalidOperationException
    {
        public MemberInfo Member { get; }
        
        public MiniExcelNotSerializableException(string message, MemberInfo member) :  base(message)
        {
            Member = member;
        }
    }
}
