using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace MToExcel.Converter
{
    /// <summary>
    /// 表达式异常
    /// </summary>
    public class ExpressionException : Exception
    {
        public ExpressionException(string? message) : base(message)
        {
        }
    }
}