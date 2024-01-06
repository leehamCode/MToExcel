using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace MToExcel.Exceptons
{
    /// <summary>
    /// 长度设置异常
    /// </summary>
    public class RowColumnLenException : Exception
    {
        public RowColumnLenException(string? message) : base(message)
        {

        }

        
    }
}