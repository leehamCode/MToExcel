using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace MToExcel.Exceptons
{
    /// <summary>
    /// Rgb颜色
    /// </summary>
    public class RgbArrayException : Exception
    {
        public RgbArrayException(string? message) : base(message)
        {
        }
    }
}