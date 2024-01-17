using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace MToExcel.Exceptons
{
    public class TitleAttrException : Exception
    {
        public TitleAttrException(string? message) : base(message)
        {
        }
    }
}