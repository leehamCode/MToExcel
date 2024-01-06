using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace MToExcel.Exceptons
{
    public class CustomHeadException : Exception
    {
        public CustomHeadException(string? message) : base(message)
        {
        }
    }
}