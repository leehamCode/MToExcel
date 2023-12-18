using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using MToExcel.Models.Enums;

namespace MToExcel.Models.Param
{
    /// <summary>
    /// 边框参数对象
    /// </summary>
    public class BorderSet
    {
        public BorderWid Wid{get;set;}

        public short Color{get;set;}

        public BorderDirect Direct{get;set;}
    }
}