using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace MToExcel.Attributes
{
    /// <summary>
    /// 冻结区域设置,
    /// 参数为冻结开始的行/列
    /// </summary>
    [AttributeUsage(AttributeTargets.Class,Inherited =false,AllowMultiple =false)]
    public class FreezeAreaAttribute:Attribute
    {
        public int FreezeStartRow{get;set;}

        public int FreezeStartCol{get;set;}
    }
}