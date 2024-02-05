using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace MToExcel.Attributes
{
    /// <summary>
    /// 是否合并大块相同值的单元格
    /// </summary>
    [AttributeUsage(AttributeTargets.Class|AttributeTargets.Property,Inherited = false, AllowMultiple =false)]
    public class MergeNearEqualBox:Attribute
    {
        /// <summary>
        /// 只合并同一列
        /// </summary>
        /// <value></value>
        public bool OnlyCol{get;set;}=true;


        public bool OnlyRow{get;set;}=false;

        public bool Both{get;set;}
    }
}