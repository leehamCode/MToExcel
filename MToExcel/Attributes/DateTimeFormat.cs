using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace MToExcel.Attributes
{
    [AttributeUsage(AttributeTargets.Property|AttributeTargets.Class,AllowMultiple =false,Inherited =false)]
    public class DateTimeFormat:Attribute
    {
        /// <summary>
        /// 制定日期格式
        /// </summary>
        /// <value></value>
        public string format{get;set;} = "yyyy/MM/dd HH:mm:ss";



    }
}