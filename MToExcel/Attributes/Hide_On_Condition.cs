using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace MToExcel.Attributes
{
    /// <summary>
    /// 是否按条件隐藏一些样式
    /// </summary>
    [AttributeUsage(AttributeTargets.Class,Inherited =false,AllowMultiple =false)]
    public class Hide_On_Condition:Attribute
    {
        public Hide_On_Condition(string condition)
        {
            this.condition = condition;
        }

        public string condition{get;set;} = "non";

        
    }
}