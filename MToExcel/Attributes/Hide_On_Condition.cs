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
        public Hide_On_Condition(string rowCondition, string colCondition)
        {
            this.rowCondition = rowCondition;
            this.colCondition = colCondition;
        }

        public string rowCondition {get;set;} = "non";

       public string colCondition {get;set;} = "non";
       
        
    }
}