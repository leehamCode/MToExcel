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
        
        //行隐藏
        public string rowCondition {get;set;} = "non";


        //列隐藏,现无法支持
        public string colCondition {get;set;} = "non";
       
        
    }
}