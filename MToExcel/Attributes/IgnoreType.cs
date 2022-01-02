using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MToExcel.Attributes
{
    /// <summary>
    /// 忽略制定标记的属性值
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false,Inherited =true)]
    public class IgnoreType:Attribute
    {

        public IgnoreType()
        {

        }
    }
}
