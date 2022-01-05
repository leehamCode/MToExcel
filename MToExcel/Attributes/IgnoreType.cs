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
        //是否真的要真的隐藏该属性
        private bool isTrueIgnore = true;


        public IgnoreType()
        {

        }

        public IgnoreType(bool isTrueIgnore)
        {
            this.isTrueIgnore = isTrueIgnore;
        }

        public bool IsTrueIgnoreIt()
        {
            return isTrueIgnore;
        }
    }
}
