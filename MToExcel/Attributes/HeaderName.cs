using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MToExcel.Attributes
{
    /// <summary>
    /// HeaderName标签，当不想以属性名为表头信息的时候，可以使用此标签来自己指定
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false,Inherited =true)]
    public class HeaderName:Attribute
    {
        /// <summary>
        /// 指定表头列名
        /// </summary>
        private string name;

        public HeaderName(string name)
        {
            this.name = name;
        }

        public string getCustomProName()
        {
            return name;
        }
    }
}
