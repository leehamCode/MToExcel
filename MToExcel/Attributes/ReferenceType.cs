using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MToExcel.Attributes
{
    /// <summary>
    /// MyConverter的功能已经可以将泛型对象中的基础数据类型成功的打印到对应的Excel之中
    /// 现在要思考如何将泛型中的引用数据类型也打印到Excel中
    /// 
    /// PS:注意:引用数据类型里面可能也有引用数据类型,这里可能造成多层的引用(这并不可怕),烦人的是可能会存在循环引用的问题,
    /// 导致永远打印不出正确的结果,
    /// 
    /// 这个初始版本暂时不设计解决循环依赖的问题
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = true)]
    public class ReferenceType : Attribute   //这个标记将属性标记为引用类型
    {
        //是否将引用类型再按属性拆分为多个列展示
        private bool isMultiColumn = true;

        public ReferenceType(bool isMultiColumn)
        {
            this.isMultiColumn = isMultiColumn;
        }

        public bool getIsMultiPart()
        {
            return this.isMultiColumn;
        }
    }

}
