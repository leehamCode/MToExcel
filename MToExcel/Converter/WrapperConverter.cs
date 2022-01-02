using MToExcel.Attributes;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace MToExcel.Converter
{
    /// <summary>
    /// 这是一个基本转化器的包装类,Design By Description Mode
    /// 
    /// </summary>
    public class WrapperConverter
    {
        //类型池,用来保存打了Attribute的标签的自定义类型
        public static Dictionary<Type, ReferenceType> TypePool = new Dictionary<Type, ReferenceType>();

        /// <summary>
        /// taget对象,用来具体的打表
        /// </summary>
        public BasicConverter basic = null;

        public WrapperConverter()
        {
            basic = new BasicConverter();

        }

        public IWorkbook ConvertToExcel<T>(List<T> list)
        {
            //在调target执行真正的方法之前,在此可以设置增强方法 advice method
            CheckAttribute(typeof(T));

            return basic.ConvertToExcel(list);
        }

        /// <summary>
        /// (增强方法)检查泛型类型的属性中是否带有自定义的Reference的标签--check
        /// </summary>
        /// <param name="type"></param>
        public void CheckAttribute(Type type)
        {
            PropertyInfo[] pros = type.GetProperties();

            foreach (PropertyInfo pro in pros)
            {
                //遍历获取属性中的Attribute对象
                ReferenceType refer = (ReferenceType)pro.GetCustomAttribute(typeof(ReferenceType));
                if (refer != null)
                {
                    //将打了标记的类型和标记本身放到类型池中
                    TypePool.Add(pro.PropertyType, refer);
                }
            }

        }
    }
}

