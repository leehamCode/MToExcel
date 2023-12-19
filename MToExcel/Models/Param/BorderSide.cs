using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace MToExcel.Models.Param
{
    /// <summary>
    /// 便设置样式的集合
    /// </summary>
    public class BorderSide
    {
        /// <summary>
        /// 各个边需要的样式
        /// </summary>
        /// <value></value>
        public HashSet<BorderSet>? Sides {get;set;}

        /// <summary>
        /// 描述
        /// </summary>
        /// <value></value>
        public string? Desctipt{get;set;}

        // override object.Equals
        public override bool Equals(object obj)
        {
            var target =  (BorderSide)obj;

            if(target.Desctipt==Desctipt&&target.Sides.Equals(Sides))
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        
        // override object.GetHashCode
        public override int GetHashCode()
        {
           return HashCode.Combine(Desctipt,Sides);
        }
        
    }
}