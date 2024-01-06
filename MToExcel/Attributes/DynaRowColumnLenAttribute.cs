using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace MToExcel.Attributes
{
    /// <summary>
    /// 动态的格子长宽设置
    /// </summary>
    [AttributeUsage(AttributeTargets.Property,Inherited =false,AllowMultiple = false)]
    public class DynaRowColumnLenAttribute:Attribute
    {

        public DynaRowColumnLenAttribute(double rowHeight, double colLength)
        {
            if(rowHeight!=0)
            {
                RowHeight = rowHeight;
            }

            if(colLength!=0)
            {
                ColLength = colLength;
            }
        }

        /// <summary>
        /// 单个设置单元格的长宽
        /// </summary>
        /// <param name="RowOrCol_Len"></param>
        /// <param name="length">长度单位</param>
        public DynaRowColumnLenAttribute(bool RowOrCol_Len,double length)
        {
            if(RowOrCol_Len)
            {
                RowHeight = length;
            }
            else
            {
                ColLength = length;
            }
        }

       

        public double RowHeight{get;set;} = 1.14514;

        public double ColLength{get;set;} = 1.14514;

        


    }
}