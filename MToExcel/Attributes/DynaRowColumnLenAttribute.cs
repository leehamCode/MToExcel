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

        public DynaRowColumnLenAttribute(double? rowHeight, double? colLength)
        {
            if(rowHeight!=null)
            {
                RowHeight = rowHeight.Value;
            }

            if(colLength!=null)
            {
                ColLength = colLength.Value;
            }
        }

        


        public double RowHeight{get;set;} = -1.00;

        public double ColLength{get;set;} = -1.00;

        


    }
}