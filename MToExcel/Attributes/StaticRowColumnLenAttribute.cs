using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace MToExcel.Attributes
{
    /// <summary>
    /// 行列高度设置Attribute，
    /// 因为Attribute无法用在对象上，且放在类上的灵活性比较差（指定行列宽度起来太麻烦）
    /// 行Leng指定所有行的长度，Col对每一列的宽度
    /// 
    /// 
    /// 最终修正：这个Attribute设置全局的行列长度
    /// </summary>
    [AttributeUsage(AttributeTargets.Class,Inherited =false,AllowMultiple =false)]
    public class StaticRowColumnLenAttribute:Attribute
    {
        public StaticRowColumnLenAttribute(double rowHeight, double colLength)
        {
            //全局行高
            RowHeight = rowHeight;
            //全局列宽
            ColLength = colLength;
        }

        public double RowHeight{get;set;} = -1.00;

        public double ColLength{get;set;} = -1.00;


        
    }



}