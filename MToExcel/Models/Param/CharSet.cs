using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace MToExcel.Models.Param
{
    /// <summary>
    /// 字体样式类
    /// </summary>
    public class CharSet
    {
        /// <summary>
        /// 字体名称
        /// </summary>
        /// <value></value>
        public string? Name{get;set;}

        /// <summary>
        /// 字体大小
        /// </summary>
        /// <value></value>
        public int Size{get;set;}

        /// <summary>
        /// 是否加粗
        /// </summary>
        /// <value></value>
        public bool IsBold{get;set;}

        /// <summary>
        /// 是否斜体
        /// </summary>
        /// <value></value>
        public bool IsItalic{get;set;}

        /// <summary>
        /// 是否带下划线
        /// </summary>
        /// <value></value>
        public bool IsUnderline{get;set;}

        /// <summary>
        /// 字体颜色
        /// </summary>
        /// <value></value>
        public short FontColor{get;set;}

        /// <summary>
        /// 是否带删除线
        /// </summary>
        /// <value></value>
        public bool IsStrikeout{get;set;}

        // override object.Equals
        public override bool Equals(object obj)
        {
            var target = (CharSet)obj;
            if(target.Name==Name&&
               target.Size==Size&&
               target.IsBold==IsBold&&
               target.IsItalic == IsItalic&&
               target.IsUnderline == IsUnderline&&
               target.FontColor == FontColor&&
               target.IsStrikeout == IsStrikeout){
                return true;
            }
            else{
                return false;
            }

        }
        
        // override object.GetHashCode
        public override int GetHashCode()
        {
            return HashCode.Combine(Name,Size,IsBold,IsItalic,IsUnderline,FontColor,IsStrikeout);
        }

    }
}