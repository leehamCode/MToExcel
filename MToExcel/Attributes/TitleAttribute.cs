using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace MToExcel.Attributes
{
    [AttributeUsage(AttributeTargets.Class,Inherited = false,AllowMultiple =false)]
    public class TitleAttribute:Attribute
    {
        /// <summary>
        /// 标题内容，后面可能需要支持表达式 
        /// </summary>
        /// <value></value>
        public string Context{get;set;}

        /// <summary>
        /// 需要合并多少个列
        /// </summary>
        /// <value></value>
        public int Col_Merge_num {get;set;} = 1;

        /// <summary>
        /// 需要合并多少行
        /// </summary>
        /// <value></value>
        public int Row_Merge_num {get;set;} = 1;

        /// <summary>
        /// 单行高度
        /// </summary>
        /// <value></value>
        public double  Single_Height{get;set;}

        
        public string Font_Name{get;set;}

        public int Font_Size{get;set;}

        public byte[] Font_color{get;set;}

        public bool IsBold{get;set;} = false;

        public bool IsItalic{get;set;} = false;

        public byte[] Back_color{get;set;}

        public byte[] Fore_color{get;set;}

    }
}