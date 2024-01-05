using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace MToExcel.Attributes
{
    [AttributeUsage(AttributeTargets.Property,AllowMultiple =false,Inherited =true)]
    public class FontSets:Attribute
    {
        /// <summary>
        /// 字体名称
        /// </summary>
        /// <value></value>
        public string Name{get;set;} = "";

        /// <summary>
        /// 字体大小
        /// </summary>
        /// <value></value>
        public double Size{get;set;} = -1.0d;

        /// <summary>
        /// 是否加粗
        /// </summary>
        /// <value></value>
        public bool IsBold{get;set;} = false;

        /// <summary>
        /// 是否斜体
        /// </summary>
        /// <value></value>
        public bool IsItalic{get;set;} = false;

        /// <summary>
        /// 是否带下划线
        /// </summary>
        /// <value></value>
        public bool IsUnderline{get;set;} = false;


        /// <summary>
        /// 是否带删除线
        /// </summary>
        /// <value></value>
        public bool IsStrikeout{get;set;} = false;

        
        /// <summary>
        /// 字体颜色
        /// </summary>
        /// <value></value>
        public byte[] FontColor{get;set;} = null;

        public string Dataformat{get;set;} = null;

        #region 这里面是各种重载的构造方法

       
        public FontSets(string name, double size, bool isBold, bool isItalic, bool isUnderline, bool isStrikeout, byte[] fontColor)
        {
            Name = name;
            Size = size;
            IsBold = isBold;
            IsItalic = isItalic;
            IsUnderline = isUnderline;
            IsStrikeout = isStrikeout;
            FontColor = fontColor;
        }

        public FontSets(string name, double size, bool isBold, bool isItalic, bool isUnderline, bool isStrikeout, byte[] fontColor,string dataformat)
        {
            Name = name;
            Size = size;
            IsBold = isBold;
            IsItalic = isItalic;
            IsUnderline = isUnderline;
            IsStrikeout = isStrikeout;
            FontColor = fontColor;
            this.Dataformat = dataformat;
        }


        public FontSets(string name, double size, bool isBold, bool isItalic, bool isUnderline, bool isStrikeout)
        {
            Name = name;
            Size = size;
            IsBold = isBold;
            IsItalic = isItalic;
            IsUnderline = isUnderline;
            IsStrikeout = isStrikeout;
            
        }

        public FontSets(string name, double size, bool isBold, bool isItalic, bool isUnderline)
        {
            Name = name;
            Size = size;
            IsBold = isBold;
            IsItalic = isItalic;
            IsUnderline = isUnderline;
        }

        public FontSets(string name, double size, bool isBold, bool isItalic)
        {
            Name = name;
            Size = size;
            IsBold = isBold;
            IsItalic = isItalic;
        }

        public FontSets(string name, double size, bool isBold)
        {
            Name = name;
            Size = size;
            IsBold = isBold;
        }

        public FontSets(string name, double size)
        {
            Name = name;
            Size = size;
        }

        public FontSets(string name)
        {
            Name = name;
        }

        public FontSets(double size, bool isBold)
        {
            this.Size = size;
            this.IsBold = isBold;
        }

        public FontSets(double size, bool isBold,string dataformat)
        {
             this.Size = size;
             this.IsBold = isBold;
             this.Dataformat = dataformat;
        }

        public FontSets(string name, double size, bool isBold,string dataformat)
        {
            this.Name = name;
            this.Size = size;
            this.IsBold = isBold;
            this.Dataformat = dataformat;
        }

        #endregion






    }
}