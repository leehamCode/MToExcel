using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Threading.Tasks;
using MToExcel.Models.Enums;
using MToExcel.Models.Param;

/// <summary>
/// 2023/12/20
/// 因为C#不支持在Attribute上通过结构体和对象传递参数,故次类需要被拆分
/// </summary>
namespace MToExcel.Attributes
{
    [AttributeUsage(AttributeTargets.Property,AllowMultiple =false,Inherited = true)]
    public class CellStyleAttribute:Attribute
    {
        /// <summary>
        /// 水平对齐样式
        /// </summary>
        /// <value></value>
        public CellStyleAttribute(Horizon horizon, VerticalHorizon verticalHorizon, short backGroundColor, bool wrapText) 
        {
            this.horizon = horizon;
            this.verticalHorizon = verticalHorizon;
            this.BackGroundColor = backGroundColor;
            this.WrapText = wrapText;
   
        }
        public Horizon horizon{get;set;}

        /// <summary>
        /// 垂直对齐样式
        /// </summary>
        /// <value></value>
        public VerticalHorizon verticalHorizon{get;set;}

        /// <summary>
        /// 边界样式设置
        /// </summary>
        /// <value></value>
        public BorderSide? borderSide {get;set;}

        /// <summary>
        /// 字体样式
        /// </summary>
        /// <value></value>
        public CharSet charSet{get;set;}

        /// <summary>
        /// 前景颜色
        /// </summary>
        /// <value></value>
        public short? ForgeColor{get;set;}

        /// <summary>
        /// 背景颜色
        /// </summary>
        /// <value></value>
        public short BackGroundColor{get;set;}

        /// <summary>
        /// 是否自动换行
        /// </summary>
        /// <value></value>
        public bool WrapText{get;set;} = false;

        public CellStyleAttribute(Horizon horizon, VerticalHorizon verticalHorizon, BorderSide? borderSide, CharSet? charSet, short? forgeColor, short backGroundColor, bool wrapText)
        {
            this.horizon = horizon;
            this.verticalHorizon = verticalHorizon;
            this.borderSide = borderSide;
            this.charSet = charSet;
            ForgeColor = forgeColor;
            BackGroundColor = backGroundColor;
            WrapText = wrapText;
        }

        public CellStyleAttribute(Horizon horizon, VerticalHorizon verticalHorizon,bool wrapText,CharSet charSet)
        {
                this.horizon = horizon;
                this.verticalHorizon = verticalHorizon;
                this.WrapText = wrapText;
                this.charSet = charSet;
        }

        public CellStyleAttribute(Horizon horizon, VerticalHorizon verticalHorizon,bool wrapText,BorderSide borderSide)
        {
                this.horizon = horizon;
                this.verticalHorizon = verticalHorizon;
                this.WrapText = wrapText;
                this.borderSide = borderSide;
        }
        
        public CellStyleAttribute(Horizon horizon, VerticalHorizon verticalHorizon,bool wrapText)
        {
            this.horizon = horizon;
            this.verticalHorizon = verticalHorizon;
            this.WrapText = wrapText;
        }

        public CellStyleAttribute()
        {
            
        }



        // override object.Equals,为了避免创建过多的CellStyle对象，在这里重写比较
        public override bool Equals(object obj)
        {
            var target =  (CellStyleAttribute)obj;

            bool varOne = target.charSet != null ? target.charSet.Equals(charSet) : true;
            bool varTwo = target.borderSide!=null? target.borderSide.Equals(borderSide):true;
            
            if( horizon == target.horizon&&
                verticalHorizon == target.verticalHorizon&&
                ForgeColor == target.ForgeColor&&
                BackGroundColor == target.BackGroundColor&&
                WrapText == target.WrapText&&
                varOne&&
                varTwo
                )
            {
                return true;        
            }
            else{
                return false;
            }
        }
        
        // override object.GetHashCode
        public override int GetHashCode()
        {
            if(borderSide==null&&charSet==null)
            {
                return HashCode.Combine(horizon,verticalHorizon,ForgeColor,BackGroundColor,WrapText);
            }
            else if(borderSide==null&&charSet!=null)
            {
                return HashCode.Combine(horizon,verticalHorizon,ForgeColor,BackGroundColor,WrapText,charSet.GetHashCode());
            }
            else if(borderSide!=null&&charSet==null)
            {
                return HashCode.Combine(horizon,verticalHorizon,ForgeColor,BackGroundColor,WrapText,borderSide.GetHashCode());
            }
            else{
                return HashCode.Combine(horizon,verticalHorizon,ForgeColor,BackGroundColor,WrapText,borderSide.GetHashCode(),charSet.GetHashCode());
            }

        }
    }
}