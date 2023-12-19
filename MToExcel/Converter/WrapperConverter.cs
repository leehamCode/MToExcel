using MToExcel.Attributes;
using MToExcel.Models.Enums;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
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
        public static Dictionary<Type,IgnoreType> IgnoreTypePool = new Dictionary<Type, IgnoreType>();
        public static Dictionary<Type,HeaderName> CustomNamePool = new Dictionary<Type,HeaderName>();

        public static Dictionary<Type,CellStyleAttribute> CellStylePool = new Dictionary<Type, CellStyleAttribute>();
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
                IgnoreType ignore = (IgnoreType)pro.GetCustomAttribute(typeof(IgnoreType));
                HeaderName header = (HeaderName)pro.GetCustomAttribute(typeof(HeaderName));
                CellStyleAttribute style = (CellStyleAttribute)pro.GetCustomAttribute(typeof(CellStyleAttribute));
                
                //检查属性是否存在标签，存在就保存到类型池中
                if (refer != null)
                {
                    //将打了标记的类型和标记本身放到类型池中
                    TypePool.Add(pro.PropertyType, refer);
                }
                if(ignore != null)
                {
                    IgnoreTypePool.Add(pro.PropertyType, ignore);
                }
                if(header != null)
                {
                    CustomNamePool.Add(pro.PropertyType, header);
                }
                if(style != null)
                {
                    CellStylePool.Add(pro.PropertyType,style);
                }


            }
        }

        /// <summary>
        /// 检查是否需要，并添加单元格样式
        /// </summary>
        /// <param name="info"></param>
        /// <param name="value_cell"></param>
        /// <returns></returns>
        public static bool PutOnCellStyle(PropertyInfo info,ICell value_cell)
        {
            if(info==null||value_cell==null)
            {
                return false;
            }

            //Type type = info.PropertyType;

            if (WrapperConverter.CellStylePool.Contains(
                    new KeyValuePair<Type, CellStyleAttribute>(info.PropertyType,
                    (CellStyleAttribute)info.GetCustomAttribute(typeof(CellStyleAttribute)) == null ? new CellStyleAttribute() : (CellStyleAttribute)info.GetCustomAttribute(typeof(CellStyleAttribute))
            )))
            {
                CellStyleAttribute cellStyleAttribute = (CellStyleAttribute)info.GetCustomAttribute(typeof(CellStyleAttribute));

                ICellStyle style = value_cell.Sheet.Workbook.CreateCellStyle();

                //水平对齐
                if(cellStyleAttribute.horizon==Horizon.Left)
                {
                    style.Alignment = HorizontalAlignment.Left;
                }
                else if(cellStyleAttribute.horizon==Horizon.Center)
                {
                    style.Alignment = HorizontalAlignment.Center;
                }
                else
                {
                    style.Alignment = HorizontalAlignment.Right;
                }
                
                //垂直对齐
                if(cellStyleAttribute.verticalHorizon==VerticalHorizon.Up)
                {
                    style.VerticalAlignment = VerticalAlignment.Top;
                }
                else if(cellStyleAttribute.verticalHorizon==VerticalHorizon.Mid)
                {
                    style.VerticalAlignment = VerticalAlignment.Center;
                }
                else
                {
                    style.VerticalAlignment = VerticalAlignment.Bottom;
                }


                //设置字体
                if(cellStyleAttribute.charSet!=null)
                {
                    var charset = cellStyleAttribute.charSet;

                    var font = value_cell.Sheet.Workbook.CreateFont();
                    
                    font.IsBold = charset.IsBold;

                    font.IsItalic = charset.IsItalic;

                    font.IsStrikeout = charset.IsStrikeout;

                    font.Underline = charset.IsUnderline? FontUnderlineType.Single:FontUnderlineType.None;

                    if(charset.Size!=null){ font.FontHeightInPoints = charset.Size.GetValueOrDefault();}

                    if(charset.Name!=null){ font.FontName = charset.Name;}

                    if(charset.FontColor!=null) {font.Color = charset.FontColor.GetValueOrDefault();}

                    style.SetFont(font);
                }


                if(cellStyleAttribute.borderSide.Sides!=null&&cellStyleAttribute.borderSide.Sides.Count!=0)
                {
                    foreach(var item in cellStyleAttribute.borderSide.Sides)
                    {
                        if(item.Direct==BorderDirect.Upper)
                        {
                            
                            if(item.Wid==BorderWid.ThinBorder){ style.BorderTop = BorderStyle.Thin; }
                            else if(item.Wid==BorderWid.NoneBorder){ style.BorderTop = BorderStyle.None; }
                            else if(item.Wid==BorderWid.MiddBorder){ style.BorderTop = BorderStyle.Medium; }
                            else if(item.Wid==BorderWid.ThickBorder){ style.BorderTop = BorderStyle.Thick; }
                            
                        }
                        else if(item.Direct==BorderDirect.Left)
                        {
                            if(item.Wid==BorderWid.ThinBorder){ style.BorderLeft = BorderStyle.Thin; }
                            else if(item.Wid==BorderWid.NoneBorder){ style.BorderLeft = BorderStyle.None; }
                            else if(item.Wid==BorderWid.MiddBorder){ style.BorderLeft = BorderStyle.Medium; }
                            else if(item.Wid==BorderWid.ThickBorder){ style.BorderLeft = BorderStyle.Thick; }
                        }
                        else if(item.Direct==BorderDirect.Right)
                        {
                            if(item.Wid==BorderWid.ThinBorder){ style.BorderRight = BorderStyle.Thin; }
                            else if(item.Wid==BorderWid.NoneBorder){ style.BorderRight = BorderStyle.None; }
                            else if(item.Wid==BorderWid.MiddBorder){ style.BorderRight = BorderStyle.Medium; }
                            else if(item.Wid==BorderWid.ThickBorder){ style.BorderRight = BorderStyle.Thick; }
                        }
                        else if(item.Direct==BorderDirect.Bottom)
                        {
                            if(item.Wid==BorderWid.ThinBorder){ style.BorderBottom = BorderStyle.Thin; }
                            else if(item.Wid==BorderWid.NoneBorder){ style.BorderBottom = BorderStyle.None; }
                            else if(item.Wid==BorderWid.MiddBorder){ style.BorderBottom = BorderStyle.Medium; }
                            else if(item.Wid==BorderWid.ThickBorder){ style.BorderBottom = BorderStyle.Thick; }
                        }
                    }
                    
                }
            }

            return true;
        }
    }
}

