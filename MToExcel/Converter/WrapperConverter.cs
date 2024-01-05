using MToExcel.Attributes;
using MToExcel.Models.Enums;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.XSSF.Model;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

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
        public static Dictionary<GuidTypePair,HeaderName> CustomNamePool = new Dictionary<GuidTypePair,HeaderName>();

        public static Dictionary<Type,CellStyleAttribute> CellStylePool = new Dictionary<Type, CellStyleAttribute>();

        //样式对象池
        public static Dictionary<CellStyleAttribute,ICellStyle> stylePool = new Dictionary<CellStyleAttribute,ICellStyle>();
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
                    CustomNamePool.Add(new GuidTypePair(pro.PropertyType), header);
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

            // if (WrapperConverter.CellStylePool.Contains(
            //         new KeyValuePair<Type, CellStyleAttribute>(info.PropertyType,
            //         (CellStyleAttribute)info.GetCustomAttribute(typeof(CellStyleAttribute)) == null ? new CellStyleAttribute() : (CellStyleAttribute)info.GetCustomAttribute(typeof(CellStyleAttribute))
            // )))

            var used_cellstyle = value_cell.Row.Sheet.Workbook.CreateCellStyle();

            var xssfstyle =  (XSSFCellStyle)used_cellstyle;

            var dataformat = new XSSFDataFormat(new StylesTable());

            var d = (XSSFDataFormat)value_cell.Row.Sheet.Workbook.CreateDataFormat();
            HSSFDataFormat.GetBuiltinFormat("###00.0");  

            

            
            #region  Attribute无法传递对象作为参数,故将其拆分并注释掉这一段

            // if(info.GetCustomAttribute(typeof(CellStyleAttribute))!=null)
            // {
            //     CellStyleAttribute cellStyleAttribute = (CellStyleAttribute)info.GetCustomAttribute(typeof(CellStyleAttribute));

            //     if(stylePool.ContainsKey(cellStyleAttribute))
            //     {
            //         var ExistStyle = stylePool[cellStyleAttribute];

            //         value_cell.CellStyle = ExistStyle;

            //         return true;

            //     }
            //     else
            //     {
            //         ICellStyle style = value_cell.Sheet.Workbook.CreateCellStyle();

            //         //水平对齐
            //         if(cellStyleAttribute.horizon==Horizon.Left)
            //         {
            //             style.Alignment = HorizontalAlignment.Left;
            //         }
            //         else if(cellStyleAttribute.horizon==Horizon.Center)
            //         {
            //             style.Alignment = HorizontalAlignment.Center;
            //         }
            //         else
            //         {
            //             style.Alignment = HorizontalAlignment.Right;
            //         }
                    
            //         //垂直对齐
            //         if(cellStyleAttribute.verticalHorizon==VerticalHorizon.Up)
            //         {
            //             style.VerticalAlignment = VerticalAlignment.Top;
            //         }
            //         else if(cellStyleAttribute.verticalHorizon==VerticalHorizon.Mid)
            //         {
            //             style.VerticalAlignment = VerticalAlignment.Center;
            //         }
            //         else
            //         {
            //             style.VerticalAlignment = VerticalAlignment.Bottom;
            //         }


            //         //设置字体
            //         if(cellStyleAttribute.charSet!=null)
            //         {
            //             var charset = cellStyleAttribute.charSet;

            //             var font = value_cell.Sheet.Workbook.CreateFont();
                        
            //             font.IsBold = charset.IsBold;

            //             font.IsItalic = charset.IsItalic;

            //             font.IsStrikeout = charset.IsStrikeout;

            //             font.Underline = charset.IsUnderline? FontUnderlineType.Single:FontUnderlineType.None;

            //             if(charset.Size!=null){ font.FontHeightInPoints = charset.Size.GetValueOrDefault();}

            //             if(charset.Name!=null){ font.FontName = charset.Name;}

            //             if(charset.FontColor!=null) {font.Color = charset.FontColor.GetValueOrDefault();}

            //             XSSFColor color = new XSSFColor(new byte[]{127,54,87});
                        
            //             font.Color = color.Index;

            //             style.SetFont(font);
            //         }

            //         if(cellStyleAttribute.borderSide!=null)
            //         {
            //             if(cellStyleAttribute.borderSide.Sides!=null&&cellStyleAttribute.borderSide.Sides.Count!=0)
            //         {
            //             foreach(var item in cellStyleAttribute.borderSide.Sides)
            //             {
            //                 if(item.Direct==BorderDirect.Upper)
            //                 {
                                
            //                     if(item.Wid==BorderWid.ThinBorder){ style.BorderTop = BorderStyle.Thin; }
            //                     else if(item.Wid==BorderWid.NoneBorder){ style.BorderTop = BorderStyle.None; }
            //                     else if(item.Wid==BorderWid.MiddBorder){ style.BorderTop = BorderStyle.Medium; }
            //                     else if(item.Wid==BorderWid.ThickBorder){ style.BorderTop = BorderStyle.Thick; }
                                
            //                 }
            //                 else if(item.Direct==BorderDirect.Left)
            //                 {
            //                     if(item.Wid==BorderWid.ThinBorder){ style.BorderLeft = BorderStyle.Thin; }
            //                     else if(item.Wid==BorderWid.NoneBorder){ style.BorderLeft = BorderStyle.None; }
            //                     else if(item.Wid==BorderWid.MiddBorder){ style.BorderLeft = BorderStyle.Medium; }
            //                     else if(item.Wid==BorderWid.ThickBorder){ style.BorderLeft = BorderStyle.Thick; }
            //                 }
            //                 else if(item.Direct==BorderDirect.Right)
            //                 {
            //                     if(item.Wid==BorderWid.ThinBorder){ style.BorderRight = BorderStyle.Thin; }
            //                     else if(item.Wid==BorderWid.NoneBorder){ style.BorderRight = BorderStyle.None; }
            //                     else if(item.Wid==BorderWid.MiddBorder){ style.BorderRight = BorderStyle.Medium; }
            //                     else if(item.Wid==BorderWid.ThickBorder){ style.BorderRight = BorderStyle.Thick; }
            //                 }
            //                 else if(item.Direct==BorderDirect.Bottom)
            //                 {
            //                     if(item.Wid==BorderWid.ThinBorder){ style.BorderBottom = BorderStyle.Thin; }
            //                     else if(item.Wid==BorderWid.NoneBorder){ style.BorderBottom = BorderStyle.None; }
            //                     else if(item.Wid==BorderWid.MiddBorder){ style.BorderBottom = BorderStyle.Medium; }
            //                     else if(item.Wid==BorderWid.ThickBorder){ style.BorderBottom = BorderStyle.Thick; }
            //                 }
            //                 else if(item.Direct==BorderDirect.Diagonal_slash)
            //                 {
            //                     style.BorderDiagonal = BorderDiagonal.Forward;
            //                     if(item.Wid==BorderWid.ThinBorder){ style.BorderDiagonalLineStyle = BorderStyle.Thin; }
            //                     else if(item.Wid==BorderWid.NoneBorder){ style.BorderDiagonalLineStyle = BorderStyle.None; }
            //                     else if(item.Wid==BorderWid.MiddBorder){ style.BorderDiagonalLineStyle = BorderStyle.Medium; }
            //                     else if(item.Wid==BorderWid.ThickBorder){ style.BorderDiagonalLineStyle = BorderStyle.Thick; }
                                
            //                 }
            //                 else if(item.Direct==BorderDirect.Diagonal_back_slash)
            //                 {
            //                     style.BorderDiagonal = BorderDiagonal.Backward;
            //                     if(item.Wid==BorderWid.ThinBorder){ style.BorderDiagonalLineStyle = BorderStyle.Thin; }
            //                     else if(item.Wid==BorderWid.NoneBorder){ style.BorderDiagonalLineStyle = BorderStyle.None; }
            //                     else if(item.Wid==BorderWid.MiddBorder){ style.BorderDiagonalLineStyle = BorderStyle.Medium; }
            //                     else if(item.Wid==BorderWid.ThickBorder){ style.BorderDiagonalLineStyle = BorderStyle.Thick; }
            //                 }
            //             }
            //         }
                    

            //         }

            //         value_cell.CellStyle = style;
            //         return true;
            //     }

                

            // }

            #endregion

            //字体设置标签
            if(info.GetCustomAttribute(typeof(FontSets))!=null)
            {
                FontSets item =  (FontSets)info.GetCustomAttribute(typeof(FontSets));

                var font =  value_cell.Row.Sheet.Workbook.CreateFont();

                if(item.IsBold==false)
                {
                    font.IsBold = false;
                }
                else
                {
                    font.IsBold = true;
                }

                if(item.IsItalic==false)
                {
                    font.IsItalic = false;
                }
                else
                {
                    font.IsItalic = true;
                }

                if(item.IsStrikeout==false)
                {
                    font.IsStrikeout = false;
                }
                else
                {
                    font.IsStrikeout = true;
                }

                if(item.IsUnderline==false)
                {
                    font.Underline = FontUnderlineType.Single;
                }
                else
                {
                    font.Underline = FontUnderlineType.None;
                }

                if(item.Size!=-1.0d)
                {
                    font.FontHeightInPoints = item.Size;
                }

                if(item.FontColor!=null)
                {
                    ((XSSFFont)font).SetColor(new XSSFColor(item.FontColor));
                }

                if(!string.IsNullOrEmpty(item.Name))
                {
                    font.FontName = item.Name;
                }

                used_cellstyle.SetFont(font);

                if(!string.IsNullOrEmpty(item.Dataformat))
                {
                    used_cellstyle.DataFormat = HSSFDataFormat.GetBuiltinFormat(item.Dataformat);
                }

            }

            //边框设置标签
            if(info.GetCustomAttribute(typeof(BorderStyleAttribute))!=null)
            {
                var bordersets =  info.GetCustomAttributes(typeof(BorderStyleAttribute)).ToList();

                //更具边的的方向distinct
                var distinct_list =  bordersets.DistinctBy(it=>((BorderStyleAttribute)it).Direct).ToList();

                distinct_list.ForEach(itt=>{

                    var bs =  (BorderStyleAttribute)itt;

                    if(bs.Direct == BorderDirect.Upper)
                    {
                        if(bs.Wid == BorderWid.NoneBorder)
                        {
                            // DoNothing 无边框什么都不做
                        }
                        else if(bs.Wid == BorderWid.ThinBorder)
                        {
                            used_cellstyle.BorderTop = BorderStyle.Thin;
                        }
                        else if(bs.Wid == BorderWid.MiddBorder)
                        {
                            used_cellstyle.BorderTop = BorderStyle.Medium;
                        }
                        else
                        {
                            used_cellstyle.BorderTop = BorderStyle.Thick;
                        }

                        if(bs.Color!=null)
                        {
                            XSSFColor color = new  XSSFColor(bs.Color);
                            ((XSSFCellStyle)used_cellstyle).SetTopBorderColor(color);
                        }
                    }
                    else if(bs.Direct == BorderDirect.Bottom)
                    {
                        if(bs.Wid == BorderWid.NoneBorder)
                        {
                            // DoNothing 无边框什么都不做
                        }
                        else if(bs.Wid == BorderWid.ThinBorder)
                        {
                            used_cellstyle.BorderBottom = BorderStyle.Thin;
                        }
                        else if(bs.Wid == BorderWid.MiddBorder)
                        {
                            used_cellstyle.BorderBottom = BorderStyle.Medium;
                        }
                        else
                        {
                            used_cellstyle.BorderBottom = BorderStyle.Thick;
                        }

                        if(bs.Color!=null)
                        {
                            XSSFColor color = new  XSSFColor(bs.Color);
                            ((XSSFCellStyle)used_cellstyle).SetBottomBorderColor(color);
                        }
                    }
                    else if(bs.Direct == BorderDirect.Right)
                    {
                        if(bs.Wid == BorderWid.NoneBorder)
                        {
                            // DoNothing 无边框什么都不做
                        }
                        else if(bs.Wid == BorderWid.ThinBorder)
                        {
                            used_cellstyle.BorderRight = BorderStyle.Thin;
                        }
                        else if(bs.Wid == BorderWid.MiddBorder)
                        {
                            used_cellstyle.BorderRight = BorderStyle.Medium;
                        }
                        else
                        {
                            used_cellstyle.BorderRight = BorderStyle.Thick;
                        }

                        if(bs.Color!=null)
                        {
                            XSSFColor color = new  XSSFColor(bs.Color);
                            ((XSSFCellStyle)used_cellstyle).SetRightBorderColor(color);
                        }
                    }
                    else if(bs.Direct == BorderDirect.Left)
                    {
                        if(bs.Wid == BorderWid.NoneBorder)
                        {
                            // DoNothing 无边框什么都不做
                        }
                        else if(bs.Wid == BorderWid.ThinBorder)
                        {
                            used_cellstyle.BorderLeft = BorderStyle.Thin;
                        }
                        else if(bs.Wid == BorderWid.MiddBorder)
                        {
                            used_cellstyle.BorderLeft = BorderStyle.Medium;
                        }
                        else
                        {
                            used_cellstyle.BorderLeft = BorderStyle.Thick;
                        }

                        if(bs.Color!=null)
                        {
                            XSSFColor color = new  XSSFColor(bs.Color);
                            ((XSSFCellStyle)used_cellstyle).SetLeftBorderColor(color);
                        }
                    }
                    else if(bs.Direct == BorderDirect.Diagonal_slash)
                    {
                        used_cellstyle.BorderDiagonal = BorderDiagonal.Forward;

                        if(bs.Wid == BorderWid.NoneBorder)
                        {
                            // DoNothing 无边框什么都不做
                        }
                        else if(bs.Wid == BorderWid.ThinBorder)
                        {
                            used_cellstyle.BorderDiagonalLineStyle = BorderStyle.Thin;
                        }
                        else if(bs.Wid == BorderWid.MiddBorder)
                        {
                            used_cellstyle.BorderDiagonalLineStyle = BorderStyle.Medium;
                        }
                        else
                        {
                            used_cellstyle.BorderDiagonalLineStyle = BorderStyle.Thick;
                        }

                        if(bs.Color!=null)
                        {
                            XSSFColor color = new  XSSFColor(bs.Color);
                            ((XSSFCellStyle)used_cellstyle).SetDiagonalBorderColor(color);
                        }
                    }
                    else
                    {
                        used_cellstyle.BorderDiagonal = BorderDiagonal.Backward;

                        if(bs.Wid == BorderWid.NoneBorder)
                        {
                            // DoNothing 无边框什么都不做
                        }
                        else if(bs.Wid == BorderWid.ThinBorder)
                        {
                            used_cellstyle.BorderDiagonalLineStyle = BorderStyle.Thin;
                        }
                        else if(bs.Wid == BorderWid.MiddBorder)
                        {
                            used_cellstyle.BorderDiagonalLineStyle = BorderStyle.Medium;
                        }
                        else
                        {
                            used_cellstyle.BorderDiagonalLineStyle = BorderStyle.Thick;
                        }

                        if(bs.Color!=null)
                        {
                            XSSFColor color = new  XSSFColor(bs.Color);
                            ((XSSFCellStyle)used_cellstyle).SetDiagonalBorderColor(color);
                        }
                    }



                    

                });

            }

            //对齐方式表
            if(info.GetCustomAttribute(typeof(HorizonAttribute))!=null)
            {
                var horizon =  (HorizonAttribute)info.GetCustomAttribute(typeof(HorizonAttribute));

                if(horizon.horizon == Horizon.Center)
                {
                    used_cellstyle.Alignment = HorizontalAlignment.Center;
                }
                else if(horizon.horizon == Horizon.Left)
                {
                    used_cellstyle.Alignment = HorizontalAlignment.Left;
                }
                else if(horizon.horizon == Horizon.Right)
                {
                    used_cellstyle.Alignment = HorizontalAlignment.Right;
                }

                if(horizon.verticalHorizon==VerticalHorizon.Up)
                {
                    used_cellstyle.VerticalAlignment = VerticalAlignment.Top;
                }

                if(horizon.verticalHorizon==VerticalHorizon.Mid)
                {
                    used_cellstyle.VerticalAlignment = VerticalAlignment.Center;
                }

                if(horizon.verticalHorizon==VerticalHorizon.Down)
                {
                    used_cellstyle.VerticalAlignment = VerticalAlignment.Bottom;
                }


            }

            
            return false;

        }
    }



}

