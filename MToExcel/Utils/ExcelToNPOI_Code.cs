using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;

namespace MToExcel.Utils
{
    /// <summary>
    /// 这个雷可以将一个Excel的固定区域反向生成NPOI样式代码
    /// </summary>
    public class ExcelToNPOI_Code
    {
        public static void InputStaticExcel_Head(string filepath,int StopRow,int StopCol)
        {
            int Height = StopRow;

            int Len = StopCol;
            StringBuilder builder = new StringBuilder($@"
                XSSFWorkbook workbook = new XSSFWorkbook();
                ISheet sheet = workbook.CreateSheet();
            ");

            Console.WriteLine("进入NPOI表头样式代码生成器(静态的)");

            FileStream fs = new FileStream(filepath,FileMode.OpenOrCreate,FileAccess.ReadWrite);

            XSSFWorkbook workbook = new XSSFWorkbook(fs);

            Stack<(XSSFCellStyle, int, int)> preStacks = new Stack<(XSSFCellStyle, int, int)>();

            Stack<(int, int)> Ijlist = new Stack<(int, int)>();                       //保存上一个样式相同的style的i,j坐标

            ISheet sheet = workbook.GetSheetAt(0);

            #region 先设置合并单元格，再设置具体单元格样式

            int RegionNum = sheet.NumMergedRegions;

            for(int k = 0; k < RegionNum; k++)
            {
                CellRangeAddress region = sheet.GetMergedRegion(k);
                if((region.FirstRow <= region.LastRow && region.LastRow <= Height) && (region.FirstColumn<=region.LastColumn&&region.LastColumn<=Len))
                {
                    builder.AppendLine($@"CellRangeAddress region{k} = new CellRangeAddress({region.FirstRow},{region.LastRow},{region.FirstColumn},{region.LastColumn});sheet.AddMergedRegion(region{k}); ");

                    (bool, List<string>) res = IsRegionHasBorder(sheet, region);
                    if (res.Item1)
                    {
                        res.Item2.ForEach(item => {

                            if (item.Equals("left"))
                                builder.Append($@" RegionUtil.SetBorderLeft(1,region{k},sheet); ");
                            else if (item.Equals("right"))
                                builder.Append($@" RegionUtil.SetBorderRight(1,region{k},sheet); ");
                            else if (item.Equals("top"))
                                builder.Append($@" RegionUtil.SetBorderTop(1,region{k},sheet); ");
                            else if (item.Equals("bottom"))
                                builder.Append($@" RegionUtil.SetBorderBottom(1,region{k},sheet); ");

                        
                        });
                        
                    }
                }
            }

            #endregion

            for (int i = 0; i < Height; i++)
            {
                IRow row = sheet.GetRow(i)?? null;

                if (row != null)
                {
                    builder.AppendLine($@" IRow row{i} =  sheet.CreateRow({i}); row{i}.HeightInPoints = {row.HeightInPoints}f; ");

                    
                                                            
                    for (int j = 0; j < Len; j++)
                    {
                        ICell cell = row.GetCell(j)==null?  null: row.GetCell(j);
                        if (i == 0)
                        {
                            builder.AppendLine($@"sheet.SetColumnWidth({j},{sheet.GetColumnWidth(j)});");
                        }
                            
                        if (cell != null)
                        {
                            
                            string cell_str = $@"  ICell row{i}cell{j} = row{i}.CreateCell({j});   ";
                            

                            switch (cell.CellType)
                            {
                                case CellType.Unknown:
                                    cell_str += $@" row{i}cell{j}.SetCellValue($@""{cell.StringCellValue}""); ";
                                    break;
                                case CellType.Numeric:
                                    cell_str += $@" row{i}cell{j}.SetCellValue({cell.NumericCellValue}); ";
                                    break;
                                case CellType.String:
                                    cell_str += $@" row{i}cell{j}.SetCellValue($@""{cell.StringCellValue}""); ";
                                    break;
                                case CellType.Formula:
                                    cell_str += $@" row{i}cell{j}.SetCellValue({cell.CellFormula}); ";
                                    break;
                                case CellType.Blank:
                                    cell_str += $@" row{i}cell{j}.SetCellValue($@""""); ";
                                    break;
                                case CellType.Boolean:
                                    cell_str += $@" row{i}cell{j}.SetCellValue({cell.BooleanCellValue}); ";
                                    break;
                                case CellType.Error:
                                    cell_str += $@" row{i}cell{j}.SetCellValue({cell.ErrorCellValue}); ";
                                    break;
                                default:
                                    cell_str += $@" row{i}cell{j}.SetCellValue($@""{cell.StringCellValue}""); ";
                                    break;
                            }       //生成设置值的代码
                            builder.AppendLine(cell_str);
                            
                            builder.AppendLine("//------------------------------------------分割线\n");

                           

                            XSSFCellStyle cellStyle = (XSSFCellStyle)cell.CellStyle;               //尝试获取样式代码

                           

                            string style_str = $@" XSSFCellStyle row{i}cell{j}Style = (XSSFCellStyle)workbook.CreateCellStyle(); ";


                            //先同一对齐方向
                            switch (cellStyle.Alignment)
                            {
                                case HorizontalAlignment.General:
                                    style_str += $@" row{i}cell{j}Style.Alignment =  HorizontalAlignment.General;";
                                    break;
                                case HorizontalAlignment.Left:
                                    style_str += $@" row{i}cell{j}Style.Alignment =  HorizontalAlignment.Left;";
                                    break;
                                case HorizontalAlignment.Center:
                                    style_str += $@" row{i}cell{j}Style.Alignment =  HorizontalAlignment.Center;";
                                    break;
                                case HorizontalAlignment.Right:
                                    style_str += $@" row{i}cell{j}Style.Alignment =  HorizontalAlignment.Center;";
                                    break;
                                case HorizontalAlignment.Justify:
                                    style_str += $@" row{i}cell{j}Style.Alignment =  HorizontalAlignment.Justify;";
                                    break;
                                case HorizontalAlignment.Fill:
                                    style_str += $@" row{i}cell{j}Style.Alignment =  HorizontalAlignment.Fill;";
                                    break;
                                case HorizontalAlignment.CenterSelection:
                                    style_str += $@" row{i}cell{j}Style.Alignment =  HorizontalAlignment.CenterSelection;";
                                    break;
                                case HorizontalAlignment.Distributed:
                                    style_str += $@" row{i}cell{j}Style.Alignment =  HorizontalAlignment.Distributed;";
                                    break;
                                default:
                                    style_str += $@" cellStyle.Alignment =  null;";
                                    break;
                            }

                            switch (cellStyle.VerticalAlignment)
                            {
                                case VerticalAlignment.None:
                                    style_str += $@" row{i}cell{j}Style.VerticalAlignment = VerticalAlignment.None; ";
                                    break;
                                case VerticalAlignment.Top:
                                    style_str += $@" row{i}cell{j}Style.VerticalAlignment = VerticalAlignment.Top; ";
                                    break;
                                case VerticalAlignment.Center:
                                    style_str += $@" row{i}cell{j}Style.VerticalAlignment = VerticalAlignment.Center; ";
                                    break;
                                case VerticalAlignment.Bottom:
                                    style_str += $@" row{i}cell{j}Style.VerticalAlignment = VerticalAlignment.Bottom; ";
                                    break;
                                case VerticalAlignment.Justify:
                                    style_str += $@" row{i}cell{j}Style.VerticalAlignment = VerticalAlignment.Justify; ";
                                    break;
                                case VerticalAlignment.Distributed:
                                    style_str += $@" row{i}cell{j}Style.VerticalAlignment = VerticalAlignment.Distributed; ";
                                    break;
                                default:
                                    style_str += $@" row{i}cell{j}Style.VerticalAlignment = VerticalAlignment.None; ";
                                    break;
                            }

                            switch (cellStyle.BorderBottom)
                            {
                                case BorderStyle.None:
                                    style_str += $@" row{i}cell{j}Style.BorderBottom = BorderStyle.None;";
                                    break;
                                case BorderStyle.Thin:
                                    style_str += $@" row{i}cell{j}Style.BorderBottom = BorderStyle.Thin;";
                                    break;
                                case BorderStyle.Medium:
                                    style_str += $@" row{i}cell{j}Style.BorderBottom = BorderStyle.Medium;";
                                    break;
                                case BorderStyle.Dashed:
                                    style_str += $@" row{i}cell{j}Style.BorderBottom = BorderStyle.Dashed;";
                                    break;
                                case BorderStyle.Dotted:
                                    style_str += $@" row{i}cell{j}Style.BorderBottom = BorderStyle.Dotted;";
                                    break;
                                case BorderStyle.Thick:
                                    style_str += $@" row{i}cell{j}Style.BorderBottom = BorderStyle.Thick;";
                                    break;
                                case BorderStyle.Double:
                                    style_str += $@" row{i}cell{j}Style.BorderBottom = BorderStyle.Double;";
                                    break;
                                case BorderStyle.Hair:
                                    style_str += $@" row{i}cell{j}Style.BorderBottom = BorderStyle.Hair;";
                                    break;
                                case BorderStyle.MediumDashed:
                                    style_str += $@" row{i}cell{j}Style.BorderBottom = BorderStyle.MediumDashed;";
                                    break;
                                case BorderStyle.DashDot:
                                    style_str += $@" row{i}cell{j}Style.BorderBottom = BorderStyle.DashDot;";
                                    break;
                                case BorderStyle.MediumDashDot:
                                    style_str += $@" row{i}cell{j}Style.BorderBottom = BorderStyle.MediumDashDot;";
                                    break;
                                case BorderStyle.DashDotDot:
                                    style_str += $@" row{i}cell{j}Style.BorderBottom = BorderStyle.DashDotDot;";
                                    break;
                                case BorderStyle.MediumDashDotDot:
                                    style_str += $@" row{i}cell{j}Style.BorderBottom = BorderStyle.MediumDashDotDot;";
                                    break;
                                case BorderStyle.SlantedDashDot:
                                    style_str += $@" row{i}cell{j}Style.BorderBottom = BorderStyle.SlantedDashDot;";
                                    break;
                                default:
                                    style_str += $@" row{i}cell{j}Style.BorderBottom = BorderStyle.None;";
                                    break;
                            }

                            switch (cellStyle.BorderLeft)
                            {
                                case BorderStyle.None:
                                    style_str += $@" row{i}cell{j}Style.BorderLeft =  BorderStyle.None;";
                                    break;
                                case BorderStyle.Thin:
                                    style_str += $@" row{i}cell{j}Style.BorderLeft =  BorderStyle.Thin;";
                                    break;
                                case BorderStyle.Medium:
                                    style_str += $@" row{i}cell{j}Style.BorderLeft =  BorderStyle.Medium;";
                                    break;
                                case BorderStyle.Dashed:
                                    style_str += $@" row{i}cell{j}Style.BorderLeft =  BorderStyle.Dashed;";
                                    break;
                                case BorderStyle.Dotted:
                                    style_str += $@" row{i}cell{j}Style.BorderLeft =  BorderStyle.Dotted;";
                                    break;
                                case BorderStyle.Thick:
                                    style_str += $@" row{i}cell{j}Style.BorderLeft =  BorderStyle.Thick;";
                                    break;
                                case BorderStyle.Double:
                                    style_str += $@" row{i}cell{j}Style.BorderLeft =  BorderStyle.Double;";
                                    break;
                                case BorderStyle.Hair:
                                    style_str += $@" row{i}cell{j}Style.BorderLeft =  BorderStyle.Hair;";
                                    break;
                                case BorderStyle.MediumDashed:
                                    style_str += $@" row{i}cell{j}Style.BorderLeft =  BorderStyle.MediumDashed;";
                                    break;
                                case BorderStyle.DashDot:
                                    style_str += $@" row{i}cell{j}Style.BorderLeft =  BorderStyle.DashDot;";
                                    break;
                                case BorderStyle.MediumDashDot:
                                    style_str += $@" row{i}cell{j}Style.BorderLeft =  BorderStyle.MediumDashDot;";
                                    break;
                                case BorderStyle.DashDotDot:
                                    style_str += $@" row{i}cell{j}Style.BorderLeft =  BorderStyle.DashDotDot;";
                                    break;
                                case BorderStyle.MediumDashDotDot:
                                    style_str += $@" row{i}cell{j}Style.BorderLeft =  BorderStyle.MediumDashDotDot;";
                                    break;
                                case BorderStyle.SlantedDashDot:
                                    style_str += $@" row{i}cell{j}Style.BorderLeft =  BorderStyle.SlantedDashDot;";
                                    break;
                                default:
                                    style_str += $@" row{i}cell{j}Style.BorderLeft =  BorderStyle.None;";
                                    break;
                            }

                            switch (cellStyle.BorderRight)
                            {
                                case BorderStyle.None:
                                    style_str += $@" row{i}cell{j}Style.BorderRight =  BorderStyle.None; ";
                                    break;
                                case BorderStyle.Thin:
                                    style_str += $@" row{i}cell{j}Style.BorderRight =  BorderStyle.Thin; ";
                                    break;
                                case BorderStyle.Medium:
                                    style_str += $@" row{i}cell{j}Style.BorderRight =  BorderStyle.Medium; ";
                                    break;
                                case BorderStyle.Dashed:
                                    style_str += $@" row{i}cell{j}Style.BorderRight =  BorderStyle.Dashed; ";
                                    break;
                                case BorderStyle.Dotted:
                                    style_str += $@" row{i}cell{j}Style.BorderRight =  BorderStyle.Dotted; ";
                                    break;
                                case BorderStyle.Thick:
                                    style_str += $@" row{i}cell{j}Style.BorderRight =  BorderStyle.Thick; ";
                                    break;
                                case BorderStyle.Double:
                                    style_str += $@" row{i}cell{j}Style.BorderRight =  BorderStyle.Double; ";
                                    break;
                                case BorderStyle.Hair:
                                    style_str += $@" row{i}cell{j}Style.BorderRight =  BorderStyle.Hair; ";
                                    break;
                                case BorderStyle.MediumDashed:
                                    style_str += $@" row{i}cell{j}Style.BorderRight =  BorderStyle.MediumDashed; ";
                                    break;
                                case BorderStyle.DashDot:
                                    style_str += $@" row{i}cell{j}Style.BorderRight =  BorderStyle.DashDot; ";
                                    break;
                                case BorderStyle.MediumDashDot:
                                    style_str += $@" row{i}cell{j}Style.BorderRight =  BorderStyle.MediumDashDot; ";
                                    break;
                                case BorderStyle.DashDotDot:
                                    style_str += $@" row{i}cell{j}Style.BorderRight =  BorderStyle.DashDotDot; ";
                                    break;
                                case BorderStyle.MediumDashDotDot:
                                    style_str += $@" row{i}cell{j}Style.BorderRight =  BorderStyle.MediumDashDotDot; ";
                                    break;
                                case BorderStyle.SlantedDashDot:
                                    style_str += $@" row{i}cell{j}Style.BorderRight =  BorderStyle.SlantedDashDot; ";
                                    break;
                                default:
                                    style_str += $@" row{i}cell{j}Style.BorderRight =  BorderStyle.None; ";
                                    break;
                            }

                            switch (cellStyle.BorderTop)
                            {
                                case BorderStyle.None:
                                    style_str += $@" row{i}cell{j}Style.BorderTop = BorderStyle.None; ";
                                    break;
                                case BorderStyle.Thin:
                                    style_str += $@" row{i}cell{j}Style.BorderTop = BorderStyle.Thin; ";
                                    break;
                                case BorderStyle.Medium:
                                    style_str += $@" row{i}cell{j}Style.BorderTop = BorderStyle.Medium; ";
                                    break;
                                case BorderStyle.Dashed:
                                    style_str += $@" row{i}cell{j}Style.BorderTop = BorderStyle.Dashed; ";
                                    break;
                                case BorderStyle.Dotted:
                                    style_str += $@" row{i}cell{j}Style.BorderTop = BorderStyle.Dotted; ";
                                    break;
                                case BorderStyle.Thick:
                                    style_str += $@" row{i}cell{j}Style.BorderTop = BorderStyle.Thick; ";
                                    break;
                                case BorderStyle.Double:
                                    style_str += $@" row{i}cell{j}Style.BorderTop = BorderStyle.Double; ";
                                    break;
                                case BorderStyle.Hair:
                                    style_str += $@" row{i}cell{j}Style.BorderTop = BorderStyle.Hair; ";
                                    break;
                                case BorderStyle.MediumDashed:
                                    style_str += $@" row{i}cell{j}Style.BorderTop = BorderStyle.MediumDashed; ";
                                    break;
                                case BorderStyle.DashDot:
                                    style_str += $@" row{i}cell{j}Style.BorderTop = BorderStyle.DashDot; ";
                                    break;
                                case BorderStyle.MediumDashDot:
                                    style_str += $@" row{i}cell{j}Style.BorderTop = BorderStyle.MediumDashDot; ";
                                    break;
                                case BorderStyle.DashDotDot:
                                    style_str += $@" row{i}cell{j}Style.BorderTop = BorderStyle.DashDotDot; ";
                                    break;
                                case BorderStyle.MediumDashDotDot:
                                    style_str += $@" row{i}cell{j}Style.BorderTop = BorderStyle.MediumDashDotDot; ";
                                    break;
                                case BorderStyle.SlantedDashDot:
                                    style_str += $@" row{i}cell{j}Style.BorderTop = BorderStyle.SlantedDashDot; ";
                                    break;
                                default:
                                    style_str += $@" row{i}cell{j}Style.BorderTop = BorderStyle.None; ";
                                    break;
                            }

                            switch (cellStyle.FillPattern)
                            {
                                case FillPattern.NoFill:
                                    style_str += $@" row{i}cell{j}Style.FillPattern = FillPattern.NoFill; ";
                                    break;
                                case FillPattern.SolidForeground:
                                    style_str += $@" row{i}cell{j}Style.FillPattern = FillPattern.SolidForeground; ";
                                    break;
                                case FillPattern.FineDots:
                                    style_str += $@" row{i}cell{j}Style.FillPattern = FillPattern.FineDots; ";
                                    break;
                                case FillPattern.AltBars:
                                    style_str += $@" row{i}cell{j}Style.FillPattern = FillPattern.AltBars; ";
                                    break;
                                case FillPattern.SparseDots:
                                    style_str += $@" row{i}cell{j}Style.FillPattern = FillPattern.SparseDots; ";
                                    break;
                                case FillPattern.ThickHorizontalBands:
                                    style_str += $@" row{i}cell{j}Style.FillPattern = FillPattern.ThickHorizontalBands; ";
                                    break;
                                case FillPattern.ThickVerticalBands:
                                    style_str += $@" row{i}cell{j}Style.FillPattern = FillPattern.ThickVerticalBands; ";
                                    break;
                                case FillPattern.ThickBackwardDiagonals:
                                    style_str += $@" row{i}cell{j}Style.FillPattern = FillPattern.ThickBackwardDiagonals; ";
                                    break;
                                case FillPattern.ThickForwardDiagonals:
                                    style_str += $@" row{i}cell{j}Style.FillPattern = FillPattern.ThickForwardDiagonals; ";
                                    break;
                                case FillPattern.BigSpots:
                                    style_str += $@" row{i}cell{j}Style.FillPattern = FillPattern.BigSpots; ";
                                    break;
                                case FillPattern.Bricks:
                                    style_str += $@" row{i}cell{j}Style.FillPattern = FillPattern.Bricks; ";
                                    break;
                                case FillPattern.ThinHorizontalBands:
                                    style_str += $@" row{i}cell{j}Style.FillPattern = FillPattern.ThinHorizontalBands; ";
                                    break;
                                case FillPattern.ThinVerticalBands:
                                    style_str += $@" row{i}cell{j}Style.FillPattern = FillPattern.ThinVerticalBands; ";
                                    break;
                                case FillPattern.ThinBackwardDiagonals:
                                    style_str += $@" row{i}cell{j}Style.FillPattern = FillPattern.ThinBackwardDiagonals; ";
                                    break;
                                case FillPattern.ThinForwardDiagonals:
                                    style_str += $@" row{i}cell{j}Style.FillPattern = FillPattern.ThinForwardDiagonals; ";
                                    break;
                                case FillPattern.Squares:
                                    style_str += $@" row{i}cell{j}Style.FillPattern = FillPattern.Squares; ";
                                    break;
                                case FillPattern.Diamonds:
                                    style_str += $@" row{i}cell{j}Style.FillPattern = FillPattern.Diamonds; ";
                                    break;
                                case FillPattern.LessDots:
                                    style_str += $@" row{i}cell{j}Style.FillPattern = FillPattern.LessDots; ";
                                    break;
                                case FillPattern.LeastDots:
                                    style_str += $@" row{i}cell{j}Style.FillPattern = FillPattern.LeastDots; ";
                                    break;
                                default:
                                    style_str += $@" row{i}cell{j}Style.FillPattern = FillPattern.NoFill; ";
                                    break;
                            }

                            //获取线条/背前景颜色

                            XSSFColor fillcolor = cellStyle.FillForegroundXSSFColor;

                            if (fillcolor != null)
                            {
                                if (fillcolor.IsRGB)
                                {
                                    style_str += $@"row{i}cell{j}Style.FillForegroundXSSFColor = new XSSFColor(new byte[]{{{fillcolor.RGB[0]},{fillcolor.RGB[1]},{fillcolor.RGB[2]}}});";
                                }
                                else if(fillcolor.IsAuto)
                                {

                                }
                                
                            }

                            #region 给样式设置字体

                            if (cellStyle.GetFont() != null)
                            {
                                XSSFFont font = cellStyle.GetFont();
                                
                                string font_str = $@"XSSFFont font{i}{j} = (XSSFFont)workbook.CreateFont();font{i}{j}.FontName = ""{font.FontName}"";font{i}{j}.FontHeightInPoints = {font.FontHeightInPoints};font{i}{j}.Family = {font.Family};font{i}{j}.Charset = {font.Charset};";

                                if (font.IsBold)
                                    font_str += $@" font{i}{j}.IsBold = true;";
                                if (font.IsItalic)
                                    font_str += $@" font{i}{j}.IsItalic = true;";
                                if (font.IsStrikeout)
                                    font_str += $@" font{i}{j}.IsStrikeout = true;";

                                switch (font.Underline)
                                {
                                    case FontUnderlineType.None:
                                        font_str += $@" font{i}{j}.Underline = FontUnderlineType.None;";
                                        break;
                                    case FontUnderlineType.Single:
                                        font_str += $@" font{i}{j}.Underline = FontUnderlineType.Single;";
                                        break;
                                    case FontUnderlineType.Double:
                                        font_str += $@" font{i}{j}.Underline = FontUnderlineType.Double;";
                                        break;
                                    case FontUnderlineType.SingleAccounting:
                                        font_str += $@" font{i}{j}.Underline = FontUnderlineType.SingleAccounting;";
                                        break;
                                    case FontUnderlineType.DoubleAccounting:
                                        font_str += $@" font{i}{j}.Underline = FontUnderlineType.DoubleAccounting;";
                                        break;
                                    default:
                                        font_str += $@" font{i}{j}.Underline = FontUnderlineType.None;";
                                        break;
                                }

                                switch (font.TypeOffset)
                                {
                                    case FontSuperScript.None:
                                        font_str += $@"font{i}{j}.TypeOffset = FontSuperScript.None;";
                                        break;
                                    case FontSuperScript.Super:
                                        font_str += $@"font{i}{j}.TypeOffset = FontSuperScript.Super;";
                                        break;
                                    case FontSuperScript.Sub:
                                        font_str += $@"font{i}{j}.TypeOffset = FontSuperScript.Sub;";
                                        break;
                                    default:
                                        font_str += $@"font{i}{j}.TypeOffset = FontSuperScript.None;";
                                        break;
                                }

                                
                                if(font.GetXSSFColor()!=null)
                                {
                                    XSSFColor fontColor = font.GetXSSFColor();
                                    
                                    if (fontColor.IsAuto)
                                    {
                                        //是系统设置的话什么都不做
                                    }
                                    if (fontColor.IsRGB)
                                    {
                                        font_str += $@"font{i}{j}.SetColor(new XSSFColor(new byte[3]{{{fontColor.RGB[0]},{fontColor.RGB[1]},{fontColor.RGB[2]}}}));";
                                    }
                                }

                                //把字体设置到样式上
                                font_str += $@"row{i}cell{j}Style.SetFont(font{i}{j});";

                                builder.AppendLine(font_str);
                            }

                            #endregion

                            
                            if (cellStyle.WrapText == true)
                                style_str += $@" row{i}cell{j}Style.WrapText = true; ";
                            else
                                style_str += $@" row{i}cell{j}Style.WrapText = false; ";

                            if (cellStyle.ShrinkToFit == true)
                                style_str += $@" row{i}cell{j}Style.ShrinkToFit = true;";
                            else
                                style_str += $@" row{i}cell{j}Style.ShrinkToFit = false;";

                            

                            style_str += $@"row{i}cell{j}.CellStyle = row{i}cell{j}Style;";

                            builder.AppendLine(style_str);

                        }

                    }
                }
                
            }

            

            FileStream outline = new FileStream("../../NpoiCode.txt", FileMode.OpenOrCreate, FileAccess.ReadWrite);

            byte[] vs = System.Text.Encoding.UTF8.GetBytes(builder.ToString());

            outline.Write(vs, 0, vs.Length);

            

        }
    
        /// <summary>
        /// 判断Region合并单元格是否被设置了边框
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="region"></param>
        /// <returns></returns>
        public static (bool,List<string>) IsRegionHasBorder(ISheet sheet,CellRangeAddress region)
        {
            bool hasBoreder = false;

            //保存方向用
            List<string> which_direct = new List<string>();

            for (int i = region.FirstRow; i <= region.LastRow; i++)
            {
                for (int j = region.FirstColumn; j <= region.LastColumn; j++)
                {
                    ICellStyle style = sheet.GetRow(i).GetCell(j).CellStyle;

                    if (style.BorderLeft != BorderStyle.None)
                        which_direct.Add("left");
                    if (style.BorderRight != BorderStyle.None)
                        which_direct.Add("right");
                    if (style.BorderTop != BorderStyle.None)
                        which_direct.Add("top");
                    if (style.BorderBottom != BorderStyle.None)
                        which_direct.Add("bottom");
                    if (style.BorderBottom != BorderStyle.None
                            || style.BorderLeft != BorderStyle.None
                            || style.BorderRight != BorderStyle.None
                            || style.BorderTop != BorderStyle.None)
                    {
                        hasBoreder = true;
                        break;
                    }
                }
            }

            return (hasBoreder,which_direct);


        }


        /// <summary>
        /// 获取要合并列的位置[行位置]
        /// 生成By文心一言
        /// </summary>
        /// <param name="objlist"></param>
        /// <returns></returns>
        public static Dictionary<object,List<List<int>>> GetMergedPosition(object[] objlist){

            Dictionary<object, List<List<int>>> consecutiveOccurrences = new Dictionary<object, List<List<int>>>();  
  
            for (int i = 0; i < objlist.Length; i++)  
            {  
                var current = objlist[i];  
    
                // 检查下一个元素是否与当前元素相同  
                  // 检查下一个元素是否与当前元素相同  
                if (i < objlist.Length - 1 && objlist[i + 1] == current)  
                {  
                    // 如果下一个元素相同，找到所有连续的位置  
                    List<int> consecutivePositions = new List<int> { i };  
                    while (i < objlist.Length - 1 && objlist[i + 1] == current)  
                    {  
                        i++;  
                        consecutivePositions.Add(i);  
                    }  
            
                    // 添加到字典中  
                    if (consecutiveOccurrences.ContainsKey(current))  
                    {  
                        consecutiveOccurrences[current].Add(consecutivePositions);  
                    }  
                    else  
                    {  
                        consecutiveOccurrences.Add(current, new List<List<int>> { consecutivePositions });  
                    }  
                }  
                
                // 注意：上面的else块包含一个错误，它不应该重置i。  
                // 正确的做法是移除else块，因为当找到连续的字符串时，while循环已经处理了跳过它们。  
            }  


            return consecutiveOccurrences;

        }

    }
}