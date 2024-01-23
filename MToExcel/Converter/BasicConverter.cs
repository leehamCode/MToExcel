using MToExcel.Attributes;
using MToExcel.Exceptons;
using MToExcel.Models.Enums;
using MToExcel.Utils;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace MToExcel.Converter
{
    public class BasicConverter : MTConverter
    {

        /// <summary>
        /// 这个布尔变量控制打印的Excel的版本信息
        /// true =  07版
        /// false = 03版
        /// </summary>
        public bool Version { get; set; }

        /// <summary>
        /// 自定义表头函数
        /// </summary>
        /// <value>转化中需要导出的Workbook</value>
        /// <value>自定义表头需要占用的行数，如果没有，则默认一行</value>
        public Action<IWorkbook> CustomHeadMethod{get;set;} = null;

        /// <summary>
        /// 自定义表头需要占用的行数
        /// </summary>
        public int? CustomHeadRows {get;set;} = null;   //使用自定义表头必须设置


        /// <summary>
        /// 自定义Excel尾部函数
        /// </summary>
        /// <value>Excel本体</value>
        /// <value>单个Sheet的最后一行</value>
        public Action<IWorkbook,int> CustomTailMethod{get;set;} = null;

        /// <summary>
        /// 行打印前事件
        /// </summary>
        /// <value></value>
        public Action<IRow,Object> item_change_before_event { get; set; }


        /// <summary>
        /// 打印后事件
        /// </summary>
        /// <value></value>
        public Action<IRow,Object> item_change_after_event{get;set;}


        //这里放一些bool值用来检查是否有一些Class上的Attribute

        /// <summary>
        /// 是否有标题行
        /// </summary>
        private bool has_title_attr = false;

        /// <summary>
        /// 是否有行隐藏条件
        /// </summary>
        private bool has_condition_hide = false;

        /// <summary>
        /// 是否有冻结区域
        /// </summary>
        private bool has_freeze_area = false;

        /// <summary>
        /// 是否合并相同值
        /// </summary>
        private bool has_merge_near = false;

        public BasicConverter()
        {
            Version = true;
        }

        public IWorkbook ConvertToExcel<T>(List<T> list)
        {
            WrapperConverter.CellStylePool.Clear();

            IWorkbook workbook = null;

            if (Version)
                workbook = new XSSFWorkbook();
            else
                workbook = new HSSFWorkbook();

            ISheet defaultSheet = workbook.CreateSheet("SheetOne");

            //获取传递的泛型类型
            Type type = typeof(T);
            Check_Class_Attr(type);
            //首先判断泛型T是否为基础数据类型

            //如果泛型类型为基础数据类型,则为写一行数据
            if (isBasicType(type))
            {
                IRow uniqueRow = defaultSheet.CreateRow(0);
                int Count = 0;
                list.ForEach(item => {
                    uniqueRow.CreateCell(Count).SetCellValue(Convert.ToString(item));
                    Count++;
                });
                return workbook;
            }

            //如果不是基础数据类型就反射获取其属性写入Excel

            PropertyInfo[] properties = type.GetProperties();

            //-------------------------------------------------------------------------------------------------------------------分割线

            if(CustomHeadMethod!=null&&CustomHeadRows!=null)
            {
                if(CustomHeadRows<=0)
                {
                    throw new CustomHeadException("自定义表头长度必须大于0!");
                }
                CustomHeadMethod.Invoke(workbook);
            }
            
            
            
            
            int i = 0;

            //r如果有titleattr则保存头行开始的行数
            int title_attr_start = 0;

            //如果没有自定义头部则打印默认头部
            if(CustomHeadMethod==null)
            {

                
                if(has_title_attr)
                {
                    var title_attr =  (TitleAttribute)type.GetCustomAttribute(typeof(TitleAttribute));

                    if(title_attr.Col_Merge_num<=0||title_attr.Row_Merge_num<=0)
                    {
                        throw new TitleAttrException("合并行列不能小于等于0");
                    }

                    title_attr_start = title_attr.Row_Merge_num;

                    //标题行不合并单元格的情况
                    if(title_attr.Col_Merge_num==1&&title_attr.Row_Merge_num==1)
                    {
                        IRow title = defaultSheet.CreateRow(0);

                        //设置一下表头样式,将表头设置为加粗字体
                        ICellStyle t_style = workbook.CreateCellStyle();
                        IFont t_font = workbook.CreateFont();

                        //标题背景颜色
                        if(title_attr.Back_color!=null)
                        {
                            if(title_attr.Back_color.Length!=3){ throw new RgbArrayException("颜色数组固定长度为3"); }
                            if(Version)
                            {
                                ((XSSFCellStyle)t_style).FillForegroundXSSFColor = new XSSFColor(title_attr.Back_color);
                            }
                            else
                            {
                                ((HSSFCellStyle)t_style).FillForegroundColor = HSSFColor.ToHSSFColor(new XSSFColor(title_attr.Back_color)).Indexed;
                            }
                        }

                        if(title_attr.Fore_color!=null)
                        {
                            if(title_attr.Fore_color.Length!=3){ throw new RgbArrayException("颜色数组固定长度为3"); }
                            if(Version)
                            {
                                ((XSSFCellStyle)t_style).FillBackgroundXSSFColor = new XSSFColor(title_attr.Fore_color);
                            }
                            else
                            {
                                ((HSSFCellStyle)t_style).FillBackgroundColor = HSSFColor.ToHSSFColor(new XSSFColor(title_attr.Fore_color)).Indexed;
                            }
                        }

                        //标题字体名
                        if(!string.IsNullOrEmpty(title_attr.Font_Name))
                        {
                            t_font.FontName = title_attr.Font_Name;
                        }
                        
                        //标题字体颜色
                        if(title_attr.Font_color!=null)
                        {
                            if(title_attr.Font_color.Length!=3){ throw new RgbArrayException("颜色数组固定长度为3"); }
                            if(Version)
                            {
                                ((XSSFFont)t_font).SetColor(new XSSFColor(title_attr.Font_color));
                            }
                            else
                            {
                                ((HSSFFont)t_font).Color = HSSFColor.ToHSSFColor(new XSSFColor(title_attr.Font_color)).Indexed;
                            }
                        }
                        
                        //字体大小
                        if(title_attr.Font_Size>0)
                        {
                            t_font.FontHeightInPoints = title_attr.Font_Size;
                        }
                        else{
                            throw new Exception("字体大小不能为负数");
                        }

                        //字体属性
                        if(title_attr.IsBold){t_font.IsBold = true;}
                        if(title_attr.IsItalic){t_font.IsItalic = true;}

                        t_style.SetFont(t_font);

                        var titlecell =  title.CreateCell(0);

                        //解析字符串中的通配符
                        titlecell.SetCellValue(ExpressionHelp.Read_Title_Content<T>(title_attr.Context,list,true));

                        //此处检查其他标签
                        //-------------------

                        #region 

                        if(type.GetCustomAttribute(typeof(HorizonAttribute))!=null)
                        {
                            var horzion =  (HorizonAttribute)type.GetCustomAttribute(typeof(HorizonAttribute));

                            if(horzion.horizon == Models.Enums.Horizon.Center)
                            {
                                t_style.Alignment = HorizontalAlignment.Center;
                            }
                            else if(horzion.horizon == Models.Enums.Horizon.Left)
                            {
                                t_style.Alignment = HorizontalAlignment.Left;
                            }
                            else
                            {
                                t_style.Alignment = HorizontalAlignment.Right;
                            }

                            if(horzion.verticalHorizon == Models.Enums.VerticalHorizon.Up)
                            {
                                t_style.VerticalAlignment = VerticalAlignment.Top;
                            }
                            else if(horzion.verticalHorizon == Models.Enums.VerticalHorizon.Mid)
                            {
                                t_style.VerticalAlignment = VerticalAlignment.Center;
                            }
                            else 
                            {
                                t_style.VerticalAlignment = VerticalAlignment.Bottom;
                            }

                        }

                        if(type.GetCustomAttribute(typeof(BorderStyleAttribute))!=null)
                        {
                            var borders = (IEnumerable<BorderStyleAttribute>)type.GetCustomAttributes(typeof(BorderStyleAttribute));

                            var distinct_list =  borders.DistinctBy(it=>it.Direct).ToList();

                            distinct_list.ForEach(item=>{
                                
                                if(item.Direct == BorderDirect.Upper)
                                {
                                    if(item.Wid == BorderWid.NoneBorder)
                                    {
                                        // DoNothing 无边框什么都不做
                                    }
                                    else if(item.Wid == BorderWid.ThinBorder)
                                    {
                                        t_style.BorderTop = BorderStyle.Thin;
                                    }
                                    else if(item.Wid == BorderWid.MiddBorder)
                                    {
                                        t_style.BorderTop = BorderStyle.Medium;
                                    }
                                    else
                                    {
                                        t_style.BorderTop = BorderStyle.Thick;
                                    }

                                    if(item.Color!=null)
                                    {
                                        XSSFColor color = new  XSSFColor(item.Color);
                                        ((XSSFCellStyle)t_style).SetTopBorderColor(color);
                                    }
                                }
                                else if(item.Direct == BorderDirect.Bottom)
                                {
                                    if(item.Wid == BorderWid.NoneBorder)
                                    {
                                        // DoNothing 无边框什么都不做
                                    }
                                    else if(item.Wid == BorderWid.ThinBorder)
                                    {
                                        t_style.BorderBottom = BorderStyle.Thin;
                                    }
                                    else if(item.Wid == BorderWid.MiddBorder)
                                    {
                                        t_style.BorderBottom = BorderStyle.Medium;
                                    }
                                    else
                                    {
                                        t_style.BorderBottom = BorderStyle.Thick;
                                    }

                                    if(item.Color!=null)
                                    {
                                        XSSFColor color = new  XSSFColor(item.Color);
                                        ((XSSFCellStyle)t_style).SetBottomBorderColor(color);
                                    }
                                }
                                else if(item.Direct == BorderDirect.Right)
                                {
                                    if(item.Wid == BorderWid.NoneBorder)
                                    {
                                        // DoNothing 无边框什么都不做
                                    }
                                    else if(item.Wid == BorderWid.ThinBorder)
                                    {
                                        t_style.BorderRight = BorderStyle.Thin;
                                    }
                                    else if(item.Wid == BorderWid.MiddBorder)
                                    {
                                        t_style.BorderRight = BorderStyle.Medium;
                                    }
                                    else
                                    {
                                        t_style.BorderRight = BorderStyle.Thick;
                                    }

                                    if(item.Color!=null)
                                    {
                                        XSSFColor color = new  XSSFColor(item.Color);
                                        ((XSSFCellStyle)t_style).SetRightBorderColor(color);
                                    }
                                }
                                else if(item.Direct == BorderDirect.Left)
                                {
                                    if(item.Wid == BorderWid.NoneBorder)
                                    {
                                        // DoNothing 无边框什么都不做
                                    }
                                    else if(item.Wid == BorderWid.ThinBorder)
                                    {
                                        t_style.BorderLeft = BorderStyle.Thin;
                                    }
                                    else if(item.Wid == BorderWid.MiddBorder)
                                    {
                                        t_style.BorderLeft = BorderStyle.Medium;
                                    }
                                    else
                                    {
                                        t_style.BorderLeft = BorderStyle.Thick;
                                    }

                                    if(item.Color!=null)
                                    {
                                        XSSFColor color = new  XSSFColor(item.Color);
                                        ((XSSFCellStyle)t_style).SetLeftBorderColor(color);
                                    }
                                }
                                else if(item.Direct == BorderDirect.Diagonal_slash)
                                {
                                    t_style.BorderDiagonal = BorderDiagonal.Forward;

                                    if(item.Wid == BorderWid.NoneBorder)
                                    {
                                        // DoNothing 无边框什么都不做
                                    }
                                    else if(item.Wid == BorderWid.ThinBorder)
                                    {
                                        t_style.BorderDiagonalLineStyle = BorderStyle.Thin;
                                    }
                                    else if(item.Wid == BorderWid.MiddBorder)
                                    {
                                        t_style.BorderDiagonalLineStyle = BorderStyle.Medium;
                                    }
                                    else
                                    {
                                        t_style.BorderDiagonalLineStyle = BorderStyle.Thick;
                                    }

                                    if(item.Color!=null)
                                    {
                                        XSSFColor color = new  XSSFColor(item.Color);
                                        ((XSSFCellStyle)t_style).SetDiagonalBorderColor(color);
                                    }
                                }
                                else
                                {
                                    t_style.BorderDiagonal = BorderDiagonal.Backward;

                                    if(item.Wid == BorderWid.NoneBorder)
                                    {
                                        // DoNothing 无边框什么都不做
                                    }
                                    else if(item.Wid == BorderWid.ThinBorder)
                                    {
                                        t_style.BorderDiagonalLineStyle = BorderStyle.Thin;
                                    }
                                    else if(item.Wid == BorderWid.MiddBorder)
                                    {
                                        t_style.BorderDiagonalLineStyle = BorderStyle.Medium;
                                    }
                                    else
                                    {
                                        t_style.BorderDiagonalLineStyle = BorderStyle.Thick;
                                    }

                                    if(item.Color!=null)
                                    {
                                        XSSFColor color = new  XSSFColor(item.Color);
                                        ((XSSFCellStyle)t_style).SetDiagonalBorderColor(color);
                                    }
                                }

                            });
                        }


                        #endregion

                        


                        titlecell.CellStyle = t_style;
                    }
                    else{
                        IRow title = defaultSheet.CreateRow(0);

                        //设置一下表头样式,将表头设置为加粗字体
                        ICellStyle t_style = workbook.CreateCellStyle();
                        IFont t_font = workbook.CreateFont();

                        //标题背景颜色
                        if(title_attr.Back_color!=null)
                        {
                            if(title_attr.Back_color.Length!=3){ throw new RgbArrayException("颜色数组固定长度为3"); }
                            if(Version)
                            {
                                ((XSSFCellStyle)t_style).FillForegroundXSSFColor = new XSSFColor(title_attr.Back_color);
                            }
                            else
                            {
                                ((HSSFCellStyle)t_style).FillForegroundColor = HSSFColor.ToHSSFColor(new XSSFColor(title_attr.Back_color)).Indexed;
                            }
                        }

                        if(title_attr.Fore_color!=null)
                        {
                            if(title_attr.Fore_color.Length!=3){ throw new RgbArrayException("颜色数组固定长度为3"); }
                            if(Version)
                            {
                                ((XSSFCellStyle)t_style).FillBackgroundXSSFColor = new XSSFColor(title_attr.Fore_color);
                            }
                            else
                            {
                                ((HSSFCellStyle)t_style).FillBackgroundColor = HSSFColor.ToHSSFColor(new XSSFColor(title_attr.Fore_color)).Indexed;
                            }
                        }

                        //标题字体名
                        if(!string.IsNullOrEmpty(title_attr.Font_Name))
                        {
                            t_font.FontName = title_attr.Font_Name;
                        }
                        
                        //标题字体颜色
                        if(title_attr.Font_color!=null)
                        {
                            if(title_attr.Font_color.Length!=3){ throw new RgbArrayException("颜色数组固定长度为3"); }
                            if(Version)
                            {
                                ((XSSFFont)t_font).SetColor(new XSSFColor(title_attr.Back_color));
                            }
                            else
                            {
                                ((HSSFFont)t_font).Color = HSSFColor.ToHSSFColor(new XSSFColor(title_attr.Back_color)).Indexed;
                            }
                        }
                        
                        //字体大小
                        if(title_attr.Font_Size>0)
                        {
                            t_font.FontHeightInPoints = title_attr.Font_Size;
                        }
                        else{
                            throw new Exception("字体大小不能为负数");
                        }

                        //字体属性
                        if(title_attr.IsBold){t_font.IsBold = true;}
                        if(title_attr.IsItalic){t_font.IsItalic = true;}

                        t_style.SetFont(t_font);

                        var titlecell =  title.CreateCell(0);

                        titlecell.SetCellValue(ExpressionHelp.Read_Title_Content<T>(title_attr.Context,list,true));

                        CellRangeAddress titleBox = new CellRangeAddress(0,title_attr.Row_Merge_num-1,0,title_attr.Col_Merge_num-1);

                        titlecell.Row.Sheet.AddMergedRegion(titleBox);

                        

                        //此处检查其他Attriubte
                        //Do Something--------------------------------------------------

                        #region 

                        //是否对齐
                        if(type.GetCustomAttribute(typeof(HorizonAttribute))!=null)
                        {
                            var horzion =  (HorizonAttribute)type.GetCustomAttribute(typeof(HorizonAttribute));

                            if(horzion.horizon == Models.Enums.Horizon.Center)
                            {
                                t_style.Alignment = HorizontalAlignment.Center;
                            }
                            else if(horzion.horizon == Models.Enums.Horizon.Left)
                            {
                                t_style.Alignment = HorizontalAlignment.Left;
                            }
                            else
                            {
                                t_style.Alignment = HorizontalAlignment.Right;
                            }

                            if(horzion.verticalHorizon == Models.Enums.VerticalHorizon.Up)
                            {
                                t_style.VerticalAlignment = VerticalAlignment.Top;
                            }
                            else if(horzion.verticalHorizon == Models.Enums.VerticalHorizon.Mid)
                            {
                                t_style.VerticalAlignment = VerticalAlignment.Center;
                            }
                            else 
                            {
                                t_style.VerticalAlignment = VerticalAlignment.Bottom;
                            }

                        }
                        
                        /*
                            合并单元格后就不能用一般的边框设置了
                        if(type.GetCustomAttribute(typeof(BorderStyleAttribute))!=null)
                        {
                            var borders = (IEnumerable<BorderStyleAttribute>)type.GetCustomAttributes(typeof(BorderStyleAttribute));

                            var distinct_list =  borders.DistinctBy(it=>it.Direct).ToList();

                            distinct_list.ForEach(item=>{
                                
                                if(item.Direct == BorderDirect.Upper)
                                {
                                    if(item.Wid == BorderWid.NoneBorder)
                                    {
                                        // DoNothing 无边框什么都不做
                                    }
                                    else if(item.Wid == BorderWid.ThinBorder)
                                    {
                                        t_style.BorderTop = BorderStyle.Thin;
                                    }
                                    else if(item.Wid == BorderWid.MiddBorder)
                                    {
                                        t_style.BorderTop = BorderStyle.Medium;
                                    }
                                    else
                                    {
                                        t_style.BorderTop = BorderStyle.Thick;
                                    }

                                    if(item.Color!=null)
                                    {
                                        XSSFColor color = new  XSSFColor(item.Color);
                                        ((XSSFCellStyle)t_style).SetTopBorderColor(color);
                                    }
                                }
                                else if(item.Direct == BorderDirect.Bottom)
                                {
                                    if(item.Wid == BorderWid.NoneBorder)
                                    {
                                        // DoNothing 无边框什么都不做
                                    }
                                    else if(item.Wid == BorderWid.ThinBorder)
                                    {
                                        t_style.BorderBottom = BorderStyle.Thin;
                                    }
                                    else if(item.Wid == BorderWid.MiddBorder)
                                    {
                                        t_style.BorderBottom = BorderStyle.Medium;
                                    }
                                    else
                                    {
                                        t_style.BorderBottom = BorderStyle.Thick;
                                    }

                                    if(item.Color!=null)
                                    {
                                        XSSFColor color = new  XSSFColor(item.Color);
                                        ((XSSFCellStyle)t_style).SetBottomBorderColor(color);
                                    }
                                }
                                else if(item.Direct == BorderDirect.Right)
                                {
                                    if(item.Wid == BorderWid.NoneBorder)
                                    {
                                        // DoNothing 无边框什么都不做
                                    }
                                    else if(item.Wid == BorderWid.ThinBorder)
                                    {
                                        t_style.BorderRight = BorderStyle.Thin;
                                    }
                                    else if(item.Wid == BorderWid.MiddBorder)
                                    {
                                        t_style.BorderRight = BorderStyle.Medium;
                                    }
                                    else
                                    {
                                        t_style.BorderRight = BorderStyle.Thick;
                                    }

                                    if(item.Color!=null)
                                    {
                                        XSSFColor color = new  XSSFColor(item.Color);
                                        ((XSSFCellStyle)t_style).SetRightBorderColor(color);
                                    }
                                }
                                else if(item.Direct == BorderDirect.Left)
                                {
                                    if(item.Wid == BorderWid.NoneBorder)
                                    {
                                        // DoNothing 无边框什么都不做
                                    }
                                    else if(item.Wid == BorderWid.ThinBorder)
                                    {
                                        t_style.BorderLeft = BorderStyle.Thin;
                                    }
                                    else if(item.Wid == BorderWid.MiddBorder)
                                    {
                                        t_style.BorderLeft = BorderStyle.Medium;
                                    }
                                    else
                                    {
                                        t_style.BorderLeft = BorderStyle.Thick;
                                    }

                                    if(item.Color!=null)
                                    {
                                        XSSFColor color = new  XSSFColor(item.Color);
                                        ((XSSFCellStyle)t_style).SetLeftBorderColor(color);
                                    }
                                }
                                else if(item.Direct == BorderDirect.Diagonal_slash)
                                {
                                    t_style.BorderDiagonal = BorderDiagonal.Forward;

                                    if(item.Wid == BorderWid.NoneBorder)
                                    {
                                        // DoNothing 无边框什么都不做
                                    }
                                    else if(item.Wid == BorderWid.ThinBorder)
                                    {
                                        t_style.BorderDiagonalLineStyle = BorderStyle.Thin;
                                    }
                                    else if(item.Wid == BorderWid.MiddBorder)
                                    {
                                        t_style.BorderDiagonalLineStyle = BorderStyle.Medium;
                                    }
                                    else
                                    {
                                        t_style.BorderDiagonalLineStyle = BorderStyle.Thick;
                                    }

                                    if(item.Color!=null)
                                    {
                                        XSSFColor color = new  XSSFColor(item.Color);
                                        ((XSSFCellStyle)t_style).SetDiagonalBorderColor(color);
                                    }
                                }
                                else
                                {
                                    t_style.BorderDiagonal = BorderDiagonal.Backward;

                                    if(item.Wid == BorderWid.NoneBorder)
                                    {
                                        // DoNothing 无边框什么都不做
                                    }
                                    else if(item.Wid == BorderWid.ThinBorder)
                                    {
                                        t_style.BorderDiagonalLineStyle = BorderStyle.Thin;
                                    }
                                    else if(item.Wid == BorderWid.MiddBorder)
                                    {
                                        t_style.BorderDiagonalLineStyle = BorderStyle.Medium;
                                    }
                                    else
                                    {
                                        t_style.BorderDiagonalLineStyle = BorderStyle.Thick;
                                    }

                                    if(item.Color!=null)
                                    {
                                        XSSFColor color = new  XSSFColor(item.Color);
                                        ((XSSFCellStyle)t_style).SetDiagonalBorderColor(color);
                                    }
                                }

                            });
                        }
                        
                        
                        */
                        


                        #endregion

                        
                        var bordersets =  (IEnumerable<BorderStyleAttribute>)type.GetCustomAttributes(typeof(BorderStyleAttribute));

                        //设置合并单元格边框
                        bordersets.DistinctBy(it=>it.Direct).ToList().ForEach(item=>{

                            if(item.Direct==BorderDirect.Upper)
                            {
                               switch (item.Wid)
                               {
                                
                                case BorderWid.MiddBorder:{
                                    RegionUtil.SetBorderTop((int)BorderStyle.Medium,titleBox,titlecell.Row.Sheet);
                                    if(item.Color!=null){
                                        var color =  new XSSFColor(item.Color);
                                        RegionUtil.SetTopBorderColor(color.Index,titleBox,titlecell.Row.Sheet);
                                    }
                                    break;
                                }
                                case BorderWid.ThickBorder:{
                                    RegionUtil.SetBorderTop((int)BorderStyle.Thick,titleBox,titlecell.Row.Sheet);
                                    if(item.Color!=null){
                                        var color =  new XSSFColor(item.Color);
                                        RegionUtil.SetTopBorderColor(color.Index,titleBox,titlecell.Row.Sheet);
                                    }
                                    break;
                                }
                                case BorderWid.ThinBorder:{
                                    RegionUtil.SetBorderTop((int)BorderStyle.Thin,titleBox,titlecell.Row.Sheet);
                                    if(item.Color!=null){
                                        var color =  new XSSFColor(item.Color);
                                        RegionUtil.SetTopBorderColor(color.Index,titleBox,titlecell.Row.Sheet);
                                    }
                                    break;
                                }
                                case BorderWid.NoneBorder:{
                                    RegionUtil.SetBorderTop((int)BorderStyle.None,titleBox,titlecell.Row.Sheet);
                                    if(item.Color!=null){
                                        var color =  new XSSFColor(item.Color);
                                        RegionUtil.SetTopBorderColor(color.Index,titleBox,titlecell.Row.Sheet);
                                    }
                                    break;
                                }
                                default:
                                    break;
                               }
                            }
                            else if(item.Direct==BorderDirect.Bottom)
                            {
                               switch (item.Wid)
                               {
                                
                                case BorderWid.MiddBorder:{
                                    RegionUtil.SetBorderBottom((int)BorderStyle.Medium,titleBox,titlecell.Row.Sheet);
                                    if(item.Color!=null){
                                        var color =  new XSSFColor(item.Color);
                                        RegionUtil.SetBottomBorderColor(color.Index,titleBox,titlecell.Row.Sheet);
                                    }
                                    break;
                                }
                                case BorderWid.ThickBorder:{
                                    RegionUtil.SetBorderBottom((int)BorderStyle.Thick,titleBox,titlecell.Row.Sheet);
                                    if(item.Color!=null){
                                        var color =  new XSSFColor(item.Color);
                                        RegionUtil.SetBottomBorderColor(color.Index,titleBox,titlecell.Row.Sheet);
                                    }
                                    break;
                                }
                                case BorderWid.ThinBorder:{
                                    RegionUtil.SetBorderBottom((int)BorderStyle.Thin,titleBox,titlecell.Row.Sheet);
                                    if(item.Color!=null){
                                        var color =  new XSSFColor(item.Color);
                                        RegionUtil.SetBottomBorderColor(color.Index,titleBox,titlecell.Row.Sheet);
                                    }
                                    break;
                                }
                                case BorderWid.NoneBorder:{
                                    RegionUtil.SetBorderBottom((int)BorderStyle.None,titleBox,titlecell.Row.Sheet);
                                    if(item.Color!=null){
                                        var color =  new XSSFColor(item.Color);
                                        RegionUtil.SetBottomBorderColor(color.Index,titleBox,titlecell.Row.Sheet);
                                    }
                                    break;
                                }
                                default:
                                    break;
                               }
                            }
                            else if(item.Direct==BorderDirect.Left)
                            {
                                switch (item.Wid)
                               {
                                
                                case BorderWid.MiddBorder:{
                                    RegionUtil.SetBorderLeft((int)BorderStyle.Medium,titleBox,titlecell.Row.Sheet);
                                    if(item.Color!=null){
                                        var color =  new XSSFColor(item.Color);
                                        RegionUtil.SetLeftBorderColor(color.Index,titleBox,titlecell.Row.Sheet);
                                    }
                                    break;
                                }
                                case BorderWid.ThickBorder:{
                                    RegionUtil.SetBorderLeft((int)BorderStyle.Thick,titleBox,titlecell.Row.Sheet);
                                    if(item.Color!=null){
                                        var color =  new XSSFColor(item.Color);
                                        RegionUtil.SetLeftBorderColor(color.Index,titleBox,titlecell.Row.Sheet);
                                    }
                                    break;
                                }
                                case BorderWid.ThinBorder:{
                                    RegionUtil.SetBorderLeft((int)BorderStyle.Thin,titleBox,titlecell.Row.Sheet);
                                    if(item.Color!=null){
                                        var color =  new XSSFColor(item.Color);
                                        RegionUtil.SetLeftBorderColor(color.Index,titleBox,titlecell.Row.Sheet);
                                    }
                                    break;
                                }
                                case BorderWid.NoneBorder:{
                                    RegionUtil.SetBorderLeft((int)BorderStyle.None,titleBox,titlecell.Row.Sheet);
                                    if(item.Color!=null){
                                        var color =  new XSSFColor(item.Color);
                                        RegionUtil.SetLeftBorderColor(color.Index,titleBox,titlecell.Row.Sheet);
                                    }
                                    break;
                                }
                                default:
                                    break;
                               }
                            }
                            else if(item.Direct==BorderDirect.Diagonal_back_slash)
                            {
                                switch (item.Wid)
                               {
                                
                                case BorderWid.MiddBorder:{
                                
                                    //RegionUtil.SetBorderBottom((int)BorderStyle.Medium,titleBox,titlecell.Row.Sheet);
                                    t_style.BorderDiagonal = BorderDiagonal.Backward;
                                    t_style.BorderDiagonalLineStyle = BorderStyle.Medium;
                                    if(item.Color!=null){
                                        t_style.BorderDiagonalColor = (new XSSFColor(item.Color)).Index;
                                    }

                                    break;
                                }
                                case BorderWid.ThickBorder:{

                                    //RegionUtil.SetBorderLeft((int)BorderStyle.Thick,titleBox,titlecell.Row.Sheet);
                                    t_style.BorderDiagonal = BorderDiagonal.Backward;
                                    t_style.BorderDiagonalLineStyle = BorderStyle.Medium;
                                    if(item.Color!=null){
                                        t_style.BorderDiagonalColor = (new XSSFColor(item.Color)).Index;
                                    }

                                    break;
                                }
                                case BorderWid.ThinBorder:{
                                    //RegionUtil.SetBorderLeft((int)BorderStyle.Thin,titleBox,titlecell.Row.Sheet);
                                    t_style.BorderDiagonal = BorderDiagonal.Backward;
                                    t_style.BorderDiagonalLineStyle = BorderStyle.Thin;
                                    if(item.Color!=null){
                                        t_style.BorderDiagonalColor = (new XSSFColor(item.Color)).Index;
                                    }
                                    break;
                                }
                                case BorderWid.NoneBorder:{
                                    //RegionUtil.SetBorderLeft((int)BorderStyle.None,titleBox,titlecell.Row.Sheet);
                                    t_style.BorderDiagonal = BorderDiagonal.Backward;
                                    t_style.BorderDiagonalLineStyle = BorderStyle.None;
                                    if(item.Color!=null){
                                        t_style.BorderDiagonalColor = (new XSSFColor(item.Color)).Index;
                                    }
                                    break;
                                }
                                default:
                                    break;
                               }
                            }
                            else if(item.Direct==BorderDirect.Diagonal_slash)
                            {
                                 switch (item.Wid)
                               {
                                
                                case BorderWid.MiddBorder:{
                                    //RegionUtil.SetBorderRight((int)BorderStyle.Medium,titleBox,titlecell.Row.Sheet);
                                    t_style.BorderDiagonal = BorderDiagonal.Forward;
                                    t_style.BorderDiagonalLineStyle = BorderStyle.Medium;
                                    if(item.Color!=null){
                                        t_style.BorderDiagonalColor = (new XSSFColor(item.Color)).Index;
                                    }
                                    break;
                                }
                                case BorderWid.ThickBorder:{
                                    //RegionUtil.SetBorderRight((int)BorderStyle.Thick,titleBox,titlecell.Row.Sheet);
                                    t_style.BorderDiagonal = BorderDiagonal.Forward;
                                    t_style.BorderDiagonalLineStyle = BorderStyle.Thick;
                                    if(item.Color!=null){
                                        t_style.BorderDiagonalColor = (new XSSFColor(item.Color)).Index;
                                    }
                                    break;
                                }
                                case BorderWid.ThinBorder:{
                                    //RegionUtil.SetBorderRight((int)BorderStyle.Thin,titleBox,titlecell.Row.Sheet);
                                    t_style.BorderDiagonal = BorderDiagonal.Forward;
                                    t_style.BorderDiagonalLineStyle = BorderStyle.Thin;
                                    if(item.Color!=null){
                                        t_style.BorderDiagonalColor = (new XSSFColor(item.Color)).Index;
                                    }
                                    break;
                                }
                                case BorderWid.NoneBorder:{
                                    //RegionUtil.SetBorderRight((int)BorderStyle.None,titleBox,titlecell.Row.Sheet);
                                    t_style.BorderDiagonal = BorderDiagonal.Forward;
                                    t_style.BorderDiagonalLineStyle = BorderStyle.None;
                                    if(item.Color!=null){
                                        t_style.BorderDiagonalColor = (new XSSFColor(item.Color)).Index;
                                    }
                                    break;
                                }
                                default:
                                    break;
                               }
                            }
                            else
                            {
                               switch (item.Wid)
                               {
                                
                                case BorderWid.MiddBorder:{
                                    RegionUtil.SetBorderRight((int)BorderStyle.Medium,titleBox,titlecell.Row.Sheet);
                                    if(item.Color!=null){
                                        var color =  new XSSFColor(item.Color);
                                        RegionUtil.SetBorderRight(color.Index,titleBox,titlecell.Row.Sheet);
                                    }
                                    break;
                                }
                                case BorderWid.ThickBorder:{
                                    RegionUtil.SetBorderRight((int)BorderStyle.Thick,titleBox,titlecell.Row.Sheet);
                                    if(item.Color!=null){
                                        var color =  new XSSFColor(item.Color);
                                        RegionUtil.SetBorderRight(color.Index,titleBox,titlecell.Row.Sheet);
                                    }
                                    break;
                                }
                                case BorderWid.ThinBorder:{
                                    RegionUtil.SetBorderRight((int)BorderStyle.Thin,titleBox,titlecell.Row.Sheet);
                                    if(item.Color!=null){
                                        var color =  new XSSFColor(item.Color);
                                        RegionUtil.SetBorderRight(color.Index,titleBox,titlecell.Row.Sheet);
                                    }
                                    break;
                                }
                                case BorderWid.NoneBorder:{
                                    RegionUtil.SetBorderRight((int)BorderStyle.None,titleBox,titlecell.Row.Sheet);
                                    if(item.Color!=null){
                                        var color =  new XSSFColor(item.Color);
                                        RegionUtil.SetBorderRight(color.Index,titleBox,titlecell.Row.Sheet);
                                    }
                                    break;
                                }
                                default:
                                    break;
                               }
                            }


                        });


                        titlecell.CellStyle = t_style;
                    }
                }


                //先用属性名打印一行表头

                IRow header = defaultSheet.CreateRow(title_attr_start);

                //设置一下表头样式,将表头设置为加粗字体
                ICellStyle style = workbook.CreateCellStyle();
                var Font = workbook.CreateFont();
                Font.IsBold = true;
                style.SetFont(Font);
                
                //int i = 0;
                

                foreach(PropertyInfo pro in properties)
                {
                    
                    
                    
                    if(pro.GetCustomAttribute(typeof(IgnoreType))!=null)
                    {
                        //如果在忽略类型中就直接Continue，开始下一轮循环
                        continue;
                    }

                    
                    if(pro.GetCustomAttribute(typeof(HeaderName))!=null)
                    {
                        HeaderName name = (HeaderName)pro.GetCustomAttribute(typeof(HeaderName));

                        header.CreateCell(i).SetCellValue(name.getCustomProName());
                        header.GetCell(i).CellStyle = style;
                        i++;
                        continue;
                    }

                    //判断泛型的该属性是否在(引用)标记类型池中
                    if(pro.GetCustomAttribute(typeof(ReferenceType))!=null)
                    {
                        ReferenceType refer = WrapperConverter.TypePool.GetValueOrDefault(pro.PropertyType);

                        if (refer.getIsMultiPart()) //判断是否要将引用类型拆成多列 :多列
                        {
                            PropertyInfo[] pros = pro.PropertyType.GetProperties();
                            
                            //PropertyInfo.PropertyType 可以属性的Type信息
                            //PropertyInfo.DeclaredType 可以取出定义这个属性的类型信息
                            //再将属性类型的属性全部取出

                            
                            foreach(PropertyInfo pi in pros)
                            {
                                header.CreateCell(i).SetCellValue(Convert.ToString(pi.DeclaringType.Name+":"+pi.Name));
                                header.GetCell(i).CellStyle = style;
                                i++;
                            }
                            //如果打印了引用类型的属性，需要Continue跳一下循环，避免再次打印该类型（Type）的信息
                            continue;
                        }
                        else  //~：单列,打印单列表头的话，不需要额外增加列数，所以直接退出循环即可
                        {
                            header.CreateCell(i).SetCellValue(Convert.ToString(pro.Name));
                            continue;
                        }
                    }

                    header.CreateCell(i).SetCellValue(pro.Name);
                    header.GetCell(i).CellStyle = style;
                    i++;
                }
            }
            

            


            //开始处理表体部分


            int RowNumber = 1;            //控制行号增加的变量
            if(has_title_attr)
            {
                RowNumber = title_attr_start+1;
            }

            if(CustomHeadMethod!=null)    //如果有自定表头，则从自定义表头占用行的下一行开始表体
            {
                RowNumber = CustomHeadRows.Value;
            }
            list.ForEach(item => {

                IRow row = defaultSheet.CreateRow(RowNumber); //创建一行写一行的数据

                PropertyInfo[] properties = item.GetType().GetProperties();

                //处理行打印前数据
                if(item_change_before_event!=null)
                {
                    item_change_before_event.Invoke(row,item);
                }

                int ColumnNumber = 0;     //控制列增加的变量
                foreach (PropertyInfo pro in properties)
                {
                    Type temp = pro.PropertyType;

                    
                    if(pro.GetCustomAttribute(typeof(IgnoreType)) != null)
                    {
                        //如果在表体上的话，这个循环就不需要了，可以直接退出这一层循环
                        continue;
                    }

                    

                    //判断泛型的该属性是否在（引用）标记类型池中
                    if(pro.GetCustomAttribute(typeof(ReferenceType)) != null)
                    {
                        ReferenceType refer = WrapperConverter.TypePool.GetValueOrDefault(pro.PropertyType);

                        if (refer.getIsMultiPart())   //多列打印引用类型的值
                        {
                            PropertyInfo[] pros = pro.PropertyType.GetProperties();
                            
                            foreach(PropertyInfo property in pros)
                            {
                                if (property.GetValue(pro.GetValue(item)) == null)
                                {
                                    row.CreateCell(ColumnNumber).SetCellValue("空值属性");
                                }
                                else
                                {
                                    row.CreateCell(ColumnNumber).SetCellValue(Convert.ToString(property.GetValue(pro.GetValue(item))));
                                }
                                ColumnNumber++;
                            }
                            
                        }
                        else
                        {
                            PropertyInfo[] pros = pro.PropertyType.GetProperties();

                            //单列的话，直接都追加到一列里去
                            string appending = "";

                            foreach (PropertyInfo property in pros)
                            {
                                if (property.GetValue(pro.GetValue(item)) == null)
                                {
                                    appending += "空置属性|";
                                }
                                else
                                {
                                    appending += (Convert.ToString(property.GetValue(pro.GetValue(item))) + '|');
                                }
                                
                            }
                            row.CreateCell(i).SetCellValue(appending);

                        }
                        continue;//同样在打印引用类型的属性完成后，需要跳一下循环，防止再打印一遍全限定名
                    }


                    if (pro.GetValue(item) == null)   //在这里进行属性判空
                    {
                        
                        row.CreateCell(ColumnNumber).SetCellValue("空值属性");
                        ColumnNumber++;
                    }
                    else if (isBasicArrayType(pro.GetValue(item).GetType())) //判断属性的类型是否为基础数据类型数组,这里先把数组的内容全写到一个格子中
                    {

                        //试一下是否能在反射中将属性值强制转化为数组,这里将所有数组的数据都写到一格数据中,以后应该提供更多的方式

                        row.CreateCell(ColumnNumber).SetCellValue("");

                        string appendingStr = "";

                        Array unknownArray = (Array)pro.GetValue(item);   //將數組屬性轉化為Array類型進行遍歷
                        for (int i = 0; i < unknownArray.Length; i++)
                        {
                            appendingStr += (Convert.ToString(unknownArray.GetValue(i)) + ',');
                        }

                        row.GetCell(ColumnNumber).SetCellValue(appendingStr);

                    }
                    else         //打印剩下的屬性類型是基礎數據類型的情況
                    {

                        

                        //打印基础类型数据
                        var value_cell =  row.CreateCell(ColumnNumber);
                        value_cell.SetCellValue(Convert.ToString(pro.GetValue(item)));

                        WrapperConverter.PutOnCellStyle(pro,value_cell,Version);
                        

                        ColumnNumber++;

                    }

                }
                RowNumber++;

                //处理行打印后事件

                //处理是否隐藏一行的条件,现只支持行隐藏,不支持列隐藏
                if(typeof(T).GetCustomAttribute(typeof(Hide_On_Condition))!=null)
                {
                    var hide_attr =  (Hide_On_Condition)typeof(T).GetCustomAttribute(typeof(Hide_On_Condition));
                    bool ishide = ExpressionHelp.Read_Condition_expression<T>(hide_attr.rowCondition,item);

                    if(ishide)
                    {
                        row.Hidden = true;
                    }
                }


            });

            if(CustomTailMethod!=null)
            {
                //给出最后一行开始的RowNumber
                CustomTailMethod.Invoke(workbook,RowNumber);
            }

            //最后如果有冻结标签就设置冻结设置的行列

            if(type.GetCustomAttribute(typeof(FreezeAreaAttribute))!=null)
            {
                var freeze = (FreezeAreaAttribute)type.GetCustomAttribute(typeof(FreezeAreaAttribute));

                workbook.GetSheetAt(0).CreateFreezePane(freeze.FreezeStartCol,freeze.FreezeStartRow);
            }

            return workbook;
        }


        /// <summary>
        /// 待完善
        /// </summary>
        /// <param name="list"></param>
        /// <typeparam name="T"></typeparam>
        /// <returns></returns>
        public IWorkbook ConvertToExcel_Double<T>(List<List<T>> list)
        {
            throw new NotImplementedException();
        }

        //判断一个类型是否为基础数据类型
        /// <summary>
        /// 是为true,否为false
        /// </summary>
        /// <param name="type"></param>
        /// <returns></returns>
        public static bool isBasicType(Type type)
        {
            if (type.Equals(typeof(int)) ||
                type.Equals(typeof(double)) ||
                type.Equals(typeof(float)) ||
                type.Equals(typeof(bool)) ||
                type.Equals(typeof(string)) ||
                type.Equals(typeof(byte)) ||
                type.Equals(typeof(char)) ||
                type.Equals(typeof(long)) ||
                type.Equals(typeof(DateTime)) ||
                type.Equals(typeof(decimal))
                )
            {
                return true;
            }
            return false;
        }

        /// <summary>
        /// 判断类型是否为基础数据类型的数组
        /// </summary>
        /// <param name="type"></param>
        /// <returns></returns>
        public static bool isBasicArrayType(Type type)
        {
            if (type.Equals(typeof(int[])) ||
               type.Equals(typeof(double[])) ||
               type.Equals(typeof(float[])) ||
               type.Equals(typeof(bool[])) ||
               type.Equals(typeof(string[])) ||
               type.Equals(typeof(byte[])) ||
               type.Equals(typeof(char[])) ||
               type.Equals(typeof(long[])) ||
               type.Equals(typeof(DateTime[])) ||
               type.Equals(typeof(decimal[]))
               )
            {
                return true;
            }
            return false;
        }


        /// <summary>
        /// 检查Class上的Attribute
        /// </summary>
        /// <param name="type"></param>
        private void Check_Class_Attr(Type type)
        {
            if(type.GetCustomAttribute(typeof(TitleAttribute))!=null)
            {
                has_title_attr = true;
            }
            else
            {
                has_title_attr = false;
            }

            if(type.GetCustomAttribute(typeof(Hide_On_Condition))!=null)
            {
                has_condition_hide = true;
            }
            else
            {
                has_condition_hide = false;
            }

            if(type.GetCustomAttribute(typeof(FreezeAreaAttribute))!=null)
            {
                has_freeze_area = true;
            }
            else
            {
                has_freeze_area = false;
            }

            if(type.GetCustomAttribute(typeof(MergeNearEqualBox))!=null)
            {
                has_merge_near = true;
            }
            else
            {
                has_merge_near = false;
            }
            
        }
    
    }

}
