using MToExcel.Attributes;
using NPOI.HSSF.UserModel;
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
    public class BasicConverter : MTConverter
    {

        /// <summary>
        /// 这个布尔变量控制打印的Excel的版本信息
        /// true =  03版
        /// false = 07版
        /// </summary>
        public bool Version { get; set; }

        public BasicConverter()
        {
            Version = true;
        }

        public IWorkbook ConvertToExcel<T>(List<T> list)
        {

            IWorkbook workbook = null;

            if (Version)
                workbook = new HSSFWorkbook();
            else
                workbook = new XSSFWorkbook();

            ISheet defaultSheet = workbook.CreateSheet("SheetOne");

            //获取传递的泛型类型
            Type type = typeof(T);

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

            //先用属性名打印一行表头

            IRow header = defaultSheet.CreateRow(0);

            int i = 0;
            foreach(PropertyInfo pro in properties)
            {
                //设置一下表头样式,将表头设置为加粗字体
                ICellStyle style = workbook.CreateCellStyle();
                var Font = workbook.CreateFont();
                Font.IsBold = true;
                style.SetFont(Font);

                if (WrapperConverter.TypePool.ContainsKey(pro.PropertyType))  //判断泛型的该属性是否在标记类型池中
                {
                    ReferenceType refer = WrapperConverter.TypePool.GetValueOrDefault(pro.PropertyType);

                    Type typeTemp = pro.PropertyType;
                   

                    if (refer.getIsMultiPart()) //判断是否要将引用类型拆成多列
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
                }

                header.CreateCell(i).SetCellValue(pro.Name);
                header.GetCell(i).CellStyle = style;
                i++;
            }

            int RowNumber = 1;            //控制行号增加的变量
            list.ForEach(item => {

                IRow row = defaultSheet.CreateRow(RowNumber); //创建一行写一行的数据

                PropertyInfo[] properties = item.GetType().GetProperties();

                int ColumnNumber = 0;     //控制列增加的变量
                foreach (PropertyInfo pro in properties)
                {

                    if (WrapperConverter.TypePool.ContainsKey(pro.PropertyType))  //判断泛型的该属性是否在标记类型池中
                    {
                        ReferenceType refer = WrapperConverter.TypePool.GetValueOrDefault(pro.PropertyType);

                        if (refer.getIsMultiPart())
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
                        row.CreateCell(ColumnNumber).SetCellValue(Convert.ToString(pro.GetValue(item)));

                        ColumnNumber++;

                    }

                }
                RowNumber++;

            });

            return workbook;
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


    }

}
