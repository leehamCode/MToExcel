﻿using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace MToExcel.Converter
{
    public interface MTConverter
    {
        /// <summary>
        /// 基础的转化方法
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="list"></param>
        /// <returns></returns>
        public IWorkbook ConvertToExcel<T>(List<T> list);


        /// <summary>
        /// 如果需要做阶段性的总结行应该使用此方法
        /// </summary>
        /// <param name="list"></param>
        /// <typeparam name="T"></typeparam>
        /// <returns></returns>
        public IWorkbook ConvertToExcel_Double<T>(List<List<T>> list);

        
    }

}
