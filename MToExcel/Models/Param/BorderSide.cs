using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace MToExcel.Models.Param
{
    /// <summary>
    /// 便设置样式的集合
    /// </summary>
    public class BorderSide
    {
        /// <summary>
        /// 各个边需要的样式
        /// </summary>
        /// <value></value>
        public List<BorderSet>? sides {get;set;}

        /// <summary>
        /// 描述
        /// </summary>
        /// <value></value>
        public string? desctipt{get;set;}
        
    }
}