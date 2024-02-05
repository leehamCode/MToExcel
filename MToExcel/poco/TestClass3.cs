using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using MToExcel.Attributes;

namespace MToExcel.poco
{
    /// <summary>
    /// 日期格式转化测试Model
    /// </summary>
    public class TestClass3
    {
        [HeaderName("学校名称")]
        public string Name{get;set;}

        [HeaderName("学校地址")]
        public string Region{get;set;}

        [HeaderName("创办日期")]
        [DateTimeFormat(format ="yyyy-MM-dd")]
        [FontSets("標楷體",16,true,false,false,false,new byte[]{233,236,0})]
        public DateTime Create_date{get;set;}

        [HeaderName("排名")]
        public int Rank{get;set;}

        [HeaderName("校领导")]
        public string head_teacher{get;set;}

        [HeaderName("备注")]
        public string remark{get;set;}
    }
}