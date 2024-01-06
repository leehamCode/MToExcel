using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using MToExcel.Attributes;
using MToExcel.Attributes.TestAttrs;
using MToExcel.Models.Param;
using NPOI.HSSF.Util;



namespace MToExcel.poco
{
    public class TestClass
    {
        public static ddc obj = new ddc("asda");


        [HeaderName("姓名")]
        public string thename{get;set;}

        [HeaderName("年龄")]
        public int age{get;set;}

        [HeaderName("地址")]
        [StructTestAttriubte(new String[]{ "西游记","水浒传"})]
        public string address{get;set;}

        [HeaderName("电话")]
        [Horizon(Models.Enums.Horizon.Center,Models.Enums.VerticalHorizon.Up)]
        [BackForeColor(true,new byte[3]{ 50,187,176})]
        [DynaRowColumnLen(123.45,123.45)]
        //[CellStyle(Models.Enums.Horizon.Center, Models.Enums.VerticalHorizon.Up, false, charSet = new CharSet() { Size = 13.1d })]
        public string phone{get;set;}
    }
}