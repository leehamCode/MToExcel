using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using MToExcel.Attributes;

namespace MToExcel.poco
{
    [TitleAttribute(Context ="测试打印标题行$1",Font_Name ="新細明體",Font_Size =16,Font_color =new byte[]{ 80,235,227 })]
    [Hide_On_Condition(rowCondition ="$1==\"九江\"")]
    public class TestClass2
    {
        [HeaderName("你的姓名:")]
        public string Name { get; set; }

        [HeaderName("地址:")]
        [BackForeColor(true,new byte[] {94,89,244})]
        [BorderStyle(MToExcel.Models.Enums.BorderWid.ThinBorder,new byte[] { 252,28,3 },MToExcel.Models.Enums.BorderDirect.Upper)]
        [DynaRowColumnLen(false,400.00)]
        public string Address { get; set; }

        [HeaderName("手机号")]
        [FontSets("Brush Script MT",14,true,true,true,true,new byte[] { 119,251,232 })]
        public string Phone { get; set; }

        [HeaderName("生日")]
        [Horizon(MToExcel.Models.Enums.Horizon.Left,MToExcel.Models.Enums.VerticalHorizon.Up)]
        public string Birth { get; set; }

        [HeaderName("邮箱")]
        public string Email { get; set; }
    }
}