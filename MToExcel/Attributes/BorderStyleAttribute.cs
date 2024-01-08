using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using MToExcel.Models.Enums;

namespace MToExcel.Attributes
{
    [AttributeUsage(AttributeTargets.Property|AttributeTargets.Class,Inherited =true,AllowMultiple =true)]
    public class BorderStyleAttribute:Attribute
    {

        public BorderStyleAttribute(BorderWid wid, byte[] color, BorderDirect direct)
        {
            Wid = wid;
            Color = color;
            Direct = direct;
        }

        public BorderStyleAttribute(BorderWid wid,BorderDirect direct)
        {
            this.Wid = wid;
            this.Direct = direct;
        }

        public BorderWid Wid{get;set;}

        public byte[] Color{get;set;} = null;

        public BorderDirect Direct{get;set;}


        
    }
}