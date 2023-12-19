using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Policy;
using System.Threading.Tasks;
using MToExcel.Models.Enums;

namespace MToExcel.Models.Param
{
    /// <summary>
    /// 边框参数对象
    /// </summary>
    public class BorderSet
    {
        public BorderWid Wid{get;set;}

        public short Color{get;set;}

        public BorderDirect Direct{get;set;}

        // override object.Equals
        public override bool Equals(object obj)
        {
            var target = (BorderSet)obj;

            if(Wid==target.Wid&&Color==target.Color&&Direct==target.Direct)
            {
                return true;
            }
            else{
                return false;
            }

    
        }
        
        // override object.GetHashCode
        public override int GetHashCode()
        {
            return HashCode.Combine(Wid,Color,Direct);
        }
    }
}