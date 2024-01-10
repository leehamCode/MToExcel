using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using MToExcel.Models.Enums;

namespace MToExcel.Attributes
{
    /// <summary>
    /// 设置对齐标签
    /// </summary>
    [AttributeUsage(AttributeTargets.Property|AttributeTargets.Class,Inherited =false,AllowMultiple =false)]
    public class HorizonAttribute:Attribute
    {
        public HorizonAttribute(Horizon horizon, VerticalHorizon verticalHorizon)
        {
            this.horizon = horizon;
            this.verticalHorizon = verticalHorizon;
        }

        public Horizon horizon{get;set;}

        public VerticalHorizon verticalHorizon{get;set;}

        
    }
}