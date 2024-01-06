using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using MToExcel.Exceptons;

namespace MToExcel.Attributes
{
    [AttributeUsage(AttributeTargets.Property,Inherited = false,AllowMultiple =false)]
    public class BackForeColorAttribute:Attribute
    {
        public byte[] back_rgb = null;

        public byte[] fore_rgb = null;

        public BackForeColorAttribute(byte[] back_rgb, byte[] fore_rgb)
        {
            if(back_rgb.Length<3||fore_rgb.Length<3)
            {
                throw new RgbArrayException("长度必须固定为3");
            }
            this.back_rgb = back_rgb;
            this.fore_rgb = fore_rgb;
        }

        public BackForeColorAttribute(bool BakcOrFore,byte[] rgb)
        {
            if(rgb.Length<3)
            {
                throw new RgbArrayException("长度必须固定为3");
            }

            if(BakcOrFore)
            {
                this.back_rgb = rgb;
            }
            else
            {
                this.fore_rgb = rgb;
            }
        }

        
    }
}