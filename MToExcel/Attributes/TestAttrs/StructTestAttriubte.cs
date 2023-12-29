using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace MToExcel.Attributes.TestAttrs
{
    [AttributeUsage(AttributeTargets.Class | AttributeTargets.Property,AllowMultiple = true)]
    public class StructTestAttriubte : Attribute
    {
        public  ddc s1{get;set;}

        public string[] args{get;set;}

        public StructTestAttriubte(ddc s1)
        {
            this.s1 = s1;
        }

        public StructTestAttriubte(string[] args)
        {
            this.args = args;
        }
    }


    public struct ddc{
        public string shit{get;set;}

        public ddc(string str)
        {
            this.shit = str;
        }
    }
}