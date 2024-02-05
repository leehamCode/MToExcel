using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using MToExcel.Attributes;

namespace MToExcel.poco
{
    
    public class CustomTestClass
    {
        [HeaderName("省份名称")]
        public string Name{get;set;}

        [HeaderName("旧时名称")]
        public string OldName{get;set;}

        [HeaderName("何处")]
        public string Address{get;set;}

        [HeaderName("河流")]
        public string River{get;set;}

        [HeaderName("山川")]
        public string Mountain{get;set;}
        
    }
}