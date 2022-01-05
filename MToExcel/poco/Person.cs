using MToExcel.Attributes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MToExcel.poco
{
    public class Person
    {
        [HeaderName("学生学号")]
        public string id { get; set; }

        public string name { get; set; }

        [IgnoreType]
        public float tall { get; set; }

        [ReferenceType(true)]
        public Animal pet { get; set; }
    }

}
