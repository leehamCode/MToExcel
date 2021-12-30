using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MToExcel.poco
{
    public class Animal
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public string Category { get; set; }

        public string LivingArea { get; set; }

        public int[] testArray { get; set; }
    }

}
