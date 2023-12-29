using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace MToExcel.Converter
{
    public class GuidTypePair
    {
        public GuidTypePair()
        {
            this.guid = Guid.NewGuid();
        }

        public GuidTypePair(Type type){
            this.guid = Guid.NewGuid();
            this.type = type;
        }

        public GuidTypePair(Guid guid, Type type)
        {
            this.guid = guid;
            this.type = type;
        }

        

        /// <summary>
        /// 当作一个唯一标识
        /// </summary>
        /// <value></value>
        public Guid guid{get;set;}

        /// <summary>
        /// 必要类型
        /// </summary>
        /// <value></value>
        public Type type{get;set;}


        
    }
}