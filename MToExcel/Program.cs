using System.Reflection;
using MToExcel.Attributes.TestAttrs;
using MToExcel.Converter;
using MToExcel.poco;
using NPOI.HSSF.Util;
using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace MToExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            List<TestClass> listOne = new List<TestClass>(){
                new TestClass(){ thename = "弗里斯兰", age = 800, address = "荷兰低地", phone = "shitU" },
                new TestClass(){ thename = "布列塔尼", age = 1200, address = "布列塔尼", phone = "franc" },
                new TestClass(){ thename = "伊利里亚", age = 2300, address = "亚得里亚", phone = "ita" },
                new TestClass(){ thename = "东色雷斯", age = 2500, address = "黑海", phone = "asdqa" }
            };

            WrapperConverter wrapper = new WrapperConverter();
            
            wrapper.basic = new BasicConverter();
            wrapper.basic.CustomHeadMethod = (workbook)=>{
                Console.WriteLine("测试自定义表头-----！！！");
            };

            IWorkbook workbook = wrapper.ConvertToExcel<TestClass>(listOne);

            ICellStyle cs = workbook.CreateCellStyle();
            

            FileStream file = new FileStream("DEMO.xlsx", FileMode.Create);

            workbook.Write(file);

            file.Close();

            
        }

        public void InitTest()
        {
            List<Person> list = new List<Person>()
            {
                new Person { id = "202101",name = "张三",tall = 1.7f,pet = new Animal { Id = 1,Name = "佩奇",Category="猪",LivingArea="新日暮里" } },
                new Person { id = "202102",name = "李四",tall = 1.8f,pet = new Animal { Id = 2,Name = "旺财",Category="狗",LivingArea="SomeWhere" } },
            };
            
            WrapperConverter wrapper = new WrapperConverter();

            //wrapper.basic = new BasicConverter();
            
            IWorkbook workbook = wrapper.ConvertToExcel<Person>(list);

            FileStream fileStream = new FileStream("./Demo3.xlsx", FileMode.Create);
            
            workbook.Write(fileStream);
            
            fileStream.Close();
        }

        public void TestOne()
        {
            int[] arrayOne = new int[4] { 1, 2, 3, 4 };

            int[] arrayTwo = new int[6] { 7, 9, 11, 15, 22, 76 };

            List<Animal> list = new List<Animal>();
            list.Add(new Animal() { Id = 1, Name = "老虎", Category = "哺乳动物", LivingArea = "东亚", testArray = arrayOne });
            list.Add(new Animal() { Id = 2, Name = "熊猫", Category = "哺乳动物", LivingArea = "四川,陕西,云南", testArray = arrayTwo });
            list.Add(new Animal() { Id = 3, Name = "水牛", Category = "哺乳动物", LivingArea = "整个南方" });

            WrapperConverter wrapper = new WrapperConverter();
            wrapper.basic = new BasicConverter();

            IWorkbook workbook = wrapper.ConvertToExcel<Animal>(list);

            ICellStyle cs = workbook.CreateCellStyle();
            

            FileStream file = new FileStream("C:/Users/F1338705/Desktop/DEMO.xls", FileMode.Create);

            workbook.Write(file);

            file.Close();
        }

        public void TestTwo()
        {
            WrapperConverter wrapper = new WrapperConverter();
            wrapper.basic = new BasicConverter();
            List<string> list = new List<string>() { "阿瑪尼亞克", "奧爾良", "波旁", "阿讓宋" };
            IWorkbook workbook = wrapper.ConvertToExcel<string>(list);

            FileStream stream = new FileStream("C:/Users/F1338705/Desktop/wdnmd.xls", FileMode.Create);

            workbook.Write(stream);

            stream.Close();
        }
    }

}