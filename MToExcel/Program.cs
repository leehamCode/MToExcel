using MToExcel.Converter;
using MToExcel.poco;
using NPOI.SS.UserModel;

namespace MToExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            List<Person> list = new List<Person>()
            {
                new Person { id = "202101",name = "张三",tall = 1.7f,pet = new Animal { Id = 1,Name = "佩奇",Category="猪",LivingArea="新日暮里" } },
                new Person { id = "202102",name = "李四",tall = 1.8f,pet = new Animal { Id = 2,Name = "旺财",Category="狗",LivingArea="SomeWhere" } },
            };

            WrapperConverter wrapper = new WrapperConverter();
            wrapper.basic = new BasicConverter();

            IWorkbook workbook = wrapper.ConvertToExcel<Person>(list);

            FileStream fileStream = new FileStream("C:/Users/F1338705/Desktop/Demo3.xls", FileMode.Create);

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