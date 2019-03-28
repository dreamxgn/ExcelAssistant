using ExcelAssit;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Test
{
    class Program
    {
        static void Main(string[] args)
        {
            /*List<Object> list = new List<Object>() {
                new Person() { Name = "Tom", Age = 19,Phone="phone.jpg",Remark="这是备注" },
                new Person() { Name = "Tom", Age = 19,Phone="phone.jpg",Remark="这是备注" },
                new Person() { Name = "Tom", Age = 19,Phone="phone.jpg",Remark="这是备注" },
                new Person() { Name = "Tom", Age = 19,Phone=null,Remark="这是备注" },
                new Person() { Name = "Tom", Age = 19,Phone="phone.jpg",Remark="这是备注" },
                new Person() { Name = "Tom", Age = 19,Phone="phone.jpg",Remark="这是备注" },
                new Person() { Name = "Tom", Age = 19,Phone="phone.jpg",Remark="这是备注" }
            };

            bool ret = ExcelAssit.ExcelAssit.AppendExcel("aa.xls", "bbb",list);
            */

            var list=ExcelAssit.ExcelAssit.ReadExcel<Person>("aa.xls", "bbb");

            Console.WriteLine("OK");
            Console.ReadKey();
        }
    }


    class Person
    {
        [AssitCell(CellTitle ="姓名",CellType =CellType.String)]
        public string Name { get; set; }

        [AssitCell(CellTitle = "年龄", CellType = CellType.Int)]
        public int Age { get; set; }

        [AssitCell(CellTitle ="电话",CellType =CellType.Image)]
        public string Phone { get; set; }

        [AssitCell(CellTitle ="备注",CellType=CellType.String)]
        public string Remark { get; set; }
    }
}
