using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication1
{
    class Program
    {
        public class Data
        {
            [CsvExport("Name")]
            public string Name { get; set; }

            [CsvExport("Gender")]
            public string Gender { get; set; }

            [CsvExport("SecretNumber")]
            public int Age { get; set; }
        }

        public class CsvExport : Attribute
        {
            public readonly string Type;

            public CsvExport(string type)
            {
                Type = type;
            }
        }

        private static List<Data> _datalist = new List<Data>();
        private static Type _dataType;
        private static PropertyInfo[] _dataProperties;

        static void Main(string[] args)
        {
            InitializeData();
            GetDataTypeAndProperties();
            GenerateReport();
        }

        private static void InitializeData()
        {
            _datalist.Add(new Data { Name = "Scot", Gender = "Male", Age = 30 });
            _datalist.Add(new Data { Name = "Ethan", Gender = "Male", Age = 30 });
            _datalist.Add(new Data { Name = "Eric", Gender = "Female", Age = 50 });
        }

        private static void GetDataTypeAndProperties()
        {
            _dataType = typeof(Data);
            _dataProperties = _dataType.GetProperties();
        }

        private static void GenerateReport()
        {
            foreach (PropertyInfo property in _dataProperties)
                GetDataValueByAttribute(property);
        }

        private static void GetDataValueByAttribute(PropertyInfo property)
        {
            if (Attribute.IsDefined(property, typeof(CsvExport)))
            {
                var attribute = property.GetCustomAttribute(typeof(CsvExport));
                Console.WriteLine(((CsvExport)attribute).Type);

                foreach (Data d in _datalist)
                {
                    Console.WriteLine(property.GetValue(d));
                    Console.ReadLine();
                }
            }
        }



        
    }
}
