using System;
using System.Collections.Generic;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;


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

        private static Excel.Application _excel_app;
        private static Excel.Workbook _excel_wb;
        private static Excel.Worksheet _excel_ws;

        private static int _row = 1;
        private static int _col = 1;

        static void Main(string[] args)
        {
            InitializeData();
            InitializeExcel();
            GetDataTypeAndProperties();
            GenerateReport();
            SaveExcel();
        }

        private static void InitializeData()
        {
            _datalist.Add(new Data { Name = "Scot", Gender = "Male", Age = 30 });
            _datalist.Add(new Data { Name = "Ethan", Gender = "Male", Age = 30 });
            _datalist.Add(new Data { Name = "Eric", Gender = "Female", Age = 50 });
        }

        private static void InitializeExcel()
        {
            _excel_app = new Excel.Application();
            _excel_wb = _excel_app.Workbooks.Add();
            _excel_ws = _excel_wb.Worksheets[1];
            _excel_ws.Name = "001";
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
                _excel_app.Cells[_row, _col] = ((CsvExport)attribute).Type;
       
                foreach (Data d in _datalist)
                {
                    _row++;
                    _excel_app.Cells[_row, _col] = property.GetValue(d);
                }

                _col++;
                _row = 1;
            }
        }

        private static void SaveExcel()
        {
            _excel_wb.SaveAs(Environment.CurrentDirectory + @"\result.xlsx");
            _excel_wb.Close();
            _excel_app.Quit();
        }
    }
}
