using System.Collections.Generic;
using System.IO;
using Anakshiant.ExcelHelper;

namespace Anakshiant.Main
{
    class Program
    {
        static void Main(string[] args)
        {
            List<ExcelColumn> columns = new List<ExcelColumn>()
            {
                new ExcelColumn{ColumnName = "Name"},
                new ExcelColumn{ColumnName = "Customer Reference Number"},
                new ExcelColumn{ColumnName = "Gender",ValidationDataSet = new List<string>(){"Male","Female"}}
            };

            using (IExcelGenerator generator = new ExcelGenerator("anand.xlsx", "serviceSheet", "data"))
            {
                ExcelColumns excelColumns = new ExcelColumns(columns);

                var dataPopulater = generator.GetExcelDataPopulater(excelColumns);

                dataPopulater.PopulateData();

                generator.NativeExcelPackageObject.SaveAs(new FileInfo("./students.xlsx"));
            }
        }
    }






}
