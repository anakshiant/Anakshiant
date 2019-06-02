using System;
using System.Collections.Generic;
using System.Text;
using Xunit;

namespace Anakshiant.ExcelHelper.Test
{
    public class ExcelGenerator_Test : IDisposable
    {
        private readonly IExcelGenerator excelGenerator;
        public ExcelGenerator_Test()
        {
            excelGenerator = new ExcelGenerator("anand.xlsx", "uiSheet", "dataSheet");
        }
      
        [Fact]
        public void ShouldGetExcelDataPopulater()
        {
            ExcelColumns columns = new ExcelColumns();
            ExcelColumn column = new ExcelColumn { ColumnName = "Test", ErrorMessage = "not" };
            columns.Add(column);

            IExcelDataPopulater excelDataPopulater = excelGenerator.GetExcelDataPopulater(columns);

            Assert.NotNull(excelDataPopulater);
        }

        [Fact]
        public void ShouldGetByteArray()
        {
            ExcelColumns columns = new ExcelColumns();
            ExcelColumn column = new ExcelColumn { ColumnName = "Test", ErrorMessage = "not" };
            columns.Add(column);

            IExcelDataPopulater excelDataPopulater = excelGenerator.GetExcelDataPopulater(columns);

            byte[] byteArray = excelGenerator.GetAsByteArray();
            Assert.NotNull(byteArray);
        }

        public void Dispose()
        {
            excelGenerator.Dispose();
        }

        ~ExcelGenerator_Test()
        {
            Dispose();
        }
    }
}
