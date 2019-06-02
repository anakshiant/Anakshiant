using System;
using System.Collections.Generic;
using System.Text;
using Xunit;

namespace Anakshiant.ExcelHelper.Test
{
    public class ExcelDataPopulater_Test
    {
        private readonly IExcelGenerator excelGenerator;

        public ExcelDataPopulater_Test()
        {
            excelGenerator = new ExcelGenerator("test.xlsx", "uiSheet", "dataSheet");
        }

        [Fact]
        public void ShoudPopulateData()
        {
            ExcelColumns columns = new ExcelColumns();
            ExcelColumn column = new ExcelColumn { ColumnName = "Test", ErrorMessage = "not" };
            columns.Add(column);

            IExcelDataPopulater excelDataPopulater = new ExcelDataPopulater(excelGenerator.NativeExcelPackageObject, excelGenerator.UiSheetName, excelGenerator.DataSheetName)
            {
                ExcelColumns = columns
            };

            excelDataPopulater.PopulateData();

            byte[] data = excelGenerator.GetAsByteArray();

            Assert.NotNull(data);
        }
    }
}
