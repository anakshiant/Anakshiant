using System;
using Xunit;

namespace Anakshiant.ExcelHelper.Test
{
    public class ExcelColumn_Validation
    {
        private readonly ExcelColumn column;
        public ExcelColumn_Validation()
        {
            column = new ExcelColumn();
        }

        [Fact]
        public void ValidationDataSetAdd()
        {
            column.ValidationDataSet.Add("Testing data set");

            Assert.True(column.ValidationDataSet.Count == 1);
        }

    }
}
