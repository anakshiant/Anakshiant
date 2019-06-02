using System;
using System.Collections.Generic;
using System.Text;
using Xunit;

namespace Anakshiant.ExcelHelper.Test
{
    public class ExcelColumns_Test
    {
        
        [Fact]
        public void ShouldAddOneColumn()
        {
            ExcelColumns columns = new ExcelColumns();
            ExcelColumn column = new ExcelColumn { ColumnName = "Test", ErrorMessage = "not" };
            columns.Add(column);
            Assert.True(columns.Count() == 1);
        }

        [Fact]
        public void ShouldRemoveColumn()
        {
            ExcelColumns columns = new ExcelColumns();
            ExcelColumn column = new ExcelColumn { ColumnName = "Test", ErrorMessage = "not" };
            columns.Add(column);
            columns.Remove(column);
            Assert.True(columns.Count() == 0);
        }

        [Fact]
        public void ShoudGetAllColumns()
        {
            ExcelColumns columns = new ExcelColumns();
            ExcelColumn column = new ExcelColumn { ColumnName = "Test", ErrorMessage = "not" };
            columns.Add(column);
            Assert.IsType<List<ExcelColumn>>(columns.GetColumns());
        }
    }
}
