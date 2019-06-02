using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;

namespace Anakshiant.ExcelHelper
{
    public class ExcelColumns : IEnumerable
    {
        private List<ExcelColumn> Columns;
        public ExcelColumns()
        {
            Columns = new List<ExcelColumn>();
        }
        public ExcelColumns(List<ExcelColumn> columns)
        {
            Columns = columns;
        }
        public void Add(ExcelColumn column) => Columns.Add(column);
        public void Remove(ExcelColumn column) => Columns.Remove(column);
        public int Count() => Columns.Count;
        public List<ExcelColumn> GetColumns() => Columns;
        public IEnumerator GetEnumerator() => Columns.GetEnumerator();
    }
}
