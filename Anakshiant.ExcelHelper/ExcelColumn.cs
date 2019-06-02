using System;
using System.Collections.Generic;
using System.Text;

namespace Anakshiant.ExcelHelper
{

    public class ExcelColumn 
    {
        public string ColumnName { get; set; }
        public string ErrorMessage { get; set; } = "Select/Input a value from the dropdown";

        public List<string> ValidationDataSet { get; set; }

        public ExcelColumn()
        {
            ValidationDataSet = new List<string>();
        }
    }
}
