using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Text;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace Anakshiant.ExcelHelper
{
    public class ExcelGenerator : IDisposable, IExcelGenerator
    {
        private string fileName;
        private ExcelPackage excelPackage;
        private string uiSheetName;
        private string dataSheetName;


        public string UiSheetName => uiSheetName;
        public string DataSheetName => dataSheetName;
        public string FileName => fileName;
        public ExcelPackage NativeExcelPackageObject => excelPackage;


        public ExcelGenerator(string fileName, string uiSheetName, string dataSheetName)
        {
            this.fileName = fileName;
            this.uiSheetName = uiSheetName;
            this.dataSheetName = dataSheetName;
            excelPackage = new ExcelPackage(new FileInfo($"./{FileName}"));
        }

        public IExcelDataPopulater GetExcelDataPopulater(ExcelColumns excelColumns)
        {
            return new ExcelDataPopulater(excelPackage, uiSheetName, dataSheetName)
            {
                ExcelColumns = excelColumns
            };
        }

        public void PopulateData(ExcelDataPopulater excelDataPopulater)
        {
            excelDataPopulater.PopulateData();
        }

        public void Dispose()
        {
            excelPackage.Dispose();
        }

        public byte[] GetAsByteArray() => excelPackage.GetAsByteArray();

        ~ExcelGenerator()
        {
            Dispose();
        }
    }

}
