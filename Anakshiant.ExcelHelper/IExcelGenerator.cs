using OfficeOpenXml;
using System;

namespace Anakshiant.ExcelHelper
{
    public interface IExcelGenerator : IDisposable
    {
        string DataSheetName { get; }
        string FileName { get; }
        ExcelPackage NativeExcelPackageObject { get; }
        string UiSheetName { get; }

        void Dispose();
        byte[] GetAsByteArray();
        IExcelDataPopulater GetExcelDataPopulater(ExcelColumns excelColumns);
        void PopulateData(ExcelDataPopulater excelDataPopulater);
    }
}