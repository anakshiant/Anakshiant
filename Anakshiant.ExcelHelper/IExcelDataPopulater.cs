namespace Anakshiant.ExcelHelper
{
    public interface IExcelDataPopulater
    {
        string DataSheetName { get; }
        ExcelColumns ExcelColumns { get; set; }
        string UiSheetName { get; }

        void PopulateData();
    }
}