using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Text;
using System.Linq;
using OfficeOpenXml.Style;
using System.Drawing;
using OfficeOpenXml.DataValidation;

namespace Anakshiant.ExcelHelper
{
    public class ExcelDataPopulater : IExcelDataPopulater
    {
        private ExcelWorksheet uiWorkSheet;
        private ExcelWorksheet dataWorkSheet;
        private ExcelPackage excelPackage;
        private string uiSheetName;
        private string dataSheetName;

        public string UiSheetName => uiSheetName;
        public string DataSheetName => dataSheetName;
        public ExcelColumns ExcelColumns { get; set; }


        public ExcelDataPopulater(ExcelPackage excelPackage, string uiSheetName, string dataSheetName)
        {
            this.uiSheetName = uiSheetName;
            this.dataSheetName = dataSheetName;

            this.excelPackage = excelPackage;

            this.dataWorkSheet = this.excelPackage.Workbook.Worksheets.Add(dataSheetName);
            this.uiWorkSheet = this.excelPackage.Workbook.Worksheets.Add(uiSheetName);

            this.dataWorkSheet.Hidden = eWorkSheetHidden.VeryHidden;
        }

        public void PopulateData()
        {
            if (ExcelColumns != null && ExcelColumns.Count() > 0)
            {
                List<ExcelColumn> columns = ExcelColumns.GetColumns();
                columns.ForEach(column =>
                {
                    ExcelRange validationCellReference = null;
                    if (column.ValidationDataSet != null && column.ValidationDataSet.Count > 0)
                    {
                        validationCellReference = AddDataCell(dataWorkSheet, column.ValidationDataSet);
                    }

                    AddExcelData(uiWorkSheet, column, validationCellReference);
                });
            }
        }


        private ExcelRange AddDataCell(ExcelWorksheet dataSheet, List<string> data)
        {
            int lastColumnIndex = dataSheet.Dimension != null
                ? dataSheet.Dimension.End.Column + 1
                : 1;
            int fromRow = 1;
            int lastRow = fromRow;

            data.ForEach(d =>
            {
                dataSheet.Cells[lastRow, lastColumnIndex].Value = d;
                lastRow++;
            });

            return dataSheet.Cells[fromRow, lastColumnIndex, lastRow, lastColumnIndex];
        }

        private void AddExcelData(ExcelWorksheet uiSheet, ExcelColumn column, ExcelRange validationCellReference = null)
        {
            int columnNumber = uiSheet.Dimension != null
                ? uiSheet.Dimension.End.Column + 1
                : 1;
            const int ROW_NUMBER = 1;

            var uiCellRangeReference = uiSheet.Cells[ROW_NUMBER, columnNumber];

            uiCellRangeReference.Value = column.ColumnName;
            uiCellRangeReference.Style.Fill.PatternType = ExcelFillStyle.Solid;
            uiCellRangeReference.Style.Fill.BackgroundColor.SetColor(Color.LightGray);
            uiCellRangeReference.Style.Locked = true;
            uiCellRangeReference.Style.WrapText = true;

            if (validationCellReference != null)
            {
                uiSheet.Names.Add(column.ColumnName, validationCellReference);
                var validation = uiSheet.DataValidations.AddListValidation(uiCellRangeReference.Address);
                validation.Formula.ExcelFormula = $"={column.ColumnName}";
                validation.ErrorStyle = ExcelDataValidationWarningStyle.stop;
                validation.ShowErrorMessage = true;
                validation.Error = column.ErrorMessage;
            }

        }
    }
}

