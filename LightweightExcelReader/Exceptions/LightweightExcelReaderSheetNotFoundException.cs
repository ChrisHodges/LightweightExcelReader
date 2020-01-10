using System;

namespace LightWeightExcelReader.Exceptions
{
    public class LightweightExcelReaderSheetNotFoundException : Exception
    {
        public LightweightExcelReaderSheetNotFoundException(string sheetName) : base(
            $"Sheet with name '{sheetName}' was not found in the workbook")
        {
        }

        public LightweightExcelReaderSheetNotFoundException(int sheetNumber, int numberOfSheets):base($"Sheet with zero-based index {sheetNumber} not found in the workbook. Workbook contains {numberOfSheets} sheets")
        {
        }
    }
}