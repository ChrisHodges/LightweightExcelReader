using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using BenchmarkDotNet.Attributes;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelDataReader;
using GemBox.Spreadsheet;
using LightWeightExcelReader;
using Net.SourceForge.Koogra;
using NPOI.XSSF.UserModel;

namespace LightweightExcelReader.Benchmarks
{
    public class Benchmarks
    {
        [Benchmark]
        public void LightweightExcelReader()
        {
            var fileName = TestHelper.TestsheetPath("TestSpreadsheet1.xlsx");
            var list = new List<object>();
            var reader = new ExcelReader(fileName);
            var sheet = reader["Sheet5Booleans"];
            while (sheet.ReadNext())
            {
                list.Add(sheet.Value);
            }

            if (list.Count != 1540)
            {
                throw new Exception();
            }
        }

        [Benchmark]
        public void OpenXml()
        {
            var fileName = TestHelper.TestsheetPath("TestSpreadsheet1.xlsx");
            var list = new List<object>();
            using (var fs = new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                using (var doc = SpreadsheetDocument.Open(fs, false))
                {
                    var workbookPart = doc.WorkbookPart;
                    var sstpart = workbookPart.GetPartsOfType<SharedStringTablePart>().First();
                    var sst = sstpart.SharedStringTable;

                    var worksheetPart = workbookPart.WorksheetParts.First();
                    var sheet = worksheetPart.Worksheet;

                    var cells = sheet.Descendants<Cell>();

                    // One way: go through each cell in the sheet
                    foreach (var cell in cells)
                    {
                        if (cell.DataType != null && cell.DataType == CellValues.SharedString)
                        {
                            var ssid = int.Parse(cell.CellValue.Text);
                            var str = sst.ChildElements[ssid].InnerText;
                            list.Add(str);
                        }
                        else if (cell.CellValue != null)
                        {
                            list.Add(cell.CellValue.Text);
                        }
                    }
                }
            }

            if (list.Count != 1540)
            {
                throw new Exception($"{list.Count}");
            }
        }

        [Benchmark]
        public void KoograTest()
        {
            var fileName = TestHelper.TestsheetPath("TestSpreadsheet1.xlsx");
            var reader = WorkbookFactory.GetExcel2007Reader(fileName);
            var worksheet = reader.Worksheets.GetWorksheetByIndex(0);
            var list = new List<object>();
            for (var r = worksheet.FirstRow; r <= worksheet.LastRow; ++r)
            {
                var row = worksheet.Rows.GetRow(r);
                if (row != null)
                {
                    for (var colCount = worksheet.FirstCol; colCount <= worksheet.LastCol; ++colCount)
                    {
                        var cellData = string.Empty;
                        if (row.GetCell(colCount) != null && row.GetCell(colCount).Value != null)
                        {
                            cellData = row.GetCell(colCount).Value.ToString();
                        }

                        list.Add(cellData);
                    }
                }
            }

            if (list.Count != 1540)
            {
                throw new Exception($"{list.Count}");
            }
        }

        [Benchmark]
        public void ExcelDataReader()
        {
            var list = new List<object>();
            var fileName = TestHelper.TestsheetPath("TestSpreadsheet1.xlsx");
            using (var stream = File.Open(fileName, FileMode.Open, FileAccess.Read))
            {
                // Auto-detect format, supports:
                //  - Binary Excel files (2.0-2003 format; *.xls)
                //  - OpenXml Excel files (2007 format; *.xlsx)
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    // Choose one of either 1 or 2:

                    // 1. Use the reader methods
                    do
                    {
                        while (reader.Read())
                        {
                            for (var i = 0; i < reader.FieldCount; i++)
                            {
                                list.Add(reader.GetValue(i));
                            }
                        }
                    } while (reader.NextResult());
                }
            }

            if (list.Count != 2052)
            {
                throw new Exception($"{list.Count}");
            }
        }

        [Benchmark]
        public void OleDbDataAdapter()
        {
            var fileName = TestHelper.TestsheetPath("TestSpreadsheet1.xlsx");
            var connectionString =
                $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={fileName}; Extended Properties=Excel 12.0;";

            var adapter = new OleDbDataAdapter("SELECT * FROM [workSheetNameHere$]", connectionString);
            var ds = new DataSet();

            adapter.Fill(ds, "anyNameHere");

            var data = ds.Tables["anyNameHere"];
        }

        public void GemBox()
        {
            var list = new List<object>();
            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");
            var fileName = TestHelper.TestsheetPath("TestSpreadsheet1.xlsx");
            var workbook = ExcelFile.Load(fileName);
            var worksheet = workbook.Worksheets.First();
            foreach (var cell in worksheet.Cells)
            {
                list.Add(cell.Value);
            }

            if (list.Count != 2052)
            {
                throw new Exception($"{list.Count}");
            }
        }

        //[Benchmark]
        public void NPoi()
        {
            var list = new List<object>();
            var fileName = TestHelper.TestsheetPath("TestSpreadsheet1.xlsx");
            var stream = File.Open(fileName, FileMode.Open, FileAccess.Read);
            var workbook = new XSSFWorkbook(stream);
            var datatypeSheet = workbook.GetSheetAt(0);
            for (var i = 0; i < datatypeSheet.PhysicalNumberOfRows; i++)
            {
                var row = datatypeSheet.GetRow(i);
                foreach (var cell in row.Cells)
                {
                    list.Add(cell.StringCellValue);
                }
            }

            if (list.Count != 2052)
            {
                throw new Exception($"{list.Count}");
            }
        }
    }
}