using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using FluentAssertions;
using Xunit;

namespace LightweightExcelReader.Tests
{
    public class LightweightExcelReaderTests
    {
        [Fact]
        public void BySheetIndexWorks()
        {
            var testFileLocation = TestHelper.TestsheetPath("TestSpreadsheet1.xlsx");
            var lightWeightExcelReader = new ExcelReader(testFileLocation);
            lightWeightExcelReader[1]["A2"].Should().Be("abc123sheet2");
            lightWeightExcelReader[1]["D4"].Should().Be(-5);
            lightWeightExcelReader[1]["C3"].Should().Be(new DateTime(2098, 10, 9));
        }

        [Fact]
        public void DifferentSheetWorks()
        {
            var testFileLocation = TestHelper.TestsheetPath("TestSpreadsheet1.xlsx");
            var lightWeightExcelReader = new ExcelReader(testFileLocation);
            lightWeightExcelReader["Sheet4DuplicateStringValue"]["A2"].Should().Be("abc123");
            lightWeightExcelReader["Sheet4DuplicateStringValue"]["D4"].Should().Be(9.876);
            lightWeightExcelReader["Sheet4DuplicateStringValue"]["C3"].Should().Be(new DateTime(2015, 10, 9));
        }

        [Fact]
        public void IsoStandardDateWorks()
        {
            var testFileLocation = TestHelper.TestsheetPath("ISOStandardDateTest.xlsx");
            var lightWeightExcelReader = new ExcelReader(testFileLocation);
            lightWeightExcelReader["Data"]["B2"].Should().Be(new DateTime(2013,3,3));
        }
        
        [Fact]
        public void NonExistingSheetThrowsMeaningfulError()
        {
            var testFileLocation = TestHelper.TestsheetPath("TestSpreadsheet1.xlsx");
            var lightWeightExcelReader = new ExcelReader(testFileLocation);
            ArgumentOutOfRangeException exception = null;
            try
            {
                var sheet = lightWeightExcelReader["ThisSheetDoesNotExist"];
                var test = sheet["A1"];
            }
            catch (Exception e)
            {
                exception = e as ArgumentOutOfRangeException;
            }

            exception.Should().NotBe(null);
            exception.Message.Should().StartWith("Sheet with name 'ThisSheetDoesNotExist' was not found in the workbook");
        }
        
        [Fact]
        public void NonExistingSheetIndexThrowsMeaningfulError()
        {
            var testFileLocation = TestHelper.TestsheetPath("TestSpreadsheet1.xlsx");
            var lightWeightExcelReader = new ExcelReader(testFileLocation);
            ArgumentOutOfRangeException exception = null;
            try
            {
                var sheet = lightWeightExcelReader[999];
                var test = sheet["A1"];
            }
            catch (Exception e)
            {
                exception = e as ArgumentOutOfRangeException;
            }

            exception.Should().NotBe(null);
            exception.Message.Should().StartWith("Sheet with zero-based index 999 was not found in the workbook. Workbook contains 10 sheets");
        }

        [Fact]
        public void MoreThanNineSheetsWorks()
        {
            var testFileLocation = TestHelper.TestsheetPath("TestSpreadsheet1.xlsx");
            var lightWeightExcelReader = new ExcelReader(testFileLocation);
            var sheet10 = lightWeightExcelReader["Sheet10"];
            sheet10["A2"].Should().Be("SHEET10Values");
        }

        [Fact]
        public void FromStreamWorks()
        {
            var testFileLocation = TestHelper.TestsheetPath("TestSpreadsheet1.xlsx");
            var filestream = new FileStream(testFileLocation, FileMode.Open, FileAccess.Read);
            var lightWeightExcelReader = new ExcelReader(filestream);
            lightWeightExcelReader["Sheet4DuplicateStringValue"]["A2"].Should().Be("abc123");
            lightWeightExcelReader["Sheet4DuplicateStringValue"]["D4"].Should().Be(9.876);
            lightWeightExcelReader["Sheet4DuplicateStringValue"]["C3"].Should().Be(new DateTime(2015, 10, 9));
        }

        [Fact]
        public void BooleansWithFormatCodeBWork()
        {
            var testFileLocation = TestHelper.TestsheetPath("TypeBBooleans.xlsx");
            var lightWeightExcelReader = new ExcelReader(testFileLocation);
            lightWeightExcelReader["sheet"]["A1"].Should().Be(true);
        }

        [Fact]
        public void BooleansWork()
        {
            var testFileLocation = TestHelper.TestsheetPath("Booleans.xlsx");
            var lightWeightExcelReader = new ExcelReader(testFileLocation);
            lightWeightExcelReader["Sheet1"]["A1"].Should().Be(0);
            lightWeightExcelReader["Sheet1"]["A2"].Should().Be(1);
            lightWeightExcelReader["Sheet1"]["B1"].Should().Be(false);
            lightWeightExcelReader["Sheet1"]["B2"].Should().Be(true);
        }

        [Fact]
        public void GermanDecimalsWork()
        {
            Thread.CurrentThread.CurrentCulture = new CultureInfo("de-DE"); 
            var testFileLocation = TestHelper.TestsheetPath("GermanDecimals.xlsx");
            var lightWeightExcelReader = new ExcelReader(testFileLocation);
            var sheet = lightWeightExcelReader["Sheet1"];
            sheet["A2"].Should().Be(1.234);
            sheet["B2"].Should().Be(1.234);
            sheet["C2"].Should().Be("1,234");
            sheet["A3"].Should().Be(9.876);
            sheet["B3"].Should().Be(9.876);
            sheet["C3"].Should().Be("9,876");
        }
        
        [Fact]
        public void GermanDecimalsWorkAlternativeSheetFormat()
        {
            Thread.CurrentThread.CurrentCulture = new CultureInfo("de-DE"); 
            var testFileLocation = TestHelper.TestsheetPath("doubletest.xlsx");
            var lightWeightExcelReader = new ExcelReader(testFileLocation);
            var sheet = lightWeightExcelReader["Tabelle1"];
            sheet["G3"].Should().Be(289.99);
        }

        [Fact]
        public void GetOfficePrefixedSheetWorks()
        {
            var testFileLocation = TestHelper.TestsheetPath("OfficePrefixedSheetText.xlsx");
            var lightWeightExcelReader = new ExcelReader(testFileLocation);
            lightWeightExcelReader["QuoteSheetId"]["A1"].Should().Be(32231);
        }

        [Fact]
        public void GetPrefixedSheetWorks()
        {
            var testFileLocation = TestHelper.TestsheetPath("PrefixedSheetTest.xlsx");
            var lightWeightExcelReader = new ExcelReader(testFileLocation);
            lightWeightExcelReader["QuoteSheetId"]["A1"].Should().Be(27706);
        }

        [Fact]
        public void NoRangeTest()
        {
            var testFileLocation = TestHelper.TestsheetPath("TestCurrencySpreadsheet.xlsx");
            var lightWeightExcelReader = new ExcelReader(testFileLocation);
            lightWeightExcelReader["BlankSheet"]["A1"].Should().Be(null);
            lightWeightExcelReader["BlankSheet"].ContainsKey("A1").Should().BeFalse();
        }

        [Fact]
        public void ContainsKeyWorks()
        {
            var testFileLocation = TestHelper.TestsheetPath("TestSpreadsheet1.xlsx");
            var lightWeightExcelReader = new ExcelReader(testFileLocation);
            var sheet1 = lightWeightExcelReader["Sheet1"];
            sheet1.ContainsKey("A2").Should().Be(true);
            sheet1.ContainsKey("AA20000").Should().Be(false);
        }

        [Fact] public void GetFirstDateTimeStyleWorks()
        {
            var testFileLocation = TestHelper.TestsheetPath("TestSpreadsheet1.xlsx");
            var lightWeightExcelReader = new ExcelReader(testFileLocation);
            lightWeightExcelReader.GetFirstDateTimeStyle().Should().Be(2);
            var sheet1 = lightWeightExcelReader["Sheet1"];
            sheet1["A2"].Should().Be("abc123");
            sheet1["D4"].Should().Be(5);
            sheet1["C3"].Should().Be(new DateTime(2015, 10, 9));
            sheet1["C4"].Should().Be(null);
        }

        [Fact]
        public void RangeIndexWorks()
        {
            var testFileLocation = TestHelper.TestsheetPath("TestSpreadsheet1.xlsx");
            var lightWeightExcelReader = new ExcelReader(testFileLocation);
            var sheet1 = lightWeightExcelReader["Sheet1"];
            sheet1["A2"].Should().Be("abc123");
            sheet1["D4"].Should().Be(5);
            sheet1["C3"].Should().Be(new DateTime(2015, 10, 9));
            sheet1["C4"].Should().Be(null);
            Action tryGetOutOfRange = () =>
            {
                var test = sheet1["C5"];
            };
            tryGetOutOfRange.Should().Throw<IndexOutOfRangeException>();
        }

        [Fact]
        public void LargeDoubleReadsCorrectly()
        {
            var testFileLocation = TestHelper.TestsheetPath("LargeDouble.xlsx");
            var lightWeightExcelReader = new ExcelReader(testFileLocation);
            var sheet1 = lightWeightExcelReader[0];
            sheet1["A1"].Should().Be(7060151014090010);
        }

        [Fact]
        public void RangeWorks()
        {
            var testFileLocation = TestHelper.TestsheetPath("TestSpreadsheet1.xlsx");
            var lightWeightExcelReader = new ExcelReader(testFileLocation);
            var range = lightWeightExcelReader["Sheet1"]["A1", "D4"].ToList();
            range[0].Should().Be("String");
            range[1].Should().Be("Int");
            range[2].Should().Be("DateTime");
            range[3].Should().Be("Decimal");
            range[4].Should().Be("abc123");
            range[5].Should().Be(1);
            range[6].Should().Be(new DateTime(2012, 12, 31));
            range[7].Should().Be(1.234);
            range[8].Should().Be("zyx987");
            range[9].Should().Be(2);
            range[10].Should().Be(new DateTime(2015, 10, 9));
            range[11].Should().Be(9.876);
            range[12].Should().Be(null);
            range[13].Should().Be(null);
            range[14].Should().Be(null);
            range[15].Should().Be(5);
        }

        [Fact]
        public void ReadNextInRowWorks()
        {
            var testFileLocation = TestHelper.TestsheetPath("TestSpreadsheet1.xlsx");
            var lightWeightExcelReader = new ExcelReader(testFileLocation);
            var sheet1 = lightWeightExcelReader["Sheet5Booleans"];
            sheet1.ReadNextInRow();
            sheet1.Value.Should().Be("String");
            sheet1.Address.Should().Be("A1");
            sheet1.ReadNextInRow();
            sheet1.Value.Should().Be("Int");
            sheet1.Address.Should().Be("B1");
            sheet1.ReadNextInRow();
            sheet1.Value.Should().Be("Bool");
            sheet1.Address.Should().Be("C1");
            sheet1.ReadNextInRow();
            sheet1.Value.Should().Be("NullableBool");
            sheet1.Address.Should().Be("D1");
            sheet1.ReadNextInRow();
            sheet1.Value.Should().Be("DateTime");
            sheet1.Address.Should().Be("E1");
            sheet1.ReadNextInRow();
            sheet1.Value.Should().Be(null);
            sheet1.Address.Should().Be(null);
            sheet1.ReadNextInRow();
            sheet1.Value.Should().Be(null);
            sheet1.Address.Should().Be(null);
        }

        [Fact]
        public void ReadNextWorks()
        {
            var testFileLocation = TestHelper.TestsheetPath("TestSpreadsheet1.xlsx");
            var lightWeightExcelReader = new ExcelReader(testFileLocation);
            var sheet1 = lightWeightExcelReader["Sheet5Booleans"];
            sheet1.CurrentRowNumber.Should().BeNull();
            sheet1.PreviousRowNumber.Should().BeNull();
            sheet1.ReadNext();
            sheet1.Value.Should().Be("String");
            sheet1.Address.Should().Be("A1");
            sheet1.CurrentRowNumber.Should().Be(1);
            sheet1.PreviousRowNumber.Should().BeNull();
            sheet1.ReadNext();
            sheet1.Value.Should().Be("Int");
            sheet1.Address.Should().Be("B1");
            sheet1.CurrentRowNumber.Should().Be(1);
            sheet1.PreviousRowNumber.Should().BeNull();
            sheet1.ReadNext();
            sheet1.Value.Should().Be("Bool");
            sheet1.Address.Should().Be("C1");
            sheet1.ReadNext();
            sheet1.Value.Should().Be("NullableBool");
            sheet1.Address.Should().Be("D1");
            sheet1.ReadNext();
            sheet1.Value.Should().Be("DateTime");
            sheet1.Address.Should().Be("E1");
            sheet1.ReadNext();
            sheet1.CurrentRowNumber.Should().Be(2);
            sheet1.PreviousRowNumber.Should().Be(1);
            sheet1.Value.Should().Be("abc123");
            sheet1.Address.Should().Be("A2");
            sheet1.ReadNext();
            sheet1.Value.Should().Be(1);
            sheet1.Address.Should().Be("B2");
            sheet1.ReadNext();
            sheet1.Value.Should().Be(new DateTime(1990, 11, 29));
            sheet1.Address.Should().Be("E2");
            sheet1.ReadNext();
            sheet1.Value.Should().Be("abc123");
            sheet1.Address.Should().Be("A3");
            sheet1.ReadNext();
            sheet1.Value.Should().Be(1);
            sheet1.Address.Should().Be("B3");
            sheet1.ReadNext();
            sheet1.Value.Should().Be("Yes");
            sheet1.Address.Should().Be("C3");
            sheet1.ReadNext();
            sheet1.Value.Should().Be("Yes");
            sheet1.Address.Should().Be("D3");
            sheet1.ReadNext();
            sheet1.Value.Should().Be("zxy123");
            sheet1.Address.Should().Be("A4");
            sheet1.ReadNext();
            sheet1.Value.Should().Be(2);
            sheet1.Address.Should().Be("B4");
            sheet1.ReadNext();
            sheet1.Value.Should().Be("No");
            sheet1.Address.Should().Be("C4");
            sheet1.ReadNext();
            sheet1.Value.Should().Be("No");
            sheet1.Address.Should().Be("D4");
            sheet1.ReadNext();
            sheet1.Value.Should().Be(null);
            sheet1.Address.Should().Be(null);
        }

        [Fact]
        public void ServerGeneratedFileWorks()
        {
            var testFileLocation = TestHelper.TestsheetPath("Import.xlsx");
            var lightWeightExcelReader = new ExcelReader(testFileLocation);
            lightWeightExcelReader["Keys"]["M3"].Should().Be("en");
        }

        [Fact]
        public void SingleColumnRangeWorks()
        {
            var testFileLocation = TestHelper.TestsheetPath("TestSpreadsheet1.xlsx");
            var lightWeightExcelReader = new ExcelReader(testFileLocation);
            var sheet1 = lightWeightExcelReader["Sheet1"];
            var range = sheet1["A1", "A4"].ToList();
            sheet1.CurrentRowNumber.Should().Be(4);
            sheet1.PreviousRowNumber.Should().Be(3);
            range[0].Should().Be("String");
            range[1].Should().Be("abc123");
            range[2].Should().Be("zyx987");
            range[3].Should().Be(null);
        }

        [Fact]
        public void SingletonRangeTest()
        {
            var testFileLocation = TestHelper.TestsheetPath("TestCurrencySpreadsheet.xlsx");
            var lightWeightExcelReader = new ExcelReader(testFileLocation);
            lightWeightExcelReader["QuoteSheetId"]["A1"].Should().Be(22869);
        }

        [Fact]
        public void TwoSheetsWorks()
        {
            var testFileLocation = TestHelper.TestsheetPath("TestCurrencySpreadsheet.xlsx");
            var lightWeightExcelReader = new ExcelReader(testFileLocation);
            lightWeightExcelReader["Prices"]["E2"].Should().Be("SupplierDescription");
            lightWeightExcelReader["Prices"]["F2"].Should().Be("*LANGUAGE*_du");
            lightWeightExcelReader["Keys"]["M3"].Should().Be("en");
            lightWeightExcelReader["QuoteSheetId"]["A1"].Should().Be(22869);
        }

        [Fact]
        public void WholeColumnWorks()
        {
            var testFileLocation = TestHelper.TestsheetPath("TestSpreadsheet1.xlsx");
            var lightWeightExcelReader = new ExcelReader(testFileLocation);
            var sheet1 = lightWeightExcelReader["Sheet1"];
            var result = sheet1.Column("A").ToArray();
            result[0].Should().Be("String");
            result[1].Should().Be("abc123");
            result[2].Should().Be("zyx987");
            sheet1.CurrentRowNumber.Should().Be(4);
            sheet1.PreviousRowNumber.Should().Be(3);
            result[3].Should().Be(null);
        }

        [Fact]
        public void WholeRowWorks()
        {
            var testFileLocation = TestHelper.TestsheetPath("TestSpreadsheet1.xlsx");
            var lightWeightExcelReader = new ExcelReader(testFileLocation);
            var result = lightWeightExcelReader["Sheet1"].Row(4).ToArray();
            result[0].Should().Be(null);
            result[1].Should().Be(null);
            result[2].Should().Be(null);
            result[3].Should().Be(5);
        }

        [Fact]
        public void ServerPrefixedWorksheetWorks()
        {
            var list = new List<Object>();
            var testFileLocation = TestHelper.TestsheetPath("Import2.xlsx");
            var lightWeightExcelReader = new ExcelReader(testFileLocation);
            var sheet = lightWeightExcelReader["Prices"];
            while (sheet.ReadNext())
            {
                list.Add(sheet.Value);
            }
        }

        [Fact]
        public void FormulaReturnNullResultsIsIgnored()
        {
            var testFileLocation = TestHelper.TestsheetPath("FormulaReturningNullResultTest.xlsx");
            var lightWeightExcelReader = new ExcelReader(testFileLocation);
            var sheet = lightWeightExcelReader[0];
            sheet.Value.Should().BeNull();
            sheet.WorksheetDimension.TopLeft.ToString().Should().Be("A1");
            sheet.WorksheetDimension.BottomRight.ToString().Should().Be("D7");
            
            sheet.ReadNext();
            sheet.Value.Should().Be(1);
            sheet.Address.Should().Be("A1");

            sheet.ReadNext();
            sheet.Value.Should().Be(2);
            sheet.Address.Should().Be("C3");

            sheet.ReadNext();
            sheet.Address.Should().Be("A4");
            
            sheet.ReadNext();
            sheet.Address.Should().Be("A5");
            
            sheet.ReadNext();
            sheet.Address.Should().Be("A6");
            
            var result = sheet.ReadNext();
            sheet.Value.Should().BeNull();
            result.Should().BeFalse();
        }
        
        [Fact]
        public void ReadsNowFormulaCorrectly()
        {
            var testFileLocation = TestHelper.TestsheetPath("Now.xlsx");
            var lightWeightExcelReader = new ExcelReader(testFileLocation);
            var sheet = lightWeightExcelReader[0];
            sheet["A1"].Should().Be(new DateTime(2020, 12, 04,16,52,07, 977));
        }

        [Fact]
        public void ReadNextBehaviourSkipNullsWorksCorrectly()
        {
            var testFileLocation = TestHelper.TestsheetPath("ReadNextBehaviour.xlsx");
            var lightWeightExcelReader = new ExcelReader(testFileLocation);
            var sheet = lightWeightExcelReader[0];
            sheet.ReadNext();
            sheet.Value.Should().Be("A1");
            sheet.ReadNext();
            sheet.Value.Should().Be("B1");
            sheet.ReadNext();
            sheet.Value.Should().Be("C1");
            sheet.ReadNext();
            sheet.Value.Should().Be("D1");
            sheet.ReadNext();
            sheet.Value.Should().Be("E1");
            sheet.ReadNext();
            sheet.Value.Should().Be("A2");
            sheet.ReadNext();
            sheet.Value.Should().Be("C2");
            sheet.ReadNext();
            sheet.Value.Should().Be("E2");
            sheet.ReadNext();
            sheet.Value.Should().Be("A3");
            sheet.ReadNext();
            sheet.Value.Should().Be("C3");
            sheet.ReadNext();
            sheet.Value.Should().Be("A5");
            sheet.ReadNext();
            sheet.Value.Should().Be("B5");
            sheet.ReadNext();
            sheet.Value.Should().Be("C5");
            sheet.ReadNext();
            sheet.Value.Should().Be("D5");
            sheet.ReadNext();
            sheet.Value.Should().Be("E5");
        }
        
        [Fact]
        public void CellReadThenReadNextWorks()
        {
            var testFileLocation = TestHelper.TestsheetPath("ReadNextBehaviour.xlsx");
            var lightWeightExcelReader = new ExcelReader(testFileLocation);
            var sheet = lightWeightExcelReader[0];
            sheet["D5"].Should().Be("D5");
            sheet.Value.Should().Be("D5");
            sheet.ReadNext();
            sheet.Value.Should().Be("E5");
        }
        
        [Fact]
        public void ReadNextBehaviourReadsNextCorrectlyAfterAccessingViaIndexer()
        {
            var testFileLocation = TestHelper.TestsheetPath("ReadNextBehaviour.xlsx");
            var lightWeightExcelReader = new ExcelReader(testFileLocation)
            {
                ReadNextBehaviour = ReadNextBehaviour.ReadAllNulls
            };
            var sheet = lightWeightExcelReader[0];
            sheet["B2"].Should().Be(null);

            sheet.ReadNext();
            sheet.Address.Should().Be("C2");
            sheet.Value.Should().Be("C2");
        }

        [Fact]
        public void AddressIsSameAsIndexerForNullValues()
        {
            var testFileLocation = TestHelper.TestsheetPath("ReadNextBehaviour.xlsx");
            var lightWeightExcelReader = new ExcelReader(testFileLocation);
            var sheet = lightWeightExcelReader[0];
            sheet["E3"].Should().Be(null);
            sheet.Address.Should().Be("E3");
            sheet.Value.Should().Be(null);
        }

        [Fact]
        public void ReadNextBehaviourReadAllNullsClearsNextPopulatedCellIfCursorAdvancedBeyondIt()
        {
            var testFileLocation = TestHelper.TestsheetPath("ReadNextBehaviour.xlsx");
            var lightWeightExcelReader = new ExcelReader(testFileLocation)
            {
                ReadNextBehaviour = ReadNextBehaviour.ReadAllNulls
            };
            var sheet = lightWeightExcelReader[0];
            sheet["E3"].Should().Be(null);
            sheet.Address.Should().Be("E3");
            sheet.NextPopulatedCellRef.Should().BeNull();
        }
        
        [Fact]
        public void ReadNextBehaviourReadAllNullsReadsNextLineAfterReadNext()
        {
            var testFileLocation = TestHelper.TestsheetPath("ReadNextBehaviour.xlsx");
            var lightWeightExcelReader = new ExcelReader(testFileLocation)
            {
                ReadNextBehaviour = ReadNextBehaviour.ReadAllNulls
            };
            var sheet = lightWeightExcelReader[0];
            sheet["D3"].Should().BeNull();
            sheet.Address.Should().Be("D3");
            sheet.Value.Should().Be(null);
            sheet.AddressCelRef.ToString().Should().Be("D3");

            sheet.ReadNext();
            sheet.Address.Should().Be("E3");
            sheet.Value.Should().Be(null);
            sheet.AddressCelRef.ToString().Should().Be("E3");
            
            sheet.ReadNext();
            sheet.Address.Should().Be("A4");
            sheet.Value.Should().Be(null);
            sheet.AddressCelRef.ToString().Should().Be("A4");
        }

        [Fact]
        public void ReadNextBehaviourReadAllNullsWorksReadsNextLine()
        {
            var testFileLocation = TestHelper.TestsheetPath("ReadNextBehaviour.xlsx");
            var lightWeightExcelReader = new ExcelReader(testFileLocation)
            {
                ReadNextBehaviour = ReadNextBehaviour.ReadAllNulls
            };
            var sheet = lightWeightExcelReader[0];
            sheet["E3"].Should().Be(null);
            sheet.Address.Should().Be("E3");
            sheet.Value.Should().Be(null);
            
            sheet.ReadNext();
            sheet.Address.Should().Be("A4");
            sheet.Value.Should().Be(null);
        }

        [Fact]
        public void ReadNextBehaviourReadAllNullsReadsCorrectlyAfterBlankRow()
        {
            var testFileLocation = TestHelper.TestsheetPath("ReadNextBehaviour.xlsx");
            var lightWeightExcelReader = new ExcelReader(testFileLocation)
            {
                ReadNextBehaviour = ReadNextBehaviour.ReadAllNulls
            };
            var sheet = lightWeightExcelReader[0];
            sheet["E4"].Should().Be(null);
            sheet.Address.Should().Be("E4");
            sheet.Value.Should().Be(null);
            
            sheet.ReadNext();
            sheet.Address.Should().Be("A5");
            sheet.Value.Should().Be("A5");
            
            sheet.ReadNext();
            sheet.Address.Should().Be("B5");
            sheet.Value.Should().Be("B5");
        }

        [Fact]
        public void ReadNextBehaviourReadAllNullsReadsValueByIndex()
        {
            var testFileLocation = TestHelper.TestsheetPath("ReadNextBehaviour.xlsx");
            var lightWeightExcelReader = new ExcelReader(testFileLocation)
            {
                ReadNextBehaviour = ReadNextBehaviour.ReadAllNulls
            };
            var sheet = lightWeightExcelReader[0];
            sheet["A2"].Should().Be("A2");
            sheet.Value.Should().Be("A2");
        }

        [Fact]
        public void HandleInvalidFormula()
        {
            var testFileLocation = TestHelper.TestsheetPath("HandleInvalidFormulaTest.xlsx");
            var sheet1 = new ExcelReader(testFileLocation)["Sheet1"];
            sheet1["A1"].Should().Be("#DIV/0!");
            sheet1["A2"].Should().Be("#N/A");
            sheet1["A3"].Should().Be("#VALUE!");
            sheet1["A4"].Should().Be("#N/A");
        }

        [Fact]
        public void ReadNextBehaviourReadAllNullsWorksCorrectly()
        {
            var testFileLocation = TestHelper.TestsheetPath("ReadNextBehaviour.xlsx");
            var lightWeightExcelReader = new ExcelReader(testFileLocation)
            {
                ReadNextBehaviour = ReadNextBehaviour.ReadAllNulls
            };
            var sheet = lightWeightExcelReader[0];
            sheet["A2"].Should().Be("A2");
            sheet.Value.Should().Be("A2");
            
            sheet.ReadNext();
            sheet.Value.Should().Be(null);
            sheet.Address.Should().Be("B2");
            
            sheet.ReadNext();
            sheet.Value.Should().Be("C2");
            sheet.Address.Should().Be("C2");
            
            sheet.ReadNext();
            sheet.Address.Should().Be("D2");
            sheet.Value.Should().Be(null);
           
            sheet.ReadNext();
            sheet.Address.Should().Be("E2");
            sheet.Value.Should().Be("E2");
            
            sheet.ReadNext();
            sheet.Address.Should().Be("A3");
            sheet.Value.Should().Be("A3");
            
            sheet.ReadNext();
            sheet.Address.Should().Be("B3");
            sheet.Value.Should().Be(null);
            
            sheet.ReadNext();
            sheet.Address.Should().Be("C3");
            sheet.Value.Should().Be("C3");
            
            sheet.ReadNext();
            sheet.Address.Should().Be("D3");
            sheet.Value.Should().Be(null);
            
            sheet.ReadNext();
            sheet.Address.Should().Be("E3");
            sheet.Value.Should().Be(null);
            
            sheet.ReadNext();
            sheet.Address.Should().Be("A4");
            sheet.Value.Should().Be(null);
            
            sheet.ReadNext();
            sheet.Address.Should().Be("B4");
            sheet.Value.Should().Be(null);
            
            sheet.ReadNext();
            sheet.Address.Should().Be("C4");
            sheet.Value.Should().Be(null);
            
            sheet.ReadNext();
            sheet.Address.Should().Be("D4");
            sheet.Value.Should().Be(null);
            
            sheet.ReadNext();
            sheet.Address.Should().Be("E4");
            sheet.Value.Should().Be(null);
            
            sheet.ReadNext();
            sheet.Address.Should().Be("A5");
            sheet.Value.Should().Be("A5");
            
            sheet.ReadNext();
            sheet.Address.Should().Be("B5");
            sheet.Value.Should().Be("B5");
            
            sheet.ReadNext();
            sheet.Address.Should().Be("C5");
            sheet.Value.Should().Be("C5");
            
            sheet.ReadNext();
            sheet.Address.Should().Be("D5");
            sheet.Value.Should().Be("D5");
            
            sheet.ReadNext();
            sheet.Address.Should().Be("E5");
            sheet.Value.Should().Be("E5");

            sheet.ReadNext().Should().BeFalse();
        }
    }
}