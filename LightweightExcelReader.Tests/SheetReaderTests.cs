using System;
using System.Linq;
using FluentAssertions;
using SpreadsheetCellRef;
using Xunit;

namespace LightweightExcelReader.Tests
{
    public class SheetReaderTests
    {
        [Fact]
        public void StringIndexerWorks()
        {
            var testFileLocation = TestHelper.TestsheetPath("TestSpreadsheet1.xlsx");
            var sheet1 = new ExcelReader(testFileLocation)["Sheet1"];
            sheet1["A3"].Should().Be("zyx987");
        }

        [Fact]
        public void IntPairIndexerWorks()
        {
            var testFileLocation = TestHelper.TestsheetPath("TestSpreadsheet1.xlsx");
            var sheet1 = new ExcelReader(testFileLocation)["Sheet1"];
            sheet1[3, 1].Should().Be("zyx987");
        }

        [Fact]
        public void CellRefIndexerWorks()
        {
            var cellRef = new CellRef("A3");
            var testFileLocation = TestHelper.TestsheetPath("TestSpreadsheet1.xlsx");
            var sheet1 = new ExcelReader(testFileLocation)["Sheet1"];
            sheet1[cellRef].Should().Be("zyx987");
        }
        
        [Fact]
        public void StringIntIndexerWorks()
        {
            var testFileLocation = TestHelper.TestsheetPath("TestSpreadsheet1.xlsx");
            var sheet1 = new ExcelReader(testFileLocation)["Sheet1"];
            sheet1["A",3].Should().Be("zyx987");
        }

        /// <summary>
        /// Prompted by this Stack Overflow question:
        /// https://stackoverflow.com/questions/66993917/check-if-ienumerableobject-has-any-values-inside-without-getting-indexoutofran/66994213#66994213
        /// </summary>
        [Fact]
        public void OutOfRangeRowThrowsOutOfRangeExceptionImmediately()
        {
            var testFileLocation = TestHelper.TestsheetPath("TestSpreadsheet1.xlsx");
            var sheet1 = new ExcelReader(testFileLocation)["Sheet1"];
            Action action = () =>
            {
                var row = sheet1.Row(5);
            };
            action.Should().Throw<IndexOutOfRangeException>();
        }
    }
}