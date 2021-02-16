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
            sheet1[3,1].Should().Be("zyx987");
        }
        
        [Fact]
        public void CellRefIndexerWorks()
        {
            var cellRef = new CellRef("A4");
            var testFileLocation = TestHelper.TestsheetPath("TestSpreadsheet1.xlsx");
            var sheet1 = new ExcelReader(testFileLocation)["Sheet1"];
            sheet1[cellRef].Should().Be("zyx987");
        }
    }
}