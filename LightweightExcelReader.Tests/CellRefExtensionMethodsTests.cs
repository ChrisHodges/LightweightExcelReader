using FluentAssertions;
using SpreadsheetCellRef;
using Xunit;

namespace LightweightExcelReader.Tests
{
    public class CellRefExtensionMethodsTests
    {
        [Fact]
        public void IsNextAdjacentWorksForSameRow()
        {
            var a1 = new CellRef("A1");
            var b1 = new CellRef("B1");
            b1.IsNextAdjacentTo(a1).Should().BeTrue();
        }
        
        [Fact]
        public void IsNextAdjacentWorksForNextRow()
        {
            var a1 = new CellRef("A1");
            var a2 = new CellRef("A2");
            a2.IsNextAdjacentTo(a1).Should().BeTrue();
        }
        
        [Fact]
        public void IsNextAdjacentReturnsFalseFoPreviousAdjacentRow()
        {
            var a1 = new CellRef("A1");
            var b1 = new CellRef("B1");
            a1.IsNextAdjacentTo(b1).Should().BeFalse();
        }
        
        [Fact]
        public void IsNextAdjacentReturnsFalseForPreviousAdjacentRow()
        {
            var a1 = new CellRef("A1");
            var a2 = new CellRef("A2");
            a1.IsNextAdjacentTo(a2).Should().BeFalse();
        }
        
        [Fact]
        public void IsNextAdjacentReturnsFalseForNonAdjacent()
        {
            var a1 = new CellRef("A1");
            var c3 = new CellRef("C3");
            a1.IsNextAdjacentTo(c3).Should().BeFalse();
            c3.IsNextAdjacentTo(a1).Should().BeFalse();
        }
    }
}