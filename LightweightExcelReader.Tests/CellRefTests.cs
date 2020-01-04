using System.Linq;
using FluentAssertions;
using LightWeightExcelReader;
using Xunit;

namespace LightweighExcelReaderTests
{
    public class CellRefTests
    {
        [Fact]
        public void AddColumnsWorks()
        {
            var a1 = new CellRef("A1");
            var g1 = a1.AddColumns(6);
            g1.ToString().Should().Be("G1");
        }

        [Fact]
        public void AddLargeColumnsWorks()
        {
            var aa1 = new CellRef("AA1");
            var ag1 = aa1.AddColumns(6);
            ag1.ToString().Should().Be("AG1");
        }


        [Fact]
        public void AddRowsWorks()
        {
            var a1 = new CellRef("A1");
            var a7 = a1.AddRows(6);
            a7.ToString().Should().Be("A7");
        }

        [Fact]
        public void CellNumberWorks()
        {
            var aa1 = new CellRef("AA1");
            aa1.ColumnNumber.Should().Be(27);
        }

        [Fact]
        public void ColumnNameToNumberWorks()
        {
            CellRef.ColumnNameToNumber("A").Should().Be(1);
            CellRef.ColumnNameToNumber("Z").Should().Be(26);
            CellRef.ColumnNameToNumber("AA").Should().Be(27);
        }

        [Fact]
        public void NumberToColumnNameWorks()
        {
            CellRef.NumberToColumnName(27).Should().Be("AA");
        }

        [Fact]
        public void RangeWorks()
        {
            var range = CellRef.Range("AA1", "AC3").ToArray();
            range[0].ToString().Should().Be("AA1");
            range[1].ToString().Should().Be("AB1");
            range[2].ToString().Should().Be("AC1");
            range[3].ToString().Should().Be("AA2");
            range[4].ToString().Should().Be("AB2");
            range[5].ToString().Should().Be("AC2");
            range[6].ToString().Should().Be("AA3");
            range[7].ToString().Should().Be("AB3");
            range[8].ToString().Should().Be("AC3");
        }

        [Fact]
        public void SingleColumnRangeWorks()
        {
            var range = CellRef.Range("A1", "A4").ToArray();
            range[0].ToString().Should().Be("A1");
            range[1].ToString().Should().Be("A2");
            range[2].ToString().Should().Be("A3");
            range[3].ToString().Should().Be("A4");
        }

        [Fact]
        public void ToStringWorks()
        {
            var aa1 = new CellRef("AA1");
            aa1.ToString().Should().Be("AA1");
        }
    }
}