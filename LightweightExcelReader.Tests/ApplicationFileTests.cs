using System;
using System.Collections.Generic;
using FluentAssertions;
using LightWeightExcelReader;
using Xunit;

namespace LightweighExcelReaderTests
{
    public class ApplicationFileTests
    {
        [Fact]
        public void LibreOffice()
        {
            var testFileLocation = TestHelper.TestsheetPath("Applications/LibreOffice.xlsx");
            var excelReader = new ExcelReader(testFileLocation);
            var sheet = excelReader[0];
            sheet.WorksheetDimension.ToString().Should().Be("A1:K3");
            var dictionary = new Dictionary<string,object>();
            while (sheet.ReadNext())
            {
                dictionary.Add(sheet.Address, sheet.Value);
            }

            dictionary["A1"].Should().Be("String");
            dictionary["B1"].Should().Be(1);
            dictionary["C1"].Should().Be(1.2);
            dictionary["D1"].Should().Be(1.23);
            dictionary["E1"].Should().Be(new DateTime(2020,1,1));
            dictionary["F1"].Should().Be("String");
            dictionary["G1"].Should().Be(1);
            dictionary["H1"].Should().Be(new DateTime(2020,1,1));
            dictionary["I1"].Should().Be(2);
            dictionary["K3"].Should().Be("X");
            
            dictionary.Count.Should().Be(10);
        }
        
        [Fact]
        public void Numbers()
        {
            var testFileLocation = TestHelper.TestsheetPath("Applications/Numbers.xlsx");
            var excelReader = new ExcelReader(testFileLocation);
            var sheet = excelReader[0];
            sheet.WorksheetDimension.ToString().Should().Be("A1:K22");
            var dictionary = new Dictionary<string,object>();
            while (sheet.ReadNext())
            {
                dictionary.Add(sheet.Address, sheet.Value);
            }

            dictionary["A1"].Should().Be("String");
            dictionary["B1"].Should().Be(1);
            dictionary["C1"].Should().Be(1.2);
            dictionary["D1"].Should().Be(1.23);
            dictionary["E1"].Should().Be(new DateTime(2020,1,1));
            dictionary["F1"].Should().Be("String");
            dictionary["G1"].Should().Be(1);
            dictionary["H1"].Should().Be(new DateTime(2020,1,1));
            dictionary["I1"].Should().Be(2);
            dictionary["K3"].Should().Be("X");
            
            dictionary.Count.Should().Be(10);
        }
        
        [Fact]
        public void GoogleSheets()
        {
            var testFileLocation = TestHelper.TestsheetPath("Applications/GoogleSheets.xlsx");
            var excelReader = new ExcelReader(testFileLocation);
            var sheet = excelReader[0];
            var dictionary = new Dictionary<string,object>();
            while (sheet.ReadNext())
            {
                dictionary.Add(sheet.Address, sheet.Value);
            }

            dictionary["A1"].Should().Be("String");
            dictionary["B1"].Should().Be(1);
            dictionary["C1"].Should().Be(1.2);
            dictionary["D1"].Should().Be(1.23);
            dictionary["E1"].Should().Be(new DateTime(2020,1,1));//CSH 010120 Apple Numbers saves DateTimes in number format
            dictionary["F1"].Should().Be("String");
            dictionary["G1"].Should().Be(1);
            dictionary["H1"].Should().Be(new DateTime(2020,1,1));
            dictionary["I1"].Should().Be(2);
            dictionary["K3"].Should().Be("X");
            
            dictionary.Count.Should().Be(10);
        }
        
        [Fact]
        public void ExcelOnline()
        {
            var testFileLocation = TestHelper.TestsheetPath("Applications/ExcelOnline.xlsx");
            var excelReader = new ExcelReader(testFileLocation);
            var sheet = excelReader[0];
            var dictionary = new Dictionary<string,object>();
            while (sheet.ReadNext())
            {
                dictionary.Add(sheet.Address, sheet.Value);
            }

            dictionary["A1"].Should().Be("String");
            dictionary["B1"].Should().Be(1);
            dictionary["C1"].Should().Be(1.2);
            dictionary["D1"].Should().Be(1.23);
            dictionary["E1"].Should().Be(new DateTime(2020,1,1));//CSH 010120 Apple Numbers saves DateTimes in number format
            dictionary["F1"].Should().Be("String");
            dictionary["G1"].Should().Be(1);
            dictionary["H1"].Should().Be(new DateTime(2020,1,1));
            dictionary["I1"].Should().Be(2);
            dictionary["K3"].Should().Be("X");
            
            dictionary.Count.Should().Be(10);
        }
        
    }
}