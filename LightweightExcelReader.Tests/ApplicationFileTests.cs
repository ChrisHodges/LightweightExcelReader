using System;
using System.Collections.Generic;
using System.Globalization;
using FluentAssertions;
using LightweightExcelReader.Tests.TestHelpers;
using Xunit;

namespace LightweightExcelReader.Tests
{
    public class ApplicationFileTests
    {

        [Theory]
        [ClassData(typeof(CultureGenerator))]
        public void LibreOffice(CultureInfo cultureInfo)
        {
            CultureInfo.DefaultThreadCurrentCulture = cultureInfo;
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
        
        [Theory]
        [ClassData(typeof(CultureGenerator))]
        public void Numbers(CultureInfo cultureInfo)
        {
            CultureInfo.DefaultThreadCurrentCulture = cultureInfo;
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

        [Theory]
        [ClassData(typeof(CultureGenerator))]
        public void GoogleSheets(CultureInfo cultureInfo)
        {
            CultureInfo.DefaultThreadCurrentCulture = cultureInfo;
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
            dictionary["E1"].Should().Be(new DateTime(2020,1,1));
            dictionary["F1"].Should().Be("String");
            dictionary["G1"].Should().Be(1);
            dictionary["H1"].Should().Be(new DateTime(2020,1,1));
            dictionary["I1"].Should().Be(2);
            dictionary["K3"].Should().Be("X");
            
            dictionary.Count.Should().Be(10);
        }
        
        [Theory]
        [ClassData(typeof(CultureGenerator))]
        public void ExcelOnline(CultureInfo cultureInfo)
        {
            CultureInfo.DefaultThreadCurrentCulture = cultureInfo;
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
            dictionary["E1"].Should().Be(new DateTime(2020,1,1));
            dictionary["F1"].Should().Be("String");
            dictionary["G1"].Should().Be(1);
            dictionary["H1"].Should().Be(new DateTime(2020,1,1));
            dictionary["I1"].Should().Be(2);
            dictionary["K3"].Should().Be("X");
            
            dictionary.Count.Should().Be(10);
        }
        
    }
}