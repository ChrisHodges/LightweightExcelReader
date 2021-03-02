using System.IO;
using System.Reflection;
using System.Text.RegularExpressions;
using FluentAssertions;

namespace LightweightExcelReader.Tests
{
    public static class TestHelper
    {
        public static string TestsheetPath(string spreadsheetName)
        {
            var assembly = Assembly.GetExecutingAssembly();
            var assemblyPath = Path.GetDirectoryName(assembly.Location);
            var testSpreadsheetLocation = Path.Combine(assemblyPath, "TestSpreadsheets", spreadsheetName);
            File.Exists(testSpreadsheetLocation).Should().BeTrue();
            return testSpreadsheetLocation;
        }
        
        public static Stream TestXmlContent(string fileName)
        {
            var assembly = Assembly.GetExecutingAssembly();
            var assemblyPath = Path.GetDirectoryName(assembly.Location);
            var testSpreadsheetLocation = Path.Combine(assemblyPath, "TestXml", fileName);
            File.Exists(testSpreadsheetLocation).Should().BeTrue();
            return new FileStream(testSpreadsheetLocation, FileMode.OpenOrCreate, FileAccess.Read);
        }
    }
}