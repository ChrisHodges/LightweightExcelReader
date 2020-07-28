using System.IO;
using System.Reflection;
using System.Text.RegularExpressions;
using FluentAssertions;

namespace LightweighExcelReaderTests
{
    public static class TestHelper
    {
        public static string TestsheetPath(string spreadsheetName)
        {
            var assembly = Assembly.GetExecutingAssembly();
            var assemblyPath = Path.GetDirectoryName(assembly.GetName().CodeBase).Substring(5);
            assemblyPath =
                Regex.Replace(assemblyPath, @"^\\+(?<drive>[A-Z]:)", "${drive}"); //Fix for windows based file systems
            var testSpreadsheetLocation = Path.Combine(assemblyPath, "TestSpreadsheets", spreadsheetName);
            File.Exists(testSpreadsheetLocation).Should().BeTrue();
            return testSpreadsheetLocation;
        }
        
        public static Stream TestXmlContent(string fileName)
        {
            var assembly = Assembly.GetExecutingAssembly();
            var assemblyPath = Path.GetDirectoryName(assembly.GetName().CodeBase).Substring(5);
            assemblyPath =
                Regex.Replace(assemblyPath, @"^\\+(?<drive>[A-Z]:)", "${drive}"); //Fix for windows based file systems
            var testSpreadsheetLocation = Path.Combine(assemblyPath, "TestXml", fileName);
            File.Exists(testSpreadsheetLocation).Should().BeTrue();
            return new FileStream(testSpreadsheetLocation, FileMode.OpenOrCreate, FileAccess.Read);
        }
    }
}