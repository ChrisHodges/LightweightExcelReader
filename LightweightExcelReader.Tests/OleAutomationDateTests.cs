using System;
using System.Globalization;
using FluentAssertions;
using Xunit;

namespace LightweightExcelReader.Tests
{
    public class OleAutomationDateTests
    {
        [Fact]
        public void ParsesFractionalOleAutomationDateWithGermanLocale()
        {
            var cultureInfo = new CultureInfo("de-DE");
            CultureInfo.DefaultThreadCurrentCulture = cultureInfo;
            CultureInfo.DefaultThreadCurrentUICulture = cultureInfo;
            var date = DateTime.FromOADate(double.Parse("43831.0", CultureInfo.InvariantCulture));
            date.Should().Be(new DateTime(2020, 1, 1));
        }
        
        [Fact]
        public void ParsesFractionalOleAutomationDateWithUkLocale()
        {
            var cultureInfo = new CultureInfo("en-GB");
            CultureInfo.DefaultThreadCurrentCulture = cultureInfo;
            CultureInfo.DefaultThreadCurrentUICulture = cultureInfo;
            var date = DateTime.FromOADate(double.Parse("43831.0", CultureInfo.InvariantCulture));
            date.Should().Be(new DateTime(2020, 1, 1));
        }
    }
}