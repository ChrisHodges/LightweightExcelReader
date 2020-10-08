using FluentAssertions;
using LightWeightExcelReader;
using Xunit;

namespace LightweighExcelReaderTests
{
    public class XslxlsDateTimeStreamtests
    {
        private XslxIsDateTimeStream _subject = new XslxIsDateTimeStream(TestHelper.TestXmlContent("Styles.xml"));

        [Fact]
        public void CanIdentifyDateTimeStyle()
        {
            _subject[14].Should().BeFalse();
            _subject[15].Should().BeTrue();
        }

        [Fact]
        public void CanGetFirstDateTimeStyle()
        {
            _subject.GetFirstDateTimeStyle().Should().Be(7);
        }
        
        [Fact]
        public void CanReadValuesBeforeAndThenGetFirstDateTimeStyle()
        {
            _subject[1].Should().BeFalse();
            _subject.GetFirstDateTimeStyle().Should().Be(7);
            _subject[14].Should().BeFalse();
            _subject[15].Should().BeTrue();
        }
        
        [Fact]
        public void CanReadValuesAfterAndThenGetFirstDateTimeStyle()
        {
            _subject[15].Should().BeTrue();
            _subject[14].Should().BeFalse();
            _subject.GetFirstDateTimeStyle().Should().Be(7);
            _subject[1].Should().BeFalse();
        }
    }
}