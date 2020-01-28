using System.Xml;

namespace LightWeightExcelReader
{
    public static class XmlReaderExtensionMethods
    {
        public static bool IsStartOfElement(this XmlReader xmlReader, string elementName)
        {
            return xmlReader.LocalName == elementName && xmlReader.NodeType == XmlNodeType.Element;
        }

        public static bool IsEndOfElement(this XmlReader xmlReader, string elementName)
        {
            return xmlReader.LocalName == elementName && xmlReader.NodeType == XmlNodeType.EndElement;
        }
    }
}