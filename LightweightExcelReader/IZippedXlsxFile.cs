using System;
using System.IO;

namespace LightweightExcelReader
{
    internal interface IZippedXlsxFile : IDisposable
    {
        Stream WorkbookXml { get; }
        XslxSharedStringsStream SharedStringsStream { get; }
        XslxIsDateTimeStream IsDateTimeStream { get; }
        Stream GetWorksheetStream(int i);
    }
}