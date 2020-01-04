using System;
using System.Collections.Generic;
using System.IO;
using System.Xml;

namespace LightWeightExcelReader
{
    /// <summary>
    /// A reader for the entire workbook. Access an individual worksheet by the worksheet name indexer,
    /// e.g. excelReader["WorkSheet"] or by it's zero-based index, e.g. excelReader[0]
    /// </summary>
    public class ExcelReader
    {
        private readonly string _filePath;
        private Dictionary<string, int> _sheetnameLookup;
        private XmlReader _sheetNameXmlReader;
        private int _sheetNumberIndex = -1;
        private Dictionary<int, SheetReader> _sheetReadersByInteger;
        private ZippedXlsxFile _zippedXlsxFile;

        /// <summary>
        /// Construct an ExcelReader from a file path
        /// </summary>
        /// <param name="filePath">A file path pointing towards an xlsx format workbok</param>
        public ExcelReader(string filePath)
        {
            _filePath = filePath;
            _zippedXlsxFile = new ZippedXlsxFile(filePath);
        }
        
        /// <summary>
        /// Construct an ExcelReader from a Stream
        /// </summary>
        /// <param name="stream">A stream pointing towards an xlsx format workbook</param>

        public ExcelReader(Stream stream)
        {
            _zippedXlsxFile = new ZippedXlsxFile(stream);
        }
        
        /// <summary>
        /// Get a SheetReader instance representing the worksheet at the given zero-based index
        /// </summary>
        /// <param name="sheetNumber">The zero-based index of the worksheet</param>
        public SheetReader this[int sheetNumber]
        {
            get
            {
                if (_sheetReadersByInteger == null)
                {
                    _sheetReadersByInteger = new Dictionary<int, SheetReader>();
                }

                if (_sheetReadersByInteger.ContainsKey(sheetNumber))
                {
                    return _sheetReadersByInteger[sheetNumber];
                }

                if (_zippedXlsxFile == null)
                {
                    _zippedXlsxFile = new ZippedXlsxFile(_filePath);
                }

                var sheetReader = new SheetReader(_zippedXlsxFile.GetWorksheetStream(sheetNumber),
                    _zippedXlsxFile.SharedStringsStream, _zippedXlsxFile.IsDateTimeStream);
                _sheetReadersByInteger.Add(sheetNumber, sheetReader);
                return sheetReader;
            }
        }

        /// <summary>
        /// Get a SheetReader instance representing the worksheet with the given name
        /// </summary>
        /// <param name="sheetName">The name of the worksheet</param>
        /// <exception cref="IndexOutOfRangeException">Will throw if the worksheet does not exist</exception>
        public SheetReader this[string sheetName]
        {
            get
            {
                if (_sheetnameLookup == null)
                {
                    _sheetnameLookup = new Dictionary<string, int>();
                }

                if (_zippedXlsxFile == null)
                {
                    _zippedXlsxFile = new ZippedXlsxFile(_filePath);
                }

                if (_sheetReadersByInteger == null)
                {
                    _sheetReadersByInteger = new Dictionary<int, SheetReader>();
                }

                int? sheetNumber;
                if (_sheetnameLookup.ContainsKey(sheetName))
                {
                    sheetNumber = _sheetnameLookup[sheetName];
                }
                else
                {
                    sheetNumber = ReadSheetNumberFromXml(sheetName);
                }

                if (!sheetNumber.HasValue)
                {
                    throw new IndexOutOfRangeException();
                }

                if (_sheetReadersByInteger.ContainsKey(sheetNumber.Value))
                {
                    var existingSheet = _sheetReadersByInteger[sheetNumber.Value];
                    return existingSheet;
                }

                var sheetReader = new SheetReader(_zippedXlsxFile.GetWorksheetStream(sheetNumber.Value),
                    _zippedXlsxFile.SharedStringsStream, _zippedXlsxFile.IsDateTimeStream);
                _sheetReadersByInteger.Add(sheetNumber.Value, sheetReader);
                return sheetReader;
            }
        }

        private int? ReadSheetNumberFromXml(string sheetName)
        {
            if (_sheetNameXmlReader == null)
            {
                _sheetNameXmlReader = XmlReader.Create(_zippedXlsxFile.WorkbookXml);
            }

            if (_sheetnameLookup == null)
            {
                _sheetnameLookup = new Dictionary<string, int>();
            }

            while (_sheetNameXmlReader.Read())
            {
                if (_sheetNameXmlReader.IsStartOfElement("sheet"))
                {
                    _sheetNumberIndex++;
                    var currentSheetName = _sheetNameXmlReader.GetAttribute("name");
                    _sheetnameLookup.Add(currentSheetName, _sheetNumberIndex);
                    if (currentSheetName == sheetName)
                    {
                        return _sheetNumberIndex;
                    }
                }

                if (_sheetNameXmlReader.IsEndOfElement("sheets"))
                {
                    break;
                }
            }

            return null;
        }
    }
}