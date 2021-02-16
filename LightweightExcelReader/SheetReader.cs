using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml;
using SpreadsheetCellRef;

namespace LightweightExcelReader
{
    /// <summary>
    /// Reads values from a spreadsheet
    /// </summary>
    /// <example>
    /// <code>
    /// var workbookReader = new ExcelReader("Path/To/Workbook");
    /// var sheetReader = workbookReader["Sheet1"];
    /// var cellA1 = sheetReader["A1"];
    /// </code>
    /// </example>
    public class SheetReader
    {
        private static Regex _digitsRegex = new Regex(@"[0-9]");
        private readonly Dictionary<string, object> _values;
        private readonly XslxIsDateTimeStream _xlsxIsDateTimeStream;
        private readonly XslxSharedStringsStream _xlsxSharedStringsStream;
        private readonly XmlReader _xmlReader;
        private ReadNextBehaviour _readNextBehaviour;
        internal CellRef? AddressCelRef;
        internal CellRef? NextPopulatedCellRef;
        internal object NextPopulatedCellValue;

        internal SheetReader(Stream sheetXmlStream, XslxSharedStringsStream xlsxSharedStringsStream,
            XslxIsDateTimeStream xlsxIsDateTimeStream, ReadNextBehaviour readNextBehaviour)
        {
            _xlsxSharedStringsStream = xlsxSharedStringsStream;
            _xlsxIsDateTimeStream = xlsxIsDateTimeStream;
            _values = new Dictionary<string, object>();
            _xmlReader = XmlReader.Create(sheetXmlStream);
            _readNextBehaviour = readNextBehaviour;
            GetDimension();
        }

        /// <summary>
        ///     Indexer. Returns the value of the cell at the given address, e.g. sheetReader["C3"] returns the value
        ///     of the cell at C3, if present, or null if the cell is empty.
        /// </summary>
        /// <param name="cellAddress">
        ///     The address of the cell.
        /// </param>
        /// <exception cref="IndexOutOfRangeException">
        ///     Will throw if the requested index is beyond the used range of the workbook. Avoid this exception by checking the
        ///     WorksheetDimension or Max/MinRow and Max/MinColumnNumber properties.
        /// </exception>
        public object this[string cellAddress]
        {
            get
            {
                if (_values.ContainsKey(cellAddress))
                {
                    return _values[cellAddress];
                }

                var cellRef = new CellRef(cellAddress);
                if (cellRef.ColumnNumber > WorksheetDimension.BottomRight.ColumnNumber ||
                    cellRef.Row > WorksheetDimension.BottomRight.Row)
                {
                    throw new IndexOutOfRangeException();
                }

                var value = GetValue(cellAddress);
                return value;
            }
        }

        /// <summary>
        ///     Indexer. Returns the value of the cell at the given CellRef, e.g. sheetReader[new CellRef("C3")] returns the value
        ///     of the cell at C3, if present, or null if the cell is empty.
        /// </summary>
        /// <param name="cellRef"></param>
        /// <exception cref="IndexOutOfRangeException">
        ///     Will throw if the requested index is beyond the used range of the workbook. Avoid this exception by checking the
        ///     WorksheetDimension or Max/MinRow and Max/MinColumnNumber properties.
        /// </exception>
        public object this[CellRef cellRef]
        {
            get
            {
                var cellRefString = cellRef.ToString();
                if (_values.ContainsKey(cellRefString))
                {
                    return _values[cellRefString];
                }
                
                if (cellRef.ColumnNumber > WorksheetDimension.BottomRight.ColumnNumber ||
                    cellRef.Row > WorksheetDimension.BottomRight.Row)
                {
                    throw new IndexOutOfRangeException();
                }

                var value = GetValue(cellRefString);
                return value;
            }
        }

        /// <summary>
        ///     Indexer. Returns the value of the cell at the given string column and 1-based integer row values, e.g. sheetReader["C",7] returns the value
        ///     of the cell at C7, or null if the cell is empty.
        /// </summary>
        /// <param name="column"></param>
        /// <param name="row"></param>
        /// <exception cref="IndexOutOfRangeException">
        ///     Will throw if the requested index is beyond the used range of the workbook. Avoid this exception by checking the
        ///     WorksheetDimension or Max/MinRow and Max/MinColumnNumber properties.
        /// </exception>
        public object this[string column, int row]
        {
            get
            {
                var cellRef = new CellRef(row, CellRef.ColumnNameToNumber(column));
                var cellRefString = cellRef.ToString();
                if (_values.ContainsKey(cellRefString))
                {
                    return _values[cellRefString];
                }
                
                if (cellRef.ColumnNumber > WorksheetDimension.BottomRight.ColumnNumber ||
                    cellRef.Row > WorksheetDimension.BottomRight.Row)
                {
                    throw new IndexOutOfRangeException();
                }

                var value = GetValue(cellRefString);
                return value;
            }
        }
        
        /// <summary>
        ///     Indexer. Returns the value of the cell at the given 1-based row and column values, e.g. sheetReader[5,6] returns the value
        ///     of the cell at row 5, column 6, or null if the cell is empty.
        /// </summary>
        /// <param name="row"></param>
        /// <param name="column"></param>
        /// <exception cref="IndexOutOfRangeException">
        ///     Will throw if the requested index is beyond the used range of the workbook. Avoid this exception by checking the
        ///     WorksheetDimension or Max/MinRow and Max/MinColumnNumber properties.
        /// </exception>
        public object this[int row, int column]
        {
            get
            {
                var cellRef = new CellRef(row, column);
                var cellRefString = cellRef.ToString();
                if (_values.ContainsKey(cellRefString))
                {
                    return _values[cellRefString];
                }
                
                if (cellRef.ColumnNumber > WorksheetDimension.BottomRight.ColumnNumber ||
                    cellRef.Row > WorksheetDimension.BottomRight.Row)
                {
                    throw new IndexOutOfRangeException();
                }

                var value = GetValue(cellRefString);
                return value;
            }
        }

        /// <summary>
        ///     Get a list of cell values covered by the range in the index, e.g. sheetReader["A1","B2"] will return a list of four
        ///     values,
        ///     going left-to-right and then top-to-bottom, from the cells A1, B1, A2, B2.
        /// </summary>
        /// <example>
        /// <code>
        /// var range = sheetReader["A1","B2"];
        /// </code>
        /// </example>
        /// <param name="topLeft">The top left cell of the required range</param>
        /// <param name="bottomRight">The bottom right cell of the required range</param>
        public IEnumerable<object> this[string topLeft, string bottomRight]
        {
            get
            {
                var range = CellRef.Range(topLeft, bottomRight).Select(x => x.ToString());
                var result = range.Select(x => this[x]);
                return result;
            }
        }

        /// <summary>
        ///     A <see cref="WorksheetDimension"/> representing the used range of the worksheet
        /// </summary>
        public WorksheetDimension WorksheetDimension { get; private set; }

        /// <summary>
        ///     The cell address of the most recently read cell of the spreadsheet
        /// </summary>
        public string Address { get; private set; }

        /// <summary>
        ///     The largest column number used by the spreadsheet
        /// </summary>
        public int MaxColumnNumber { get; private set; }

        /// <summary>
        ///     The smallest column number used by the spreadsheet
        /// </summary>
        public int MinColumnNumber { get; private set; }

        /// <summary>
        ///     The largest row number used by the spreadsheet
        /// </summary>
        public int MaxRow { get; private set; }

        /// <summary>
        ///     The smallest row number used by the spreadsheet
        /// </summary>
        public int MinRow { get; private set; }

        /// <summary>
        ///     The value of the last cell read by the reader. This will be null if:
        ///     - The sheet has not yet been read
        ///     - The ReadNextInRow() method read to the end of the row
        ///     - The ReadNext() method read to the end of the sheet
        /// </summary>
        public object Value { get; private set; }

        /// <summary>
        ///     The 1-based row number of the most recently read cell. This will be null if the spreadsheet has not yet been read.
        /// </summary>
        public int? CurrentRowNumber { get; private set; }

        /// <summary>
        ///     The number of the penultimate row read by the reader. This will be null if the reader has read zero or one rows.
        ///     This property is useful when checking for blank rows.
        /// </summary>
        public int? PreviousRowNumber { get; private set; }

        private object ReadNumericValueFromSpreadsheet(string sType)
        {
            object value;
            if (sType != null && _xlsxIsDateTimeStream[int.Parse(sType)])
            {
                value = DateTime.FromOADate(double.Parse(_xmlReader.Value));
            }
            else
            {
                value = double.Parse(_xmlReader.Value, CultureInfo.InvariantCulture);
            }

            return value;
        }

        private object GetValueFromCell(string nodeType, string sType)
        {
            object value;
            if (string.IsNullOrEmpty(_xmlReader.Value))
            {
                return null;
            }

            switch (nodeType)
            {
                case "d":
                    value = DateTime.Parse(_xmlReader.Value);
                    break;
                case "str":
                    value = _xmlReader.Value;
                    break;
                case "s":
                    value = _xlsxSharedStringsStream[int.Parse(_xmlReader.Value)];
                    break;
                case "b":
                    value = _xmlReader.Value == "1";
                    break;
                default:
                    value = ReadNumericValueFromSpreadsheet(sType);
                    break;
            }

            return value;
        }

        private KeyValuePair<string, object> ReadNextCell()
        {
            var sType = _xmlReader.GetAttribute("s");
            var nodeType = _xmlReader.GetAttribute("t");
            var address = _xmlReader.GetAttribute("r");
            var newValue = ReadValue(nodeType, sType);
            return new KeyValuePair<string, object>(address, newValue);
        }

        private void GetCellAttributesAndReadValue()
        {
            var nextCell = ReadNextCell();
            switch (_readNextBehaviour)
            {
                case ReadNextBehaviour.SkipNulls:
                    Address = nextCell.Key;
                    Value = nextCell.Value;
                    break;
                case ReadNextBehaviour.ReadAllNulls:
                    //If first cell read:
                    if (!AddressCelRef.HasValue)
                    {
                        Address = nextCell.Key;
                        Value = nextCell.Value;
                        AddressCelRef = new CellRef(Address);
                    }
                    else
                    {
                        var nextCellRef = new CellRef(nextCell.Key);
                        //If not first cell read bit adjacent to first cell
                        if (nextCellRef.IsNextAdjacentTo(AddressCelRef))
                        {
                            Address = nextCell.Key;
                            Value = nextCell.Value;
                            AddressCelRef = nextCellRef;
                        }
                        //If not first cell read and not adjacent to first cell
                        else
                        {
                            var nextAdjacent =
                                AddressCelRef.Value.GetNextAdjacent(WorksheetDimension.BottomRight.ColumnNumber);
                            Address = nextAdjacent.ToString();
                            Value = null;
                            AddressCelRef = nextAdjacent;
                            NextPopulatedCellRef = nextCellRef;
                            NextPopulatedCellValue = nextCell.Value;
                        }
                    }
                    break;
                default:
                    throw new NotImplementedException();
            }
        }

        private object ReadValue(string nodeType, string sType)
        {
            object newValue = null;
            while (ReadNextXmlElementAndLogRowNumber())
            {
                if (_xmlReader.IsStartOfElement("v"))
                {
                    ReadNextXmlElementAndLogRowNumber();
                    newValue = GetValueFromCell(nodeType, sType);
                }

                if (_xmlReader.IsStartOfElement("t"))
                {
                    ReadNextXmlElementAndLogRowNumber();
                    newValue = _xmlReader.Value;
                }

                if (_xmlReader.IsEndOfElement("c"))
                {
                    break;
                }
            }

            return newValue;
        }

        private bool ReadNextXmlElementAndLogRowNumber()
        {
            var result = _xmlReader.Read();
            if (_xmlReader.IsStartOfElement("row"))
            {
                PreviousRowNumber = CurrentRowNumber;
                CurrentRowNumber = int.Parse(_xmlReader.GetAttribute("r"));
            }

            return result;
        }

        private void SetCursorsToCurrentNullValue(string address)
        {
            _values[address] = null;
            AddressCelRef = new CellRef(address);
            Address = address;
            Value = null;
            NextPopulatedCellRef = null;
            NextPopulatedCellValue = null;
        }

        private object GetValue(string address)
        {
            var cellRef = new CellRef(address);
            while (ReadNextXmlElementAndLogRowNumber())
            {
                if (_xmlReader.IsStartOfElement("c") && !_xmlReader.IsEmptyElement)
                {
                    GetCellAttributesAndReadValue();
                    _values[Address] = Value;
                    if (Address == address)
                    {
                        return Value;
                    }

                    if (CurrentRowNumber == cellRef.Row)
                    {
                        var columnLetter = _digitsRegex.Replace(Address, "");
                        var currentColumnNumber = CellRef.ColumnNameToNumber(columnLetter);
                        if (currentColumnNumber > cellRef.ColumnNumber)
                        {
                            SetCursorsToCurrentNullValue(address);
                            return null;
                        }
                    }
                }

                if (_xmlReader.IsStartOfElement("row") && CurrentRowNumber > cellRef.Row)
                {
                    SetCursorsToCurrentNullValue(address);
                    return null;
                }
            }

            return null;
        }

        /// <summary>
        /// Reads the next cell in the spreadsheet, updating the readers value and address properties.
        /// </summary>
        /// <returns>False if all cells have been read, true otherwise</returns>
        public bool ReadNext()
        {
            if (NextPopulatedCellRef.HasValue)
            {
                var nextAdjacentCellRef =
                    AddressCelRef.Value.GetNextAdjacent(WorksheetDimension.BottomRight.ColumnNumber);
                if (nextAdjacentCellRef == NextPopulatedCellRef)
                {
                    AddressCelRef = NextPopulatedCellRef;
                    Address = NextPopulatedCellRef.ToString();
                    Value = NextPopulatedCellValue;
                    NextPopulatedCellRef = null;
                    NextPopulatedCellValue = null;
                }
                else
                {
                    Address = nextAdjacentCellRef.ToString();
                    AddressCelRef = new CellRef(Address);
                    Value = null;
                }

                return true;
            }

            while (ReadNextXmlElementAndLogRowNumber())
            {
                if (_xmlReader.IsStartOfElement("c") && !_xmlReader.IsEmptyElement)
                {
                    GetCellAttributesAndReadValue();
                    if (Value == null && _readNextBehaviour == ReadNextBehaviour.SkipNulls)
                    {
                        return ReadNext();
                    }

                    _values[Address] = Value;
                    return true;
                }

                if (_xmlReader.IsEndOfElement("sheetData"))
                {
                    Address = null;
                    Value = null;
                    return false;
                }
            }

            Address = null;
            Value = null;
            return false;
        }

        private void GetDimension()
        {
            while (ReadNextXmlElementAndLogRowNumber())
            {
                if (_xmlReader.IsStartOfElement("dimension"))
                {
                    var rangeRef = _xmlReader.GetAttribute("ref");
                    var regex = new Regex("^([A-Z]+[0-9]+)(?::([A-Z]+[0-9]+))?$");
                    var match = regex.Match(rangeRef);
                    var topLeftRange = match.Groups[1].Value;
                    var bottomRightRange = match.Groups[2].Value;
                    var topLeft = new CellRef(topLeftRange);
                    var bottomRight = bottomRightRange != ""
                        ? new CellRef(bottomRightRange)
                        : new CellRef(topLeftRange);
                    WorksheetDimension = new WorksheetDimension {TopLeft = topLeft, BottomRight = bottomRight};
                    MinRow = topLeft.Row;
                    MaxRow = bottomRight.Row;
                    MinColumnNumber = topLeft.ColumnNumber;
                    MaxColumnNumber = bottomRight.ColumnNumber;
                    break;
                }

                if (_xmlReader.IsStartOfElement("sheetData"))
                {
                    break;
                }
            }
        }

        /// <summary>
        /// Gets a list of all the cell values within the specified column.
        /// </summary>
        /// <param name="column">The string representation of the column, e.g. A, C, AAZ, etc. </param>
        /// <returns>An enumerable of objects representing the values of cells in the column</returns>
        public IEnumerable<object> Column(string column)
        {
            return Column(CellRef.ColumnNameToNumber(column));
        }

        /// <summary>
        /// Gets a list of all the cell values at the specified ordinal column index.
        /// </summary>
        /// <param name="column">The column index </param>
        /// <returns>An enumerable of objects representing the values of cells in the column</returns>
        public IEnumerable<object> Column(int column)
        {
            var topLeft = new CellRef(MinRow, column);
            var bottomRight = new CellRef(MaxRow, column);
            return this[topLeft.ToString(), bottomRight.ToString()];
        }

        /// <summary>
        /// Gets a list of all the cell values in the specified row
        /// </summary>
        /// <param name="row">The 1-based row index</param>
        /// <returns>An enumerable of objects representing the values of cells in the row</returns>
        public IEnumerable<object> Row(int row)
        {
            var topLeft = new CellRef(row, MinColumnNumber);
            var bottomRight = new CellRef(row, MaxColumnNumber);
            return this[topLeft.ToString(), bottomRight.ToString()];
        }

        /// <summary>
        /// Reads the next cell in the row and updates the reader's value and address properties
        /// </summary>
        /// <returns>False if there are no more cells in the row, true otherwise</returns>
        public bool ReadNextInRow()
        {
            do
            {
                if (_xmlReader.IsStartOfElement("c") && !_xmlReader.IsEmptyElement)
                {
                    GetCellAttributesAndReadValue();
                    _values[Address] = Value;
                    return true;
                }

                if (_xmlReader.IsEndOfElement("row"))
                {
                    break;
                }
            } while (ReadNextXmlElementAndLogRowNumber());

            Address = null;
            Value = null;
            return false;
        }

        /// <summary>
        /// Returns <c>true</c> if the specified cell contains a non-null value.
        /// </summary>
        /// <param name="cellRefString"></param>
        /// <returns></returns>
        public bool ContainsKey(string cellRefString)
        {
            var cellRef = new CellRef(cellRefString);
            if (cellRef.ColumnNumber > WorksheetDimension.BottomRight.ColumnNumber ||
                cellRef.Row > WorksheetDimension.BottomRight.Row)
            {
                return false;
            }

            var result = this[cellRefString];
            if (result == null)
            {
                return false;
            }

            return true;
        }
    }
}