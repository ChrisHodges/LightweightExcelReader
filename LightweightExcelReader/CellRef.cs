using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace LightweightExcelReader
{
    /// <summary>
    /// Struct. Represents the letter-number address of a cell in a worksheet.
    /// </summary>
    /// <example>
    /// <code>
    /// var a1 = new CellRef("A1");
    /// var alsoA1 = new CellRef(1,1);
    /// var b7 = new CellRef("B7"):
    /// var aa500 = new CellRef("AA500")
    /// var alsoAA500 = new CellRef(27,500);
    /// </code>
    /// </example>
    public struct CellRef
    {
        private static readonly Regex RowAndColumnMatcher = new Regex("([A-Z]+)([0-9]+)");
        private readonly string _cellRefString;

        /// <summary>
        /// Creates a cell ref from a string representation of a spreadsheet cell address, e.g. "A1" or "AZ36"
        /// </summary>
        /// <example>
        /// <code>
        /// var a1 = new CellRef("A1");
        /// </code>
        /// </example>
        /// <param name="cellRefString"></param>
        public CellRef(string cellRefString)
        {
            _cellRefString = cellRefString;
            var matches = RowAndColumnMatcher.Match(_cellRefString);
            Row = int.Parse(matches.Groups[2].Value);
            Column = matches.Groups[1].Value;
            ColumnNumber = ColumnNameToNumber(Column);
        }

        /// <summary>
        /// Creates a cell ref from the 1-based indexes of a cell row and column.
        /// </summary>
        /// <example>
        /// <code>
        /// var a1 = new CellRef(1,1);
        /// Console.WriteLine(a1.ToString()); //Outputs "A1"
        /// </code>
        /// </example>
        /// <param name="row"></param>
        /// <param name="column"></param>
        public CellRef(int row, int column)
        {
            Row = row;
            Column = NumberToColumnName(column);
            ColumnNumber = column;
            _cellRefString = $"{NumberToColumnName(column)}{row}";
        }

        /// <summary>
        /// Calling <c>ToString()</c> on a CellRef returns the address of the cell.
        /// </summary>
        /// <example>
        /// <code>
        /// var a1 = new CellRef("A1");
        /// Console.WriteLine(a1.ToString()); //Outputs "A1"
        /// </code>
        /// </example>
        /// <returns></returns>
        public override string ToString()
        {
            return _cellRefString;
        }

        /// <summary>
        /// The Column letter of the referenced cell
        /// </summary>
        public string Column { get; }

        /// <summary>
        /// The 1-based index of the CellRef column
        /// </summary>
        public int ColumnNumber { get; }

        /// <summary>
        /// The 1-based index of the CellRef row
        /// </summary>
        public int Row { get; }

        /// <summary>
        /// Static utility method to covert an column integer into its string equivalent.
        /// </summary>
        /// <example>
        /// <code>
        /// Console.WriteLine(CellRef.NumberToColumnName(27)); //Outputs "AA";
        /// </code>
        /// </example>
        /// <param name="i"></param>
        /// <returns></returns>
        public static string NumberToColumnName(int i)
        {
            var dividend = i;
            var columnName = "";
            int modulo;
            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo) + columnName;
                dividend = Convert.ToInt32((dividend - modulo) / 26);
            }

            return columnName;
        }

        /// <summary>
        /// Static utility method to covert a column string into its integer equivalent
        /// </summary>
        /// <example>
        /// <code>
        /// Console.WriteLine(CellRef.NumberToColumnName("AA").ToString()); //Outputs "27";
        /// </code>
        /// </example>
        /// <param name="c"></param>
        /// <returns></returns>
        public static int ColumnNameToNumber(string c)
        {
            var retVal = 0;
            var col = c.ToUpper();
            for (var iChar = col.Length - 1; iChar >= 0; iChar--)
            {
                var colPiece = col[iChar];
                var colNum = colPiece - 64;
                retVal = retVal + colNum * (int) Math.Pow(26, col.Length - (iChar + 1));
            }

            return retVal;
        }

        /// <summary>
        /// Returns a new CellRef that is equal to the current cell ref plus the given number of columns.
        /// </summary>
        /// <example>
        /// <code>
        /// var cellRef = new CellRef("A1");
        /// var newCellRef = cellRef.AddColumns(5);
        /// Console.WriteLine(newCellRef.ToString()); //outputs "F1"
        /// </code>
        /// </example>
        /// <param name="i"></param>
        /// <returns></returns>
        public CellRef AddColumns(int i)
        {
            var newColumnNumber = ColumnNumber + i;
            return new CellRef(Row, newColumnNumber);
        }

        /// <summary>
        /// Returns a new CellRef that is equal to the current cell ref plus the given number of columns.
        /// </summary>
        /// <example>
        /// <code>
        /// var cellRef = new CellRef("A1");
        /// var newCellRef = cellRef.AddRows(5);
        /// Console.WriteLine(newCellRef.ToString()); //outputs "A6"
        /// </code>
        /// </example>
        /// <param name="i"></param>
        /// <returns></returns>
        public CellRef AddRows(int i)
        {
            return new CellRef(Row + i, ColumnNumber);
        }

        /// <summary>
        /// Static method. Returns an enumerable of <c>CellRef</c>s representing all the cells in the given range. Cell order is
        /// left-to-right then top-to-bottom
        /// </summary>
        /// <example>
        /// <code>
        /// var cellRefs = CellRef.Range("A1","B2").ToArray();
        /// Console.WriteLine(cellRefs[0]); //outputs "A1"
        /// Console.WriteLine(cellRefs[1]); //outputs "B1"
        /// Console.WriteLine(cellRefs[2]); //outputs "A2"
        /// Console.WriteLine(cellRefs[3]); //outputs "B2"
        /// </code>
        /// </example>
        /// <param name="topLeft"></param>
        /// <param name="bottomRight"></param>
        /// <returns></returns>
        public static IEnumerable<CellRef> Range(string topLeft, string bottomRight)
        {
            var tl = new CellRef(topLeft);
            var br = new CellRef(bottomRight);
            var list = new List<CellRef>();

            for (var y = tl.Row; y <= br.Row; y++)
            {
                for (var x = tl.ColumnNumber; x <= br.ColumnNumber; x++)
                {
                    list.Add(new CellRef(y, x));
                }
            }

            return list;
        }
    }
}