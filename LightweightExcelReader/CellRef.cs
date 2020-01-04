using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace LightWeightExcelReader
{
    public struct CellRef
    {
        private static readonly Regex _rowAndColumnMatcher = new Regex("([A-Z]+)([0-9]+)");
        private readonly string _cellRefString;

        public CellRef(string cellRefString)
        {
            _cellRefString = cellRefString;
            var matches = _rowAndColumnMatcher.Match(_cellRefString);
            Row = int.Parse(matches.Groups[2].Value);
            Column = matches.Groups[1].Value;
            ColumnNumber = ColumnNameToNumber(Column);
        }

        public CellRef(int row, int column)
        {
            Row = row;
            Column = NumberToColumnName(column);
            ColumnNumber = column;
            _cellRefString = $"{NumberToColumnName(column)}{row}";
        }

        public override string ToString()
        {
            return _cellRefString;
        }

        public string Column { get; }

        public int ColumnNumber { get; }

        public int Row { get; }

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

        public CellRef AddColumns(int i)
        {
            var newColumnNumber = ColumnNumber + i;
            return new CellRef(Row, newColumnNumber);
        }

        public CellRef AddRows(int i)
        {
            return new CellRef(Row + i, ColumnNumber);
        }

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