using SpreadsheetCellRef;

namespace LightweightExcelReader
{
    internal static class CellRefExtensionMethods
    {
        internal static bool IsNextAdjacentTo(this CellRef thisCellRef, CellRef otherCellRef)
        {
            if (thisCellRef.Row == otherCellRef.Row && thisCellRef.ColumnNumber  == otherCellRef.ColumnNumber + 1)
            {
                return true;
            }
            if (thisCellRef.Row == otherCellRef.Row + 1 && thisCellRef.ColumnNumber == 1)
            {
                return true;
            }
            return false;
        }
        
        internal static bool IsNextAdjacentTo(this CellRef thisCellRef, CellRef? otherCellRef)
        {
            if (!otherCellRef.HasValue)
            {
                return false;
            }
            if (thisCellRef.Row == otherCellRef.Value.Row && thisCellRef.ColumnNumber  == otherCellRef.Value.ColumnNumber + 1)
            {
                return true;
            }
            if (thisCellRef.Row == otherCellRef.Value.Row + 1 && thisCellRef.ColumnNumber == 1)
            {
                return true;
            }
            return false;
        }

        internal static CellRef GetNextAdjacent(this CellRef cellRef, int mostRightColumn)
        {
            return cellRef.ColumnNumber >= mostRightColumn ? cellRef.GetFirstInNextRow() : cellRef.GetNextRight();
        }

        internal static CellRef GetNextRight(this CellRef cellRef)
        {
            return new CellRef(cellRef.Row, cellRef.ColumnNumber + 1);
        }

        internal static CellRef GetFirstInNextRow(this CellRef cellRef)
        {
            return new CellRef(cellRef.Row + 1, 1);
        }
    }
}