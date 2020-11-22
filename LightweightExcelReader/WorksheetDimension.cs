namespace LightweightExcelReader
{
    /// <summary>
    /// Represents the used range of a worksheet
    /// </summary>
    public class WorksheetDimension
    {
        internal WorksheetDimension()
        {
            
        }
        /// <summary>
        /// The top left cell in the used range. 
        /// </summary>
        public CellRef TopLeft { get; internal set; }
        
        /// <summary>
        /// The bottom right cell in the used range.
        /// </summary>
        public CellRef BottomRight { get; internal set; }

        /// <summary>
        /// Calling <c>ToString()</c> on a <c>WorksheetDimension</c> instance returns the top left and bottom right cell refs, separated by a colom
        /// </summary>
        /// <example>
        /// <code>
        /// Console.WriteLine(worksheet.WorksheetDimension.ToString()); //outputs, for example, 'A1:C17'
        /// </code>
        /// </example>
        /// <returns></returns>
        public override string ToString()
        {
            return $"{TopLeft}:{BottomRight}";
        }
    }
}