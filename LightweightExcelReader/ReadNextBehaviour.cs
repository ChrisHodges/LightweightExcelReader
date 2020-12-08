namespace LightweightExcelReader
{
    /// <summary>
    /// Defines how the reader will handle null cells when using <c>SheetReader.ReadNext()</c>
    /// and <c>SheetReader.ReadNextInRow()</c>
    /// </summary>
    public enum ReadNextBehaviour
    {
        /// <summary>
        /// Default behaviour - calling ReadNext() will read the next non-null value
        /// </summary>
        SkipNulls,
        
        /// <summary>
        /// calling ReadNext() will read the next null or non-null value on the current 
        /// row (as far as the reported dimension of the spreadsheet) but will skip empty rows.
        /// </summary>
        ReadNullsOnPopulatedRows,
        
        /// <summary>
        /// calling ReadNext() will read the next non-null cell that is within the 
        /// reported dimension of the spreadsheet.
        /// </summary>
        ReadAllNulls
    }
}