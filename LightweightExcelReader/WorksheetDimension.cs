namespace LightWeightExcelReader
{
    public class WorksheetDimension
    {
        public CellRef TopLeft { get; set; }
        public CellRef BottomRight { get; set; }

        public override string ToString()
        {
            return $"{TopLeft}:{BottomRight}";
        }
    }
}