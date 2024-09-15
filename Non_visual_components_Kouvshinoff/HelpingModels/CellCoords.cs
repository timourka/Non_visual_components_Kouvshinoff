namespace Non_visual_components_Kouvshinoff.HelpingModels
{
    internal class CellCoords
    {
        public string ColumnName { get; set; } = string.Empty;
        public uint RowIndex { get; set; }
        public string CellReference => $"{ColumnName}{RowIndex}";

        public CellCoords() { }
        public CellCoords(uint rowIndex, string columnName)
        {
            ColumnName = columnName;
            RowIndex = rowIndex;
        }
    }
}
