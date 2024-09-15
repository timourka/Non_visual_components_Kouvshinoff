namespace Non_visual_components_Kouvshinoff
{
    internal static class HelpingFunctions
    {
        internal static string ColumnIndexToLetter(int columnIndex)
        {
            columnIndex++;
            string columnLetter = string.Empty;
            while (columnIndex > 0)
            {
                int remainder = (columnIndex - 1) % 26;
                columnLetter = (char)(remainder + 'A') + columnLetter;
                columnIndex = (columnIndex - 1) / 26;
            }
            return columnLetter;
        }
    }
}
