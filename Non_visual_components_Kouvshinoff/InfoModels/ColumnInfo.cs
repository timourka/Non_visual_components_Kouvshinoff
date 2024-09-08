namespace Non_visual_components_Kouvshinoff.InfoModels
{
    public class ColumnInfo
    {
        /// <summary>
        /// название столбца
        /// </summary>
        public string ColumnName { get; }
        /// <summary>
        /// объединаяет ли какието столбцы
        /// </summary>
        public bool MergeColumns { get; }
        /// <summary>
        /// объеденёные столбцы
        /// </summary>
        public List<ColumnInfo>? MergedColumns { get; }
        /// <summary>
        /// ширина столбца
        /// </summary>
        public double Width { get; }
        /// <summary>
        /// название поля объекта для этого столбца
        /// </summary>
        public string? FieldName { get; }
        /// <summary>
        /// глубина дерева образуемого колонкой
        /// </summary>
        internal int deep {  get; }
        /// <summary>
        /// ширина дерева образуемого колонкой
        /// </summary>
        internal int width { get; }
        /// <summary>
        /// Создать столбец который будет объединять в себя другие столбцы, через поле mergedCollumns можно передавать вложеные столбцы
        /// </summary>
        /// <param name="columnName">название столбца в шапке</param>
        /// <param name="mergedColumns">объединяемые столбцы</param>
        public ColumnInfo(string columnName, List<ColumnInfo> mergedColumns)
        {
            ColumnName = columnName;
            MergeColumns = true;
            MergedColumns = mergedColumns;
            Width = mergedColumns.Sum(x => x.Width);
            width = mergedColumns.Sum(x => x.width);
            deep = mergedColumns.Max(x => x.deep) + 1;
        }
        /// <summary>
        /// Создать столбец который не объединяет никаких других столбцов
        /// </summary>
        /// <param name="columnName">название столбца в шапке</param>
        /// <param name="width">ширина столбца</param>
        /// <param name="fieldName">название поля объекта для этого столбца</param>
        public ColumnInfo(string columnName, double width, string fieldName)
        {
            ColumnName = columnName;
            MergeColumns = false;
            Width = width;
            FieldName = fieldName;
            deep = 0;
            this.width = 1;
        }
    }
}
