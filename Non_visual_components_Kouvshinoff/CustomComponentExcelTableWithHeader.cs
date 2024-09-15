using Non_visual_components_Kouvshinoff.HelpingEnums;
using Non_visual_components_Kouvshinoff.HelpingModels;
using Non_visual_components_Kouvshinoff.InfoModels;
using System.ComponentModel;
using System.Reflection;

namespace Non_visual_components_Kouvshinoff
{
    public partial class CustomComponentExcelTableWithHeader : Component
    {
        public CustomComponentExcelTableWithHeader()
        {
            InitializeComponent();
        }

        public CustomComponentExcelTableWithHeader(IContainer container)
        {
            container.Add(this);
            InitializeComponent();
        }

        List<List<Tuple<ColumnInfo, int>>> header;
        List<ColumnInfo> lastHeaderRow;
        private void MakeHeader(List<ColumnInfo> headerColumns, int height = 0, int lb = 0)
        {
            foreach (var column in headerColumns)
            {
                header[height].Add(new Tuple<ColumnInfo, int>(column,lb));
                if (column.MergeColumns)
                    MakeHeader(column.MergedColumns!, height+1, lb);
                else
                    lastHeaderRow.Add(column);
                lb++;
            }
        }
        /// <summary>
        /// создаёт в xlsx документе таблицу, шапка таблици заполняется по информации из headerColumns. Столбцы могут объединять другие столбцы, в этом случае подпись объединяющего столбца будет над объединяемыми. 
        /// </summary>
        /// <typeparam name="T">Тип данных</typeparam>
        /// <param name="fileName">полный путь к файлу с названием</param>
        /// <param name="title">заголовок таблицы</param>
        /// <param name="headerColumns">Шапка таблици</param>
        /// <param name="lines">данные таблицы</param>
        /// <exception cref="ArgumentNullException">если что то из входных данных отсутствует</exception>
        /// <exception cref="NullReferenceException">если в объекте нет поля который описан для колонки</exception>
        public void createExcel<T>(string fileName, string title, List<ColumnInfo> headerColumns, List<T> lines)
        {
            if (string.IsNullOrEmpty(fileName))
                throw new ArgumentNullException("fileName");
            if (string.IsNullOrEmpty(title))
                throw new ArgumentNullException("title");
            if (lines == null)
                throw new ArgumentNullException("lines");

            ExcelTable table = new ExcelTable();
            table.CreateExcel(fileName);

            table.InsertCellInWorksheet(new CellCoords { RowIndex = 1U, ColumnName = "A" }, title, ExcelStyleInfoType.Title);
            int tableRow = 2;
            int headerHeight = headerColumns.Max(x => x.deep) + 1;
            header = new List<List<Tuple<ColumnInfo, int>>>(headerHeight);
            for (int i = 0; i < headerHeight; i++)
            {
                header.Add(new List<Tuple<ColumnInfo, int>>());
            }
            lastHeaderRow = new List<ColumnInfo>();
            MakeHeader(headerColumns);
            foreach (var row in header)
            {
                for (int i = 0; i < row.Count; i++)
                {
                    var cell = row[i];

                    CellCoords leftTop = new CellCoords { RowIndex = (uint)tableRow, ColumnName = HelpingFunctions.ColumnIndexToLetter(cell.Item2) };
                    CellCoords rightBotom = new CellCoords { RowIndex = (uint)(headerHeight - cell.Item1.deep) + 1U, ColumnName = HelpingFunctions.ColumnIndexToLetter(cell.Item2 + cell.Item1.width - 1) };
                    for (int x = cell.Item2; x <= cell.Item2 + cell.Item1.width - 1; x++)
                        for (int y = tableRow; y <= headerHeight - cell.Item1.deep + 1; y++)
                            table.InsertCellInWorksheet(new CellCoords { RowIndex = (uint)y, ColumnName = HelpingFunctions.ColumnIndexToLetter(x) }, cell.Item1.ColumnName, ExcelStyleInfoType.TextWithBorder);
                    if (leftTop.RowIndex != rightBotom.RowIndex || leftTop.ColumnName != rightBotom.ColumnName)
                        table.MergeCells(leftTop, rightBotom);
                }
                tableRow++;
            }

            for (int i = 0; i < lastHeaderRow.Count; i++)
                table.SetColumnWidth((uint)i+1, lastHeaderRow[i].Width);

            foreach (var line in lines)
            {
                for (int i = 0; i < lastHeaderRow.Count; i++)
                {
                    var col = lastHeaderRow[i];
                    FieldInfo? fieldInfo = typeof(T).GetField(col.FieldName!);
                    if (fieldInfo == null)
                        throw new NullReferenceException(col.FieldName!);
                    table.InsertCellInWorksheet(new CellCoords { RowIndex = (uint)tableRow, ColumnName = HelpingFunctions.ColumnIndexToLetter(i) }, fieldInfo.GetValue(line).ToString(), ExcelStyleInfoType.TextWithBorder);
                }
                tableRow++;
            }

            table.SaveExcel();
        }
    }
}
