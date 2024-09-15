using Non_visual_components_Kouvshinoff.HelpingEnums;
using Non_visual_components_Kouvshinoff.HelpingModels;
using System.ComponentModel;

namespace Non_visual_components_Kouvshinoff
{
    public partial class CustomComponentExcelBigText : Component
    {
        public CustomComponentExcelBigText()
        {
            InitializeComponent();
        }

        public CustomComponentExcelBigText(IContainer container)
        {
            container.Add(this);

            InitializeComponent();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="fileName">имя файла (включая путь до файла)</param>
        /// <param name="title">название документа (заголовок в документе)</param>
        /// <param name="lines">массив строк (каждая строка – текст в ячейке для табличного документа)</param>
        /// <exception cref="ArgumentNullException">если какие то входные данные отсутствуют</exception>
        public void createExcel(string fileName, string title, string[] lines)
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
            for (int i = 0; i < lines.Length; i++)
            {
                table.InsertCellInWorksheet(new CellCoords { RowIndex = (uint)i + 3U, ColumnName = "A" }, lines[i], ExcelStyleInfoType.Text);
            }

        }
    }
}
