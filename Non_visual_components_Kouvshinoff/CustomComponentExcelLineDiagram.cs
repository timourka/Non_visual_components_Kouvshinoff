using Non_visual_components_Kouvshinoff.Enums;
using Non_visual_components_Kouvshinoff.HelpingEnums;
using Non_visual_components_Kouvshinoff.HelpingModels;
using System.ComponentModel;

namespace Non_visual_components_Kouvshinoff
{
    public partial class CustomComponentExcelLineDiagram : Component
    {
        public CustomComponentExcelLineDiagram()
        {
            InitializeComponent();
        }

        public CustomComponentExcelLineDiagram(IContainer container)
        {
            container.Add(this);

            InitializeComponent();
        }

        /// <summary>
        /// создать excel документ с линейным графиком 
        /// </summary>
        /// <param name="fileName">имя файла (включая путь до файла)</param>
        /// <param name="title">заголовок в документе</param>
        /// <param name="diagramTitle">заголовок для диаграммы</param>
        /// <param name="diagramLegendLocation">расположения легенды для диаграммы</param>
        /// <param name="header">шапка таблици с данными, или иначе говоря название точек в диаграмме</param>
        /// <param name="ranges">диапазоны для диаграммы</param>
        /// <exception cref="ArgumentNullException">если что то не заполненно</exception>
        public void createExcel(string fileName, string title, string diagramTitle, DiagramLegendLocation diagramLegendLocation, List<string> header, List<InfoModels.Range> ranges)
        {
            if (string.IsNullOrEmpty(fileName))
                throw new ArgumentNullException("fileName");
            if (string.IsNullOrEmpty(title))
                throw new ArgumentNullException("title");
            if (string.IsNullOrEmpty(diagramTitle))
                throw new ArgumentNullException("deiagramTitle");
            if (header == null)
                throw new ArgumentNullException("header");
            if (ranges == null)
                throw new ArgumentNullException("ranges");

            ExcelTable table = new ExcelTable();
            table.CreateExcel(fileName);

            table.InsertCellInWorksheet(new CellCoords { RowIndex = 1U, ColumnName = "A" }, title, ExcelStyleInfoType.Title);

            table.AddChart(diagramTitle, diagramLegendLocation, header, ranges);

            table.SaveExcel();
        }
    }
}
