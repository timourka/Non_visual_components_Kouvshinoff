﻿using DocumentFormat.OpenXml.Office2010.Excel;
using DocumentFormat.OpenXml.Office2013.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using System.ComponentModel;
using DocumentFormat.OpenXml.EMMA;


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

        private SpreadsheetDocument? _spreadsheetDocument;
        private SharedStringTablePart? _shareStringPart;
        private Worksheet? _worksheet;

        /// <summary>
        /// Настройка стилей для файла
        /// </summary>
        /// <param name="workbookpart"></param>
        private static void CreateStyles(WorkbookPart workbookpart)
        {
            var sp = workbookpart.AddNewPart<WorkbookStylesPart>();
            sp.Stylesheet = new Stylesheet();
            var fonts = new Fonts() { Count = 2U, KnownFonts = true };
            var fontUsual = new DocumentFormat.OpenXml.Spreadsheet.Font();
            fontUsual.Append(new FontSize() { Val = 12D });
            fontUsual.Append(new DocumentFormat.OpenXml.Office2010.Excel.Color() { Theme = 1U });
            fontUsual.Append(new FontName() { Val = "Times New Roman" });
            fontUsual.Append(new FontFamilyNumbering() { Val = 2 });
            fontUsual.Append(new FontScheme() { Val = FontSchemeValues.Minor });
            var fontTitle = new DocumentFormat.OpenXml.Spreadsheet.Font();
            fontTitle.Append(new Bold());
            fontTitle.Append(new FontSize() { Val = 14D });
            fontTitle.Append(new DocumentFormat.OpenXml.Office2010.Excel.Color() { Theme = 1U });
            fontTitle.Append(new FontName() { Val = "Times New Roman" });
            fontTitle.Append(new FontFamilyNumbering() { Val = 2 });
            fontTitle.Append(new FontScheme() { Val = FontSchemeValues.Minor });
            fonts.Append(fontUsual);
            fonts.Append(fontTitle);
            var fills = new Fills() { Count = 2U };
            var fill1 = new Fill();
            fill1.Append(new PatternFill() { PatternType = PatternValues.None });
            var fill2 = new Fill();
            fill2.Append(new PatternFill() { PatternType = PatternValues.Gray125 });
            fills.Append(fill1);
            fills.Append(fill2);
            var borders = new Borders() { Count = 1U };
            var borderNoBorder = new Border();
            borderNoBorder.Append(new LeftBorder());
            borderNoBorder.Append(new RightBorder());
            borderNoBorder.Append(new TopBorder());
            borderNoBorder.Append(new BottomBorder());
            borderNoBorder.Append(new DiagonalBorder());
            borders.Append(borderNoBorder);
            var cellStyleFormats = new CellStyleFormats() { Count = 1U };
            var cellFormatStyle = new CellFormat()
            {
                NumberFormatId = 0U,
                FontId = 0U,
                FillId = 0U,
                BorderId = 0U
            };
            cellStyleFormats.Append(cellFormatStyle);
            var cellFormats = new CellFormats() { Count = 2U };
            var cellFormatFont = new CellFormat()
            {
                NumberFormatId = 0U,
                FontId = 0U,
                FillId = 0U,
                BorderId = 0U,
                FormatId = 0U,
                Alignment = new Alignment()
                {
                    Vertical = VerticalAlignmentValues.Center,
                    WrapText = false,
                    Horizontal = HorizontalAlignmentValues.Left
                },
                ApplyFont = true
            };
            var cellFormatTitle = new CellFormat()
            {
                NumberFormatId = 0U,
                FontId = 1U,
                FillId = 0U,
                BorderId = 0U,
                FormatId = 0U,
                Alignment = new Alignment()
                {
                    Vertical = VerticalAlignmentValues.Center,
                    WrapText = false,
                    Horizontal = HorizontalAlignmentValues.Left
                },
                ApplyFont = true
            };
            cellFormats.Append(cellFormatFont);
            cellFormats.Append(cellFormatTitle);
            var cellStyles = new CellStyles() { Count = 1U };
            cellStyles.Append(new CellStyle()
            {
                Name = "Normal",
                FormatId = 0U,
                BuiltinId = 0U
            });
            var differentialFormats = new DocumentFormat.OpenXml.Office2013.Excel.DifferentialFormats() { Count = 0U };
            var tableStyles = new TableStyles()
            {
                Count = 0U,
                DefaultTableStyle = "TableStyleMedium2",
                DefaultPivotStyle = "PivotStyleLight16"
            };
            var stylesheetExtensionList = new StylesheetExtensionList();
            var stylesheetExtension1 = new StylesheetExtension() { Uri = "{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}" };
            stylesheetExtension1.AddNamespaceDeclaration("x14", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");
            stylesheetExtension1.Append(new SlicerStyles() { DefaultSlicerStyle = "SlicerStyleLight1" });
            var stylesheetExtension2 = new StylesheetExtension() { Uri = "{9260A510-F301-46a8-8635-F512D64BE5F5}" };
            stylesheetExtension2.AddNamespaceDeclaration("x15", "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main");
            stylesheetExtension2.Append(new TimelineStyles() { DefaultTimelineStyle = "TimeSlicerStyleLight1" });
            stylesheetExtensionList.Append(stylesheetExtension1);
            stylesheetExtensionList.Append(stylesheetExtension2);
            sp.Stylesheet.Append(fonts);
            sp.Stylesheet.Append(fills);
            sp.Stylesheet.Append(borders);
            sp.Stylesheet.Append(cellStyleFormats);
            sp.Stylesheet.Append(cellFormats);
            sp.Stylesheet.Append(cellStyles);
            sp.Stylesheet.Append(differentialFormats);
            sp.Stylesheet.Append(tableStyles);
            sp.Stylesheet.Append(stylesheetExtensionList);
        }

        private void CreateExcel(string fileName)
        {
            _spreadsheetDocument = SpreadsheetDocument.Create(fileName, SpreadsheetDocumentType.Workbook); ///
            // Создаем книгу (в ней хранятся листы)
            var workbookpart = _spreadsheetDocument.AddWorkbookPart();
            workbookpart.Workbook = new Workbook();
            CreateStyles(workbookpart);
            // Получаем/создаем хранилище текстов для книги
            _shareStringPart = _spreadsheetDocument.WorkbookPart!.GetPartsOfType<SharedStringTablePart>().Any()
                ?
                _spreadsheetDocument.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First()
                :
                _spreadsheetDocument.WorkbookPart.AddNewPart<SharedStringTablePart>();
            // Создаем SharedStringTable, если его нет
            if (_shareStringPart.SharedStringTable == null)
            {
                _shareStringPart.SharedStringTable = new SharedStringTable();
            }
            // Создаем лист в книгу
            var worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());
            // Добавляем лист в книгу
            var sheets = _spreadsheetDocument.WorkbookPart.Workbook.AppendChild(new Sheets());
            var sheet = new Sheet()
            {
                Id = _spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart),
                SheetId = 1,
                Name = "Лист"
            };
            sheets.Append(sheet);
            _worksheet = worksheetPart.Worksheet;
        }
        private void InsertCellInWorksheet(uint rowIndex, string columnName, string text, bool title)
        {
            string cellReference = $"{columnName}{rowIndex}";
            if (_worksheet == null || _shareStringPart == null)
            {
                return;
            }
            var sheetData = _worksheet.GetFirstChild<SheetData>();
            if (sheetData == null)
            {
                return;
            }
            // Ищем строку, либо добавляем ее
            Row row;
            if (sheetData.Elements<Row>().Where(r => r.RowIndex! == rowIndex).Any())
            {
                row = sheetData.Elements<Row>().Where(r => r.RowIndex! == rowIndex).First();
            }
            else
            {
                row = new Row() { RowIndex = rowIndex };
                sheetData.Append(row);
            }
            // Ищем нужную ячейку
            Cell cell;
            if (row.Elements<Cell>().Where(c => c.CellReference!.Value == cellReference).Any())
            {
                cell = row.Elements<Cell>().Where(c => c.CellReference!.Value == cellReference).First();
            }
            else
            {
                // Все ячейки должны быть последовательно друг за другом расположены
                // нужно определить, после какой вставлять
                Cell? refCell = null;
                foreach (Cell rowCell in row.Elements<Cell>())
                {
                    if (string.Compare(rowCell.CellReference!.Value, cellReference, true) > 0)
                    {
                        refCell = rowCell;
                        break;
                    }
                }
                var newCell = new Cell()
                {
                    CellReference = cellReference
                };
                row.InsertBefore(newCell, refCell);
                cell = newCell;
            }
            // вставляем новый текст
            _shareStringPart.SharedStringTable.AppendChild(new SharedStringItem(new Text(text)));
            _shareStringPart.SharedStringTable.Save();
            cell.CellValue = new CellValue((_shareStringPart.SharedStringTable.Elements<SharedStringItem>().Count() - 1).ToString());
            cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
            if (title)
                cell.StyleIndex = 1U;
            else
                cell.StyleIndex = 0U;
        }

        private void SaveExcel()
        {
            if (_spreadsheetDocument == null)
            {
                return;
            }
            _spreadsheetDocument.WorkbookPart!.Workbook.Save();
            _spreadsheetDocument.Close();
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

            CreateExcel(fileName);

            InsertCellInWorksheet(1, "A", title, true);
            for (int i = 0; i < lines.Length; i++)
            {
                InsertCellInWorksheet((uint)i + 3U, "A", lines[i], false);
            }

            SaveExcel();
        }
    }
}
