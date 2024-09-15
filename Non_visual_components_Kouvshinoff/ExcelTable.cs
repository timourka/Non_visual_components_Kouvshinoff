using DocumentFormat.OpenXml.Office2010.Excel;
using DocumentFormat.OpenXml.Office2013.Excel;

using Non_visual_components_Kouvshinoff.Enums;
using Non_visual_components_Kouvshinoff.HelpingModels;
using Non_visual_components_Kouvshinoff.HelpingEnums;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Drawing.Spreadsheet;

using TextProperties = DocumentFormat.OpenXml.Drawing.Charts.TextProperties;
using Index = DocumentFormat.OpenXml.Drawing.Charts.Index;
using Orientation = DocumentFormat.OpenXml.Drawing.Charts.Orientation;
using DisplayBlanksAsValues = DocumentFormat.OpenXml.Drawing.Charts.DisplayBlanksAsValues;

using Run = DocumentFormat.OpenXml.Drawing.Run;
using ParagraphProperties = DocumentFormat.OpenXml.Drawing.ParagraphProperties;
using DefaultRunProperties = DocumentFormat.OpenXml.Drawing.DefaultRunProperties;
using Paragraph = DocumentFormat.OpenXml.Drawing.Paragraph;
using BodyProperties = DocumentFormat.OpenXml.Drawing.BodyProperties;
using ListStyle = DocumentFormat.OpenXml.Drawing.ListStyle;
using EndParagraphRunProperties = DocumentFormat.OpenXml.Drawing.EndParagraphRunProperties;

namespace Non_visual_components_Kouvshinoff
{
    internal class ExcelTable
    {
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
            var borders = new Borders() { Count = 2U };
            var borderNoBorder = new Border();
            borderNoBorder.Append(new LeftBorder());
            borderNoBorder.Append(new RightBorder());
            borderNoBorder.Append(new TopBorder());
            borderNoBorder.Append(new BottomBorder());
            borderNoBorder.Append(new DiagonalBorder());
            var borderThin = new Border();
            var leftBorder = new LeftBorder() { Style = BorderStyleValues.Thin };
            leftBorder.Append(new DocumentFormat.OpenXml.Office2010.Excel.Color() { Indexed = 64U });
            var rightBorder = new RightBorder() { Style = BorderStyleValues.Thin };
            rightBorder.Append(new DocumentFormat.OpenXml.Office2010.Excel.Color() { Indexed = 64U });
            var topBorder = new TopBorder() { Style = BorderStyleValues.Thin };
            topBorder.Append(new DocumentFormat.OpenXml.Office2010.Excel.Color() { Indexed = 64U });
            var bottomBorder = new BottomBorder() { Style = BorderStyleValues.Thin };
            bottomBorder.Append(new DocumentFormat.OpenXml.Office2010.Excel.Color() { Indexed = 64U });
            borderThin.Append(leftBorder);
            borderThin.Append(rightBorder);
            borderThin.Append(topBorder);
            borderThin.Append(bottomBorder);
            borderThin.Append(new DiagonalBorder());
            borders.Append(borderNoBorder);
            borders.Append(borderThin);
            var cellStyleFormats = new CellStyleFormats() { Count = 1U };
            var cellFormatStyle = new CellFormat()
            {
                NumberFormatId = 0U,
                FontId = 0U,
                FillId = 0U,
                BorderId = 0U
            };
            cellStyleFormats.Append(cellFormatStyle);
            var cellFormats = new CellFormats() { Count = 3U };
            var cellFormatFont = new CellFormat()
            {
                NumberFormatId = 0U,
                FontId = 0U,
                FillId = 0U,
                BorderId = 0U,
                FormatId = 0U,
                ApplyFont = true
            };
            var cellFormatFontAndBorder = new CellFormat()
            {
                NumberFormatId = 0U,
                FontId = 0U,
                FillId = 0U,
                BorderId = 1U,
                FormatId = 0U,
                Alignment = new Alignment()
                {
                    Vertical = VerticalAlignmentValues.Center,
                    WrapText = true,
                    Horizontal = HorizontalAlignmentValues.Center
                },
                ApplyFont = true,
                ApplyBorder = true
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
                    WrapText = true,
                    Horizontal = HorizontalAlignmentValues.Center
                },
                ApplyFont = true
            };
            cellFormats.Append(cellFormatFont);
            cellFormats.Append(cellFormatFontAndBorder);
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
        internal void CreateExcel(string fileName)
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
        /// <summary>
        /// Получение номера стиля из типа
        /// </summary>
        /// <param name="styleInfo"></param>
        /// <returns></returns>
        private static uint GetStyleValue(ExcelStyleInfoType styleInfo)
        {
            return styleInfo switch
            {
                ExcelStyleInfoType.Title => 2U,
                ExcelStyleInfoType.TextWithBorder => 1U,
                ExcelStyleInfoType.Text => 0U,
                _ => 0U,
            };
        }
        internal void InsertCellInWorksheet(CellCoords cellCoords, string text, ExcelStyleInfoType cellType)
        {
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
            if (sheetData.Elements<Row>().Where(r => r.RowIndex == cellCoords.RowIndex).Any())
            {
                row = sheetData.Elements<Row>().Where(r => r.RowIndex == cellCoords.RowIndex).First();
            }
            else
            {
                row = new Row() { RowIndex = cellCoords.RowIndex };
                sheetData.Append(row);
            }

            // Ищем нужную ячейку
            Cell cell;
            if (row.Elements<Cell>().Where(c => c.CellReference == cellCoords.CellReference).Any())
            {
                cell = row.Elements<Cell>().Where(c => c.CellReference == cellCoords.CellReference).First();
            }
            else
            {
                // Все ячейки должны быть последовательно друг за другом расположены
                // нужно определить, после какой вставлять
                Cell? refCell = null;
                foreach (Cell rowCell in row.Elements<Cell>())
                {
                    if (string.Compare(rowCell.CellReference, cellCoords.CellReference, true) > 0)
                    {
                        refCell = rowCell;
                        break;
                    }
                }
                var newCell = new Cell()
                {
                    CellReference = cellCoords.CellReference
                };
                row.InsertBefore(newCell, refCell);
                cell = newCell;
            }

            // Проверяем, является ли значение числом
            if (double.TryParse(text, out double numericValue))
            {
                // Если это число, добавляем его как число
                cell.CellValue = new CellValue(numericValue.ToString(System.Globalization.CultureInfo.InvariantCulture));
                cell.DataType = new EnumValue<CellValues>(CellValues.Number);
            }
            else
            {
                // Если это текст, добавляем как shared string
                _shareStringPart.SharedStringTable.AppendChild(new SharedStringItem(new Text(text)));
                _shareStringPart.SharedStringTable.Save();
                cell.CellValue = new CellValue((_shareStringPart.SharedStringTable.Elements<SharedStringItem>().Count() - 1).ToString());
                cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
            }

            // Устанавливаем стиль ячейки
            cell.StyleIndex = GetStyleValue(cellType);
        }

        internal void MergeCells(CellCoords startCell, CellCoords endCell)
        {
            string merge = $"{startCell.CellReference}:{endCell.CellReference}";
            if (_worksheet == null)
            {
                return;
            }
            MergeCells mergeCells;
            if (_worksheet.Elements<MergeCells>().Any())
            {
                mergeCells = _worksheet.Elements<MergeCells>().First();
            }
            else
            {
                mergeCells = new MergeCells();
                if (_worksheet.Elements<CustomSheetView>().Any())
                {
                    _worksheet.InsertAfter(mergeCells, _worksheet.Elements<CustomSheetView>().First());
                }
                else
                {
                    _worksheet.InsertAfter(mergeCells, _worksheet.Elements<SheetData>().First());
                }
            }
            var mergeCell = new MergeCell()
            {
                Reference = new StringValue(merge)
            };
            mergeCells.Append(mergeCell);
        }
        internal void SetColumnWidth(uint columnIndex, double width)
        {
            if (_worksheet == null)
            {
                return;
            }

            Columns? columns;

            // Проверяем, есть ли уже секция Columns, если нет, создаем новую
            if (_worksheet.Elements<Columns>().Any())
            {
                columns = _worksheet.Elements<Columns>().First();
            }
            else
            {
                columns = new Columns();
                // Вставляем Columns перед SheetData
                _worksheet.InsertAt(columns, 0);
            }

            // Проверяем, существует ли уже колонка с заданным индексом
            Column? existingColumn = columns.Elements<Column>().FirstOrDefault(c => c.Min == columnIndex && c.Max == columnIndex);

            if (existingColumn != null)
            {
                // Если колонка уже есть, обновляем её ширину
                existingColumn.Width = width;
                existingColumn.CustomWidth = true;
            }
            else
            {
                // Если колонки нет, добавляем её с новой шириной
                Column column = new Column()
                {
                    Min = columnIndex,
                    Max = columnIndex,
                    Width = width,
                    CustomWidth = true
                };
                columns.Append(column);
            }

            _worksheet.Save();
        }

        private static Title GenerateTitle(string titleText)
        {
            Run run = new Run();
            run.Append(new OpenXmlElement[]
            {
                new DocumentFormat.OpenXml.Drawing.RunProperties
                {
                    FontSize = 1100
                }
            });
            run.Append(new OpenXmlElement[]
            {
                new DocumentFormat.OpenXml.Drawing.Text(titleText)
            });
            ParagraphProperties paragraphProperties = new ParagraphProperties();
            paragraphProperties.Append(new OpenXmlElement[]
            {
                new DefaultRunProperties
                {
                    FontSize = 1100
                }
            });
            Paragraph paragraph = new Paragraph();
            paragraph.Append(new OpenXmlElement[] { paragraphProperties });
            paragraph.Append(new OpenXmlElement[] { run });
            RichText richText = new RichText();
            richText.Append(new OpenXmlElement[]
            {
                new BodyProperties()
            });
            richText.Append(new OpenXmlElement[]
            {
                new ListStyle()
            });
            richText.Append(new OpenXmlElement[] { paragraph });
            ChartText chartText = new ChartText();
            chartText.Append(new OpenXmlElement[] { richText });
            Title title = new Title();
            title.Append(new OpenXmlElement[] { chartText });
            title.Append(new OpenXmlElement[]
            {
                new Layout()
            });
            title.Append(new OpenXmlElement[]
            {
                new Overlay
                {
                    Val = false
                }
            });
            return title;
        }
        private static Legend GenerateLegend(LegendPositionValues position)
        {
            ParagraphProperties paragraphProperties = new ParagraphProperties();
            paragraphProperties.Append(new OpenXmlElement[]
            {
                new DefaultRunProperties()
            });
            Paragraph paragraph = new Paragraph();
            paragraph.Append(new OpenXmlElement[] { paragraphProperties });
            paragraph.Append(new OpenXmlElement[]
            {
                new EndParagraphRunProperties()
            });
            TextProperties textProperties = new TextProperties();
            textProperties.Append(new OpenXmlElement[]
            {
                new BodyProperties()
            });
            textProperties.Append(new OpenXmlElement[]
            {
                new ListStyle()
            });
            textProperties.Append(new OpenXmlElement[] { paragraph });
            Legend legend = new Legend();
            legend.Append(new OpenXmlElement[]
            {
                new LegendPosition
                {
                    Val = position
                }
            });
            legend.Append(new OpenXmlElement[]
            {
                new Layout()
            });
            legend.Append(new OpenXmlElement[]
            {
                new Overlay
                {
                    Val = false
                }
            });
            legend.Append(new OpenXmlElement[] { textProperties });
            return legend;
        }
        internal void AddChart(string chartTitle, DiagramLegendLocation diagramLegendLocation, List<string> header, List<InfoModels.Range> ranges)
        {
            WorksheetPart worksheetPart = _worksheet.WorksheetPart;

            // Добавляем DrawingsPart и диаграмму
            DrawingsPart drawingsPart = worksheetPart.AddNewPart<DrawingsPart>();
            worksheetPart.Worksheet.Append(new Drawing() { Id = worksheetPart.GetIdOfPart(drawingsPart) });
            worksheetPart.Worksheet.Save();

            drawingsPart.WorksheetDrawing = new WorksheetDrawing();

            // Создаем диаграмму и задаем язык диаграммы
            ChartPart chartPart = drawingsPart.AddNewPart<ChartPart>();
            chartPart.ChartSpace = new ChartSpace();
            chartPart.ChartSpace.AppendChild(new EditingLanguage() { Val = "en-US" });
            DocumentFormat.OpenXml.Drawing.Charts.Chart chart = chartPart.ChartSpace.AppendChild(new DocumentFormat.OpenXml.Drawing.Charts.Chart());
            chart.AppendChild(new AutoTitleDeleted() { Val = true });

            // Добавляем заголовок диаграммы
            chart.Append(GenerateTitle(chartTitle));

            // Легенда для диаграммы
            switch (diagramLegendLocation)
            {
                case DiagramLegendLocation.Left:
                    chart.Append(GenerateLegend(LegendPositionValues.Left));
                    break;
                case DiagramLegendLocation.Top:
                    chart.Append(GenerateLegend(LegendPositionValues.Top));
                    break;
                case DiagramLegendLocation.Right:
                    chart.Append(GenerateLegend(LegendPositionValues.Right));
                    break;
                case DiagramLegendLocation.Bottom:
                    chart.Append(GenerateLegend(LegendPositionValues.Bottom));
                    break;
            }

            // Создаем область построения и добавляем линейную диаграмму
            PlotArea plotArea = chart.AppendChild(new PlotArea());
            Layout layout = plotArea.AppendChild(new Layout());

            // Линейная диаграмма вместо столбчатой
            LineChart lineChart = plotArea.AppendChild(new LineChart(
                new Grouping() { Val = new EnumValue<GroupingValues>(GroupingValues.Standard) },
                new VaryColors() { Val = false }
            ));

            // Заголовок
            Row row = new Row();
            uint rowIndex = 2;
            InsertCellInWorksheet(new(rowIndex, "A"), string.Empty, ExcelStyleInfoType.TextWithBorder);

            for (int i = 0; i < header.Count; i++)
            {
                InsertCellInWorksheet(new(rowIndex, HelpingFunctions.ColumnIndexToLetter(i + 1)), header[i], ExcelStyleInfoType.TextWithBorder);
            }

            rowIndex++;

            // Создание серий для линейной диаграммы
            for (int i = 0; i < ranges.Count; i++)
            {
                LineChartSeries lineChartSeries = lineChart.AppendChild(new LineChartSeries(
                    new Index() { Val = (uint)i },
                    new Order() { Val = (uint)i },
                    new SeriesText(new NumericValue() { Text = ranges[i].name })
                ));

                // Добавляем ось категорий
                CategoryAxisData categoryAxisData = lineChartSeries.AppendChild(new CategoryAxisData());

                // Категории
                string formulaCat = $"Лист!$B$2:${HelpingFunctions.ColumnIndexToLetter(header.Count)}$2";
                StringReference stringReference = categoryAxisData.AppendChild(new StringReference()
                {
                    Formula = new DocumentFormat.OpenXml.Drawing.Charts.Formula() { Text = formulaCat }
                });

                StringCache stringCache = stringReference.AppendChild(new StringCache());
                stringCache.Append(new PointCount() { Val = (uint)header.Count });

                for (int j = 0; j < header.Count; j++)
                {
                    stringCache.AppendChild(new NumericPoint() { Index = (uint)j }).Append(new NumericValue(header[j]));
                }
            }

            var chartSeries = lineChart.Elements<LineChartSeries>().GetEnumerator();

            for (int i = 0; i < ranges.Count; i++)
            {
                row = new Row();
                InsertCellInWorksheet(new(rowIndex, "A"), ranges[i].name, ExcelStyleInfoType.TextWithBorder);
                chartSeries.MoveNext();

                string formulaVal = string.Format("Лист!$B${0}:${1}${0}", rowIndex, HelpingFunctions.ColumnIndexToLetter(header.Count));
                DocumentFormat.OpenXml.Drawing.Charts.Values values = chartSeries.Current.AppendChild(new DocumentFormat.OpenXml.Drawing.Charts.Values());

                NumberReference numberReference = values.AppendChild(new NumberReference()
                {
                    Formula = new DocumentFormat.OpenXml.Drawing.Charts.Formula() { Text = formulaVal }
                });

                NumberingCache numberingCache = numberReference.AppendChild(new NumberingCache());
                numberingCache.Append(new PointCount() { Val = (uint)header.Count });

                for (int j = 0; j < header.Count; j++)
                {
                    string value = string.Empty;
                    if (ranges[i].data.ContainsKey(header[j]))
                        value = ranges[i].data[header[j]].ToString();
                    InsertCellInWorksheet(new(rowIndex, HelpingFunctions.ColumnIndexToLetter(j + 1)), value, ExcelStyleInfoType.TextWithBorder);
                    numberingCache.AppendChild(new NumericPoint() { Index = (uint)j }).Append(new NumericValue(value));
                }

                rowIndex++;
            }

            // Настройка осей и меток
            lineChart.AppendChild(new DataLabels(
                new ShowLegendKey() { Val = false },
                new ShowValue() { Val = false },
                new ShowCategoryName() { Val = false },
                new ShowSeriesName() { Val = false },
                new ShowPercent() { Val = false },
                new ShowBubbleSize() { Val = false }
            ));

            lineChart.Append(new AxisId() { Val = 48650112u });
            lineChart.Append(new AxisId() { Val = 48672768u });

            // Ось категорий
            plotArea.AppendChild(
                new CategoryAxis(
                    new AxisId() { Val = 48650112u },
                    new Scaling(new Orientation() { Val = new EnumValue<DocumentFormat.OpenXml.Drawing.Charts.OrientationValues>(DocumentFormat.OpenXml.Drawing.Charts.OrientationValues.MinMax) }),
                    new Delete() { Val = false },
                    new AxisPosition() { Val = new EnumValue<AxisPositionValues>(AxisPositionValues.Bottom) },
                    new TickLabelPosition() { Val = new EnumValue<TickLabelPositionValues>(TickLabelPositionValues.NextTo) },
                    new CrossingAxis() { Val = 48672768u },
                    new Crosses() { Val = new EnumValue<CrossesValues>(CrossesValues.AutoZero) },
                    new AutoLabeled() { Val = true },
                    new LabelAlignment() { Val = new EnumValue<LabelAlignmentValues>(LabelAlignmentValues.Center) }
                ));

            // Ось значений
            plotArea.AppendChild(
                new ValueAxis(
                    new AxisId() { Val = 48672768u },
                    new Scaling(new Orientation() { Val = new EnumValue<DocumentFormat.OpenXml.Drawing.Charts.OrientationValues>(DocumentFormat.OpenXml.Drawing.Charts.OrientationValues.MinMax) }),
                    new Delete() { Val = false },
                    new AxisPosition() { Val = new EnumValue<AxisPositionValues>(AxisPositionValues.Left) },
                    new MajorGridlines(),
                    new DocumentFormat.OpenXml.Drawing.Charts.NumberingFormat() { FormatCode = "General", SourceLinked = true },
                    new TickLabelPosition() { Val = new EnumValue<TickLabelPositionValues>(TickLabelPositionValues.NextTo) },
                    new CrossingAxis() { Val = 48650112u },
                    new Crosses() { Val = new EnumValue<CrossesValues>(CrossesValues.AutoZero) },
                    new CrossBetween() { Val = new EnumValue<CrossBetweenValues>(CrossBetweenValues.Between) }
                ));

            chart.Append(
                new PlotVisibleOnly() { Val = true },
                new DisplayBlanksAs() { Val = new EnumValue<DisplayBlanksAsValues>(DisplayBlanksAsValues.Gap) },
                new ShowDataLabelsOverMaximum() { Val = false }
            );

            chartPart.ChartSpace.Save();

            // Размещение диаграммы на листе
            TwoCellAnchor twoCellAnchor = drawingsPart.WorksheetDrawing.AppendChild(new TwoCellAnchor());

            twoCellAnchor.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.FromMarker(
                new ColumnId("0"),
                new ColumnOffset("0"),
                new RowId((rowIndex + 2).ToString()),
                new RowOffset("0")
            ));

            twoCellAnchor.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.ToMarker(
                new ColumnId($"{header.Count+2}"),
                new ColumnOffset("0"),
                new RowId((rowIndex + ranges.Count * 6).ToString()),
                new RowOffset("0")
            ));

            // Добавляем GraphicFrame для линейной диаграммы
            DocumentFormat.OpenXml.Drawing.Spreadsheet.GraphicFrame graphicFrame = twoCellAnchor.AppendChild(new DocumentFormat.OpenXml.Drawing.Spreadsheet.GraphicFrame());
            graphicFrame.Macro = string.Empty;

            graphicFrame.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualGraphicFrameProperties(
                new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualDrawingProperties() { Id = 2u, Name = "Line Chart" },
                new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualGraphicFrameDrawingProperties()
            ));

            graphicFrame.Append(new Transform(
                new DocumentFormat.OpenXml.Drawing.Offset() { X = 0L, Y = 0L },
                new DocumentFormat.OpenXml.Drawing.Extents() { Cx = 0L, Cy = 0L }
            ));

            graphicFrame.Append(new DocumentFormat.OpenXml.Drawing.Graphic(
                new DocumentFormat.OpenXml.Drawing.GraphicData(
                    new ChartReference() { Id = drawingsPart.GetIdOfPart(chartPart) }
                )
                { Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart" }
            ));

            twoCellAnchor.Append(new ClientData());

            drawingsPart.WorksheetDrawing.Save();
            worksheetPart.Worksheet.Save();
        }

        internal void SaveExcel()
        {
            if (_spreadsheetDocument == null)
            {
                return;
            }
            _spreadsheetDocument.WorkbookPart!.Workbook.Save();
            _spreadsheetDocument.Close();
        }
    }
}
