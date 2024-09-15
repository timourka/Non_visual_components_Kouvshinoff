using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Drawing.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Non_visual_components_Kouvshinoff.Enums;
using Index = DocumentFormat.OpenXml.Drawing.Charts.Index;
using Orientation = DocumentFormat.OpenXml.Drawing.Charts.Orientation;
using Run = DocumentFormat.OpenXml.Drawing.Run;
using TextProperties = DocumentFormat.OpenXml.Drawing.Charts.TextProperties;


namespace Non_visual_components_Kouvshinoff
{
    internal class aboba
    {
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
        public void CreateExcelDoc(string fileName, string chartTitle, DiagramLegendLocation diagramLegendLocation, List<string> header, List<InfoModels.Range> ranges)
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Create(fileName, SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet();

                Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());

                Sheet sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Students" };

                SheetData sheetData = worksheetPart.Worksheet.AppendChild(new SheetData());

                // Добавляем DrawingsPart и диаграмму
                DrawingsPart drawingsPart = worksheetPart.AddNewPart<DrawingsPart>();
                worksheetPart.Worksheet.Append(new Drawing() { Id = worksheetPart.GetIdOfPart(drawingsPart) });
                worksheetPart.Worksheet.Save();

                drawingsPart.WorksheetDrawing = new WorksheetDrawing();
                sheets.Append(sheet);
                workbookPart.Workbook.Save();

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
                int rowIndex = 1;
                row.AppendChild(ConstructCell(string.Empty, CellValues.String));

                foreach (var h in header)
                {
                    row.AppendChild(ConstructCell(h, CellValues.String));
                }

                sheetData.AppendChild(row);
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
                    string formulaCat = $"Students!$B$1:${HelpingFunctions.ColumnIndexToLetter(header.Count+1)}$1";
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
                    row.AppendChild(ConstructCell(ranges[i].name, CellValues.String));
                    chartSeries.MoveNext();

                    string formulaVal = string.Format("Students!$B${0}:$G${0}", rowIndex);
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
                        row.AppendChild(ConstructCell(value, CellValues.Number));
                        numberingCache.AppendChild(new NumericPoint() { Index = (uint)j }).Append(new NumericValue(value));
                    }

                    sheetData.AppendChild(row);
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
                    new ColumnId("8"),
                    new ColumnOffset("0"),
                    new RowId((rowIndex + 12).ToString()),
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
        }

        private Cell ConstructCell(string value, CellValues dataType)
        {
            return new Cell()
            {
                CellValue = new CellValue(value),
                DataType = new EnumValue<CellValues>(dataType),
            };
        }
    }
}
