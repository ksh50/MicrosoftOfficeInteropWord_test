using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Drawing.Spreadsheet;
using DocumentFormat.OpenXml.Drawing;

using OpenXmlExcel;

GenChart.CreateExcelWithChart();

namespace OpenXmlExcel
{
    class GenChart
    {
        public static void CreateExcelWithChart()
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Create("Sample.xlsx", SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet(new SheetData());

                Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());
                Sheet sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Sheet 1" };
                sheets.Append(sheet);

                WorkbookStylesPart stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
                stylesPart.Stylesheet = new Stylesheet();
                stylesPart.Stylesheet.Save();

                // Add data to Excel
                SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
                string[] months = { "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec" };
                double[] temperatures = { 5.6, 7.2, 10.1, 13.4, 17.2, 20.1, 22.3, 22.1, 18.9, 14.2, 9.1, 6.2 }; // Replace with actual data

                for (int i = 0; i < 12; i++)
                {
                    Row row = new Row();
                    Cell monthCell = new Cell() { DataType = CellValues.String, CellValue = new CellValue(months[i]) };
                    Cell tempCell = new Cell() { DataType = CellValues.Number, CellValue = new CellValue(temperatures[i].ToString()) };
                    row.Append(monthCell);
                    row.Append(tempCell);
                    sheetData.Append(row);
                }

                // Add a new chart to Excel
                DrawingsPart drawingsPart = worksheetPart.AddNewPart<DrawingsPart>();
                ChartPart chartPart = drawingsPart.AddNewPart<ChartPart>();
                chartPart.ChartSpace = new ChartSpace();
                chartPart.ChartSpace.AppendChild(new EditingLanguage() { Val = "en-US" });
                DocumentFormat.OpenXml.Drawing.Charts.Chart chart = chartPart.ChartSpace.AppendChild(new DocumentFormat.OpenXml.Drawing.Charts.Chart());

                // Define the 2D line chart
                PlotArea plotArea = chart.AppendChild(new PlotArea());
                LineChart lineChart = plotArea.AppendChild(new LineChart(new Grouping() { Val = GroupingValues.Standard }));

                // Define the category axis
                CategoryAxisData categoryAxisData = new CategoryAxisData();
                StringReference stringReference = categoryAxisData.AppendChild(new StringReference());
                stringReference.AppendChild(new DocumentFormat.OpenXml.Drawing.Charts.Formula($"Sheet1!$A$2:$A$13"));
                lineChart.AppendChild(categoryAxisData);

                // Define the value axis
                DocumentFormat.OpenXml.Drawing.Charts.Values values = lineChart.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.Values>();
                if (values == null)
                {
                    values = lineChart.AppendChild(new DocumentFormat.OpenXml.Drawing.Charts.Values());
                }
                NumberReference numberReference = values.GetFirstChild<NumberReference>();
                if (numberReference == null)
                {
                    numberReference = values.AppendChild(new NumberReference());
                }
                numberReference.AppendChild(new DocumentFormat.OpenXml.Drawing.Charts.Formula($"Sheet1!$B$2:$B$13"));

                // Save the chart part
                chartPart.ChartSpace.Save();

                // Add the chart to the worksheet
                DrawingsPart newDrawingsPart = worksheetPart.AddNewPart<DrawingsPart>();
                newDrawingsPart.WorksheetDrawing = new WorksheetDrawing();
                DocumentFormat.OpenXml.Drawing.Spreadsheet.GraphicFrame graphicFrame = newDrawingsPart.WorksheetDrawing.AppendChild(new DocumentFormat.OpenXml.Drawing.Spreadsheet.GraphicFrame(
                    new Transform(new Offset(), new Extents()),
                    new Graphic(new GraphicData(chartPart.GetIdOfPart(chartPart)))));
                newDrawingsPart.WorksheetDrawing.Save();

                workbookPart.Workbook.Save();
            }

        }
    }

}

