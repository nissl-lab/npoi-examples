using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XDDF.UserModel.Chart;
using NPOI.XSSF.UserModel;
using System;
using System.IO;

namespace BarChart
{
    class Program
    {
        const int NUM_OF_ROWS = 10;
        const int NUM_OF_COLUMNS = 2;
        private static void CreateChart(ISheet sheet, XSSFDrawing drawing, IClientAnchor anchor,  string serieTitle, int startDataRow, int endDataRow, int columnIndex)
        {
            var chart = drawing.CreateChart(anchor);
            var legend = chart.GetOrAddLegend();
            legend.Position = LegendPosition.Bottom;

            var bottomAxis = chart.CreateCategoryAxis(AxisPosition.Bottom);
            bottomAxis.MajorTickMark = AxisTickMark.None;
            var leftAxis = chart.CreateValueAxis(AxisPosition.Left);
            leftAxis.Crosses = AxisCrosses.AutoZero;
            leftAxis.CrossBetween =AxisCrossBetween.Between;

            var barChartData = chart.CreateData<string, double>(ChartTypes.BAR, bottomAxis, leftAxis);

            var categoryAxis = XDDFDataSourcesFactory.FromStringCellRange(sheet, new CellRangeAddress(startDataRow, endDataRow, 0, 0));
            var valueAxis = XDDFDataSourcesFactory.FromNumericCellRange(sheet, new CellRangeAddress(startDataRow, endDataRow, columnIndex, columnIndex));
            var serie = barChartData.AddSeries(categoryAxis, valueAxis);
            serie.SetTitle(serieTitle);

            chart.Plot(barChartData);
        }
        static void Main(string[] args)
        {
            using (IWorkbook wb = new XSSFWorkbook())
            {
                ISheet sheet = wb.CreateSheet();


                // Create a row and put some cells in it. Rows are 0 based.
                IRow row;
                ICell cell;
                for (int rowIndex = 0; rowIndex < NUM_OF_ROWS; rowIndex++)
                {
                    row = sheet.CreateRow((short)rowIndex);
                    for (int colIndex = 0; colIndex < NUM_OF_COLUMNS; colIndex++)
                    {
                        cell = row.CreateCell((short)colIndex);
                        if (colIndex == 0)
                            cell.SetCellValue("X" + rowIndex);
                        else
                        {
                            var x = colIndex * (rowIndex + 1);
                            cell.SetCellValue(x * x + 2 * x + 1);
                        }
                    }
                }
                XSSFDrawing drawing = (XSSFDrawing)sheet.CreateDrawingPatriarch();
                XSSFClientAnchor anchor = (XSSFClientAnchor)drawing.CreateAnchor(0, 0, 0, 0, 3, 3, 10, 12);

                CreateChart(sheet, drawing, anchor, "s1", 0, 9, 1);
                using (FileStream fs = File.Create("test.xlsx"))
                {
                    wb.Write(fs, false);
                }
            }
        }
    }
}
