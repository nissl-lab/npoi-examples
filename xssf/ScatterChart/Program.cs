
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XDDF.UserModel.Chart;
using NPOI.XSSF.UserModel;
using System.IO;

namespace ScatterChart
{
    class Program
    {
        static void Main(string[] args)
        {
            using (var wb = new XSSFWorkbook())
            {
                var sheet = wb.CreateSheet("Sheet 1");
                int NUM_OF_ROWS = 3;
                int NUM_OF_COLUMNS = 10;

                // Create a row and put some cells in it. Rows are 0 based.
                IRow row;
                ICell cell;
                for (int rowIndex = 0; rowIndex < NUM_OF_ROWS; rowIndex++)
                {
                    row = sheet.CreateRow((short)rowIndex);
                    for (int colIndex = 0; colIndex < NUM_OF_COLUMNS; colIndex++)
                    {
                        cell = row.CreateCell((short)colIndex);
                        cell.SetCellValue(colIndex * (rowIndex + 1));
                    }
                }

                var drawing = sheet.CreateDrawingPatriarch() as XSSFDrawing;
                IClientAnchor anchor = drawing.CreateAnchor(0, 0, 0, 0, 0, 5, 10, 15);

                var chart = drawing.CreateChart(anchor);
                var legend = chart.GetOrAddLegend();
                legend.Position = LegendPosition.TopRight;

                var bottomAxis = chart.CreateValueAxis(AxisPosition.Bottom);
                var leftAxis = chart.CreateValueAxis(AxisPosition.Left);

                var data = chart.CreateData<double, double>(ChartTypes.SCATTER, bottomAxis, leftAxis);
                leftAxis.Crosses = AxisCrosses.AutoZero;

                var xs = XDDFDataSourcesFactory.FromNumericCellRange(sheet, new CellRangeAddress(0, 0, 0, NUM_OF_COLUMNS - 1));
                var ys1 = XDDFDataSourcesFactory.FromNumericCellRange(sheet, new CellRangeAddress(1, 1, 0, NUM_OF_COLUMNS - 1));
                var ys2 = XDDFDataSourcesFactory.FromNumericCellRange(sheet, new CellRangeAddress(2, 2, 0, NUM_OF_COLUMNS - 1));

                var s1=data.AddSeries(xs, ys1);
                s1.SetTitle("s1");
                var s2=data.AddSeries(xs, ys2);
                s2.SetTitle("s2");
                chart.Plot(data);

                // Write the output to a file
                using (FileStream sw = File.Create("test.xlsx"))
                {
                    wb.Write(sw);
                }
            }
        }
    }
}
