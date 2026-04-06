using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XDDF.UserModel.Chart;
using NPOI.XSSF.UserModel;
using System.IO;

namespace LineChart
{
    class Program
    {
        const int NUM_OF_ROWS = 3;
        const int NUM_OF_COLUMNS = 10;

        static void CreateChart(XSSFDrawing drawing, ISheet sheet, IClientAnchor anchor, string chartTitle, string serie1, string serie2, bool enableMajorGridline=false)
        {
            var chart = drawing.CreateChart(anchor);
            chart.SetTitleText(chartTitle);
            var legend = chart.GetOrAddLegend();
            legend.Position = LegendPosition.TopRight;


            // Use a category axis for the bottom axis.
            var bottomAxis = chart.CreateCategoryAxis(AxisPosition.Bottom);
            var leftAxis = chart.CreateValueAxis(AxisPosition.Left);
            leftAxis.Crosses = AxisCrosses.AutoZero;


            var data = chart.CreateData<double, double>(ChartTypes.LINE, bottomAxis, leftAxis);

            var xs = XDDFDataSourcesFactory.FromNumericCellRange(sheet, new CellRangeAddress(0, 0, 0, NUM_OF_COLUMNS - 1));
            var ys1 = XDDFDataSourcesFactory.FromNumericCellRange(sheet, new CellRangeAddress(1, 1, 0, NUM_OF_COLUMNS - 1));
            var ys2 = XDDFDataSourcesFactory.FromNumericCellRange(sheet, new CellRangeAddress(2, 2, 0, NUM_OF_COLUMNS - 1));

            var s1 = data.AddSeries(xs, ys1);
            s1.SetTitle(serie1);
            var s2 = data.AddSeries(xs, ys2);
            s2.SetTitle(serie2);

            data.GetCategoryAxis().GetOrAddMajorGridProperties();
            data.GetValueAxis()[0].GetOrAddMajorGridProperties();

            chart.Plot(data);
            
        }

        static void Main(string[] args)
        {
            using (IWorkbook wb = new XSSFWorkbook())
            {
                ISheet sheet = wb.CreateSheet("linechart");
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
                IClientAnchor anchor1 = drawing.CreateAnchor(0, 0, 0, 0, 0, 5, 10, 15);
                CreateChart(drawing, sheet, anchor1,"Test 1","title1", "title2");
                IClientAnchor anchor2 = drawing.CreateAnchor(0, 0, 0, 0, 0, 20, 10, 35);
                CreateChart(drawing, sheet, anchor2,"Test2", "s1", "s2", true);
                using (FileStream fs = File.Create("test.xlsx"))
                {
                    wb.Write(fs, false);
                }
            }
        }
    }
}
