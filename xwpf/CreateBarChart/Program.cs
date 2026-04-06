// Category Axis Data
using NPOI.SS.Util;
using NPOI.XDDF.UserModel.Chart;
using NPOI.XWPF.UserModel;
using System.Linq;

static void ShowUsage(){
    Console.WriteLine("Usage: BarChartExample <bar-chart-template.docx> <bar-chart-data.txt>");
    Console.WriteLine("    bar-chart-template.docx     template with a bar chart");
    Console.WriteLine("    bar-chart-data.txt          the model to set. First line is chart title, " +
                       "then go pairs {axis-label value}");
}

if (args.Length < 2)
{
    ShowUsage();
    return;
}

using var argIS = File.OpenRead(args[0]);
using TextReader modelReader = new StreamReader(args[1]);
string chartTitle = modelReader.ReadLine();
string seriesText = modelReader.ReadLine();
String[] series = seriesText == null ? new string[0] : seriesText.Split(",");

// Category Axis Data
List<String> listLanguages = new List<string>(10);

// Values

List<Double> listCountries = new List<double>(10);
List<Double> listSpeakers = new List<double>(10);
String ln;
while((ln = modelReader.ReadLine()) != null) {
    String[] vals = ln.Split(",");
    listCountries.Add(Double.Parse(vals[0]));
    listSpeakers.Add(Double.Parse(vals[1]));
    listLanguages.Add(vals[2]);
}

String[] categories = listLanguages.ToArray();
Double[] values1 = listCountries.ToArray();
Double[] values2 = listSpeakers.ToArray();

using (XWPFDocument doc = new XWPFDocument(argIS))
{
    XWPFChart chart = doc.GetCharts()[0];
    SetBarData(chart, chartTitle, series, categories, values1, values2);
    // save the result
    using (var stream = File.Create("bar-chart-demo-output.docx"))
    {
        doc.Write(stream);
    }

    Console.WriteLine("Done");
}
static void SetBarData(XWPFChart chart, String chartTitle, String[] series, String[] categories, Double[] values1, Double[] values2)
{
    var data = chart.GetChartSeries<string, double>();
    var bar = data[0] as XDDFBarChartData<string,double>;
    
    int numOfPoints = categories.Length;

     String categoryDataRange = chart.FormatRange(new CellRangeAddress(1, numOfPoints, 0, 0));
     String valuesDataRange = chart.FormatRange(new CellRangeAddress(1, numOfPoints, 1, 1));
     String valuesDataRange2 = chart.FormatRange(new CellRangeAddress(1, numOfPoints, 2, 2));
    var categoriesData = XDDFDataSourcesFactory.FromArray(categories, categoryDataRange, 0);
     var valuesData = XDDFDataSourcesFactory.FromArray(values1, valuesDataRange, 1);
    values1[6] = 16.0; // if you ever want to change the underlying data
    var valuesData2 = XDDFDataSourcesFactory.FromArray(values2, valuesDataRange2, 2);

    var series1 = bar.GetSeries(0);
    series1.SetTitle(series[0], chart.SetSheetTitle(series[0], 0));

    var series2 = bar.AddSeries(categoriesData, valuesData2);
    series2.SetTitle(series[1], chart.SetSheetTitle(series[1], 1));

    bar.SetVaryColors(true);
    bar.BarDirection = BarDirection.Col;
    chart.Plot(bar);

    XDDFChartLegend legend = chart.GetOrAddLegend();
    legend.Position = LegendPosition.Left;
    legend.IsOverlay = false;

    chart.SetTitleText(chartTitle);
    chart.TitleOverlay = false;
}