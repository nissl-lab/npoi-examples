using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

using (IWorkbook wb = new XSSFWorkbook(new FileStream("Sample.xlsx", FileMode.Open)))
{
    var sheet = wb.GetSheetAt(0);
    for (int i = 0; i < 3; i++)
    { 
        sheet.AutoSizeColumn(i);
    }
    using (FileStream sw = File.Create("output.xlsx"))
    {
        wb.Write(sw, false);
    }
}