using NPOI.OpenXml4Net.OPC;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.XSSF.EventUserModel;
using NPOI.XSSF.Model;
using NSAX;
using NSAX.AElfred;
using ReadSheetsByXSSFReader;
using System.Runtime.Serialization;
using System.Xml;

OPCPackage package = OPCPackage.Open("sample.xlsx", PackageAccess.READ);

XSSFReader reader = new XSSFReader(package);

// Read styles table
var styles = reader.StylesTable;
var sst = reader.SharedStringsTable;

// Get worksheets
XSSFReader.SheetIterator sheetIterator = (XSSFReader.SheetIterator)reader.GetSheetsData();

int sheetIndex = 0;
while (sheetIterator.MoveNext())
{
    using Stream sheetStream = sheetIterator.Current;
    string sheetName = sheetIterator.SheetName;
    
    Console.WriteLine($"\n--- Sheet {sheetIndex}: {sheetName} ---");
    SAXDriver sheetParser = new SAXDriver();
    InputSource sheetSource = new InputSource(sheetStream);
    
    var handler=new XSSFSheetHandler(styles, sst,new SheetContentHandler() , true);
    sheetParser.ContentHandler = handler;
    sheetParser.Parse(sheetSource);
    sheetIndex++;
}
package.Close();
