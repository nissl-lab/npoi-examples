using NPOI.SS.UserModel;
using NPOI.XSSF.EventUserModel;
using NPOI.XSSF.Extractor;
using NPOI.XSSF.Model;
using NPOI.XSSF.UserModel;
using NSAX.Helpers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReadSheetsByXSSFReader
{
    public class SheetContentHandler : XSSFSheetHandler.ISheetContentsHandler
    {
        public SheetContentHandler()
        { 
            
        }
        public void Cell(string cellReference, string formattedValue, XSSFComment comment)
        {
            Console.Write(formattedValue);
            Console.Write("|");
        }

        public void EndRow(int rowNum)
        {
            Console.WriteLine();
        }

        public void EndSheet()
        {
            //do nothing
        }

        public void HeaderFooter(string text, bool IsHeader, string tagName)
        {
            //do nothing
        }

        public void StartRow(int rowNum)
        {
            Console.Write($"Row {rowNum}: ");
        }
    }
}
