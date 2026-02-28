/* ================================================================
 * Author: Tony Qu 
 * Author's email: tonyqus (at) gmail.com 
 * NPOI Examples: https://github.com/nissl-lab/npoi-examples
 * ==============================================================*/

using NPOI.Util;
using NPOI.XWPF.UserModel;
using System.IO;

namespace SimpleTable
{
    class Program
    {
        static void Main(string[] args)
        {
            using (XWPFDocument doc = new XWPFDocument())
            {
                XWPFParagraph para = doc.CreateParagraph();
                XWPFRun r0 = para.CreateRun();
                r0.SetText("Title1");
                para.BorderTop = Borders.Thick;
                para.FillBackgroundColor = "EEEEEE";
                para.FillPattern = NPOI.OpenXmlFormats.Wordprocessing.ST_Shd.diagStripe;

                XWPFTable table = doc.CreateTable(3, 3);

                table.GetRow(1).GetCell(1).SetText("EXAMPLE OF TABLE");


                XWPFTableCell c1 = table.GetRow(0).GetCell(0);
                XWPFParagraph p1 = c1.AddParagraph();   //don't use doc.CreateParagraph
                XWPFRun r1 = p1.CreateRun();
                r1.SetText("The quick brown fox");
                r1.IsBold = true;

                r1.FontFamily = "Courier";
                r1.Underline = UnderlinePatterns.DotDotDash;
                r1.TextPosition = 100;
                c1.SetColor("FF0000");

                table.GetRow(2).GetCell(2).SetText("only text");
                var r2 = table.GetRow(2).GetCell(0).AddParagraph().CreateRun();
                var widthEmus = (int)(400.0 * 9525);
                var heightEmus = (int)(300.0 * 9525);

                using (FileStream picData = new FileStream("image/HumpbackWhale.jpg", FileMode.Open, FileAccess.Read))
                {
                    r2.AddPicture(picData, (int)PictureType.JPEG, "image1", widthEmus, heightEmus);
                }

                using (FileStream fs = new FileStream("simpleTable.docx", FileMode.Create))
                {
                    doc.Write(fs);
                }
            }
        }
    }
}
