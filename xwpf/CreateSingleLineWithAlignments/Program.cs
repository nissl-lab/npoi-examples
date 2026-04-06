/* ================================================================
 * Author: Tony Qu 
 * Author's email: tonyqus (at) gmail.com 
 * NPOI Examples: https://github.com/nissl-lab/npoi-examples
 * ==============================================================*/

using NPOI.XWPF.UserModel;
using System.IO;

namespace CreateSingleLineWithAlignments
{
    /// <summary>
    /// https://stackoverflow.com/questions/36211987/npoi-xwpf-how-can-i-place-text-on-a-single-line-that-is-both-left-right-justif
    /// </summary>
    class Program
    {
        static void Main(string[] args)
        {
            using (XWPFDocument doc = new XWPFDocument())
            {
                var p= doc.CreateParagraph();
                var ctp = p.GetCTP();
                var pPr = ctp.AddNewPPr();
                var tab = pPr.AddNewTabs().AddNewTab();
                tab.pos = "8000";
                tab.val = NPOI.OpenXmlFormats.Wordprocessing.ST_TabJc.right;

                var r1 = p.CreateRun();
                r1.SetText("Left aligned");
                r1.AddTab();

                var r2 = p.CreateRun();
                r2.SetText("Right aligned");

                using (FileStream sw = File.Create("alignmentOnSingleLine.docx"))
                {
                    doc.Write(sw);
                }
            }
        }
    }
}