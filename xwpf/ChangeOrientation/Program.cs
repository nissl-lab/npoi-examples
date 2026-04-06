using NPOI.XWPF.UserModel;
using System;
using System.IO;
using NPOI.OpenXmlFormats.Wordprocessing;

namespace ChangeOrientation
{
    internal class Program
    {
        static void Main(string[] args)
        {
            using XWPFDocument doc = new XWPFDocument();

            var run = doc.CreateParagraph().CreateRun();
            run.SetText("Hello World!");

            doc.ChangeOrientation(ST_PageOrientation.landscape);
            using (FileStream fs = new FileStream("test.docx", FileMode.Create))
            {
                doc.Write(fs);
            }
        }
    }
}
