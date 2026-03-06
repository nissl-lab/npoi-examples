using NPOI.HPSF;
using NPOI.Util;
using NPOI.XWPF.UserModel;

using (FileStream fs = new FileStream("pictures.docx", FileMode.Open, FileAccess.ReadWrite))
{
    using (XWPFDocument doc = new XWPFDocument(fs))
    {

        IList<XWPFParagraph> paragraphs = doc.Paragraphs;   //get all paragraphs

        for (int par = 0; par < paragraphs.Count; par++)
        {
            IList<XWPFRun> runs = paragraphs[par].Runs;          //get all paragraphs

            for (int runNum = 0; runNum < runs.Count; runNum++)     //get all runs in each paragraph
            {
                if (runs[runNum].GetEmbeddedPictures().Count != 0)  //get picture 
                {
                    List<XWPFPicture> pictures = runs[runNum].GetEmbeddedPictures();

                    var ctinline = runs[runNum].GetCTR().GetDrawingArray(0).inline[0];
                    ctinline.extent.cx = Units.ToEMU(300);
                    ctinline.extent.cy = Units.ToEMU(250);

                    pictures[0].GetCTPicture().spPr.xfrm.ext.cx = Units.ToEMU(300);   //resize picture width
                    pictures[0].GetCTPicture().spPr.xfrm.ext.cy = Units.ToEMU(250);   //resize picture height

                    using (FileStream imgWrite = File.Create("output.docx"))
                    { doc.Write(imgWrite); }
                }
            }
        }
    }
}