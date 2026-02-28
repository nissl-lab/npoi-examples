using NPOI.XWPF.UserModel;

using (XWPFDocument doc = new XWPFDocument(File.OpenRead("simpleTable.docx")))
{
    foreach (var table in doc.Tables)
    {
        for (int i = 0; i < table.NumberOfRows; i++)
        {
            var row = table.GetRow(i);
            Console.Write($"Row {i + 1}: ");
            for (int j = 0; j < table.NumberOfColumns; j++)
            {
                var text= row.GetCell(j).GetText();
                Console.Write(text+"|");
            }
            Console.WriteLine();
        }
    }
}