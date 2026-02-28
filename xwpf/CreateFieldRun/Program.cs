using NPOI.XWPF.UserModel;

using var doc = new XWPFDocument();
var para = doc.CreateParagraph();

// 1. Add some leading text
var run = para.CreateRun();
run.SetText("This is page ");

// 2. Create the Field Run for the current page number
// The string argument is the Word Field Code
var fieldRun = para.CreateFieldRun();
fieldRun.FieldInstruction = "PAGE";
fieldRun.SetText("1");

// 3. Add more text
para.CreateRun().SetText(" of ");

// 4. Create another Field Run for the total number of pages
var fieldRun2=para.CreateFieldRun();
fieldRun2.FieldInstruction = "NUMPAGES";
fieldRun2.SetText("3");     //fake num of pages, you need to update it in Word

using (FileStream out1 = new FileStream("fieldruns.docx", FileMode.Create))
{
    doc.Write(out1);
}