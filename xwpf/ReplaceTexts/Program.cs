/* ================================================================
 * Author: Tony Qu 
 * Author's email: tonyqus (at) gmail.com 
 * NPOI Examples: https://github.com/nissl-lab/npoi-examples
 * ==============================================================*/

using NPOI.XWPF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;

namespace ReplaceTexts
{
    class Program
    {
        static List<string> placeHolderDictionary = new List<string>(); 
        static void Main(string[] args)
        {
            placeHolderDictionary.Add("{startdate}");
            placeHolderDictionary.Add("{enddate}");
            placeHolderDictionary.Add("{yushouzzmj}");
            placeHolderDictionary.Add("{yushouzzmj_t}");
            placeHolderDictionary.Add("{yushouzzmj_tb}");
            placeHolderDictionary.Add("{yushouzzmj_h}");
            placeHolderDictionary.Add("{yushouzzmj_hb}");
            placeHolderDictionary.Add("{yushouzzts}");
            placeHolderDictionary.Add("{yushouzzts_t}");
            placeHolderDictionary.Add("{yushouzzts_tb}");
            placeHolderDictionary.Add("{yushouzzts_h}");
            placeHolderDictionary.Add("{yushouzzts_hb}");
            placeHolderDictionary.Add("{yushouzz90mj}");
            placeHolderDictionary.Add("{yushouzz90mj_t}");
            placeHolderDictionary.Add("{yushouzz90mj_tb}");
            placeHolderDictionary.Add("{yushouzz90ts}");
            placeHolderDictionary.Add("{yushouzz90ts_t}");
            placeHolderDictionary.Add("{yushouzz90ts_tb}");
            placeHolderDictionary.Add("{yushouzz90144mj}");
            placeHolderDictionary.Add("{yushouzz90144mj_t}");
            placeHolderDictionary.Add("{yushouzz90144mj_tb}");
            placeHolderDictionary.Add("{yushouzz90144ts}");
            placeHolderDictionary.Add("{yushouzz90144ts_t}");
            placeHolderDictionary.Add("{yushouzz90144ts_tb}");

            var template = @"Template1.docx";
            using (var rs = File.OpenRead(template))
            {
                var generateFile = @"output1.docx";
                using var doc = new XWPFDocument(rs);
                //doc.FindAndReplaceText("{yushouzzmj}", "Test");   
                using (var ws = File.Create(generateFile))
                {
                    doc.Write(ws);
                }
            }
        }
    }
}
