using System;
using System.Diagnostics;
using System.IO;

namespace HtmlToWord
{
    class Program
    {
        static void Main(string[] args)
        {
            //string htmlFilename = "Demo.html";
            //string htmlFilename = "Demo2.html";
            //string htmlFilename = "Invoice.html";
            string htmlFilename = "DVLA.DeclarationLetter.htm";
            string html = File.ReadAllText(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, htmlFilename));
            byte[] wordData = WordHelper.HtmlToWord(html);
            string wordFilename = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Demo.docx");
            File.WriteAllBytes(wordFilename, wordData);

            Process process1 = new Process();
            process1.StartInfo.FileName = @"C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE";
            process1.StartInfo.Arguments = wordFilename;
            process1.StartInfo.WindowStyle = ProcessWindowStyle.Maximized;
            process1.Start();
        }
    }
}

// https://stackoverflow.com/questions/5431580/convert-html-to-docx-in-c-sharp
// https://github.com/onizet/html2openxml

// VS使用NPOI替换word模板中的关键字
// https://blog.csdn.net/duke_zcm/article/details/89541642