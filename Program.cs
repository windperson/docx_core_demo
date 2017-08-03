using System;
using Novacode;

namespace hwapp
{
    class Program
    {
        static void Main(string[] args)
        {
            var fileName = @"test.docx";

            var doc = DocX.Create(fileName);

            Console.WriteLine($" docx file {fileName} created.");

            var headLineFormat = new Formatting();
            headLineFormat.FontFamily = new Font("Arial Black");
            headLineFormat.Size = 18D;
            headLineFormat.Position = 12;
            var head = doc.InsertParagraph("This is title", false,headLineFormat);

            var para = doc.InsertParagraph();
            para.Append("Programmable insert text.");

            Console.WriteLine("text inserted.");

            doc.Save();

            Console.WriteLine("document saved.");
        }
    }
}
