using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OpenXmlBuilderDemoConsole
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello, World!");

            new MsOfficeDemo().ProcessWord("file1.docx");
            new MsOfficeDemo().CreateExcel("file2.xlsx");
        }
    }
}
