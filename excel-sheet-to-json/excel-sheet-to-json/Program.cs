using System;
using System.IO;
using OfficeOpenXml;
using System.Linq;

namespace excel_sheet_to_json
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Start exexution");
            Console.WriteLine(TranslateExcelSheet.Run() ? "Execution succes" : "Execution failed");
            Console.Read();
        }
    }
}
