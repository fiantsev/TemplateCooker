using OfficeOpenXml;
using System;
using System.IO;

namespace EpplusPlaygroundApp
{
    class Program
    {
        static void Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var fileInfo = new FileInfo("./myworkbook.xlsx");
            using (var p = new ExcelPackage(fileInfo))
            {
                var sheet = p.Workbook.Worksheets[0];
                Console.WriteLine($"Dimension: {sheet.Dimension}");

                for(var i=1; i < sheet.Dimension.End.Row+1; ++i)
                {
                    for(var j=1; j < sheet.Dimension.End.Column+1;++j)
                    Console.WriteLine($"{sheet.Cells[i,j]}: ({sheet.Cells[i, j].Value?.GetType().FullName}) {sheet.Cells[i, j].Value} ({sheet.Cells[i, j].Formula?.GetType().FullName}) {sheet.Cells[i, j].Formula} {sheet.Cells[i, j].Formula == string.Empty}");
                }
            }

            //Console.WriteLine("Press any key");
            Console.ReadKey();
        }
    }
}
