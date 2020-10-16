using Abstractions;
using ClosedXML.Excel;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ClosedXmlPlugin
{
    public class WorkbookImplementation : IWorkbookAbstraction
    {
        private XLWorkbook _workbook;

        public WorkbookImplementation()
        {
            _workbook = new XLWorkbook();
        }

        public WorkbookImplementation(Stream stream)
        {
            _workbook = new XLWorkbook(stream);
        }

        public ISheetAbstraction GetSheet(int index)
        {
            var sheet = new SheetImplementation(_workbook.Worksheet(index + 1));
            return sheet;
        }

        public IEnumerable<ISheetAbstraction> GetSheets()
        {
            var sheets = _workbook.Worksheets.Select(x => new SheetImplementation(x));
            return sheets;
        }
    }
}
