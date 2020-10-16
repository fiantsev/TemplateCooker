using Abstractions;
using ClosedXML.Excel;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXmlPlugin
{
    public class SheetImplementation : ISheetAbstraction
    {
        private IXLWorksheet _sheet;

        public SheetImplementation(IXLWorksheet sheet)
        {
            _sheet = sheet;
        }

        public IRowAbstraction GetRow(int index)
        {
            var row = new RowImplementation(_sheet.Row(index + 1));
            return row;
        }

        public IEnumerable<IRowAbstraction> GetRows()
        {
            var rows = (_sheet.Rows() as IEnumerable<IXLRow>).Select(x => new RowImplementation(x));
            return rows;
        }
    }
}