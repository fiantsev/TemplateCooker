using ClosedXML.Excel;
using PluginAbstraction;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace ClosedXmlPlugin
{
    [DebuggerDisplay("{_sheet}")]
    public class SheetImplementation : ISheetAbstraction
    {
        private IXLWorksheet _sheet;

        public SheetImplementation(IXLWorksheet sheet)
        {
            _sheet = sheet;
        }

        public int SheetIndex => _sheet.Position - 1;
        public string SheetName => _sheet.Name;

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

        public IEnumerable<IRowAbstraction> GetUsedRows()
        {
            return _sheet.RowsUsed().Select(x => new RowImplementation(x));
        }

        public void Dispose()
        {
        }
    }
}