using Abstractions;
using ClosedXML.Excel;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXmlPlugin
{
    public class RowImplementation : IRowAbstraction
    {
        private IXLRow _row;

        public RowImplementation(IXLRow row)
        {
            _row = row;
        }

        public ICellAbstraction GetCell(int index)
        {
            var cell = new CellImplementation(_row.Cell(index + 1));
            return cell;
        }

        public IEnumerable<ICellAbstraction> GetCells()
        {
            var cells = (_row.Cells() as IEnumerable<IXLCell>).Select(x => new CellImplementation(x));
            return cells;
        }
    }
}