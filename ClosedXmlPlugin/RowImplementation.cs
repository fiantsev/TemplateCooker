using Abstractions;
using ClosedXML.Excel;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace ClosedXmlPlugin
{
    [DebuggerDisplay("{_row}")]
    public class RowImplementation : IRowAbstraction
    {
        private IXLRow _row;

        public RowImplementation(IXLRow row)
        {
            _row = row;
        }

        public ICellAbstraction FirstCell()
        {
            return new CellImplementation(_row.FirstCell());
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

        public IEnumerable<ICellAbstraction> GetUsedCells()
        {
            return _row.CellsUsed().Select(x => new CellImplementation(x));
        }

        public void InsertRowsBelow(int rowsCount)
        {
            _row.InsertRowsBelow(rowsCount);
        }

        public void Dispose()
        {
        }
    }
}