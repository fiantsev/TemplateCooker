using NPOI.SS.UserModel;
using PluginAbstraction;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace NpoiPlugin
{
    [DebuggerDisplay("{_row}")]
    public class RowImplementation : IRowAbstraction
    {
        private IRow _row;

        public RowImplementation(IRow row)
        {
            _row = row;
        }

        public ICellAbstraction FirstCell()
        {
            return new CellImplementation(_row.First());
        }

        public ICellAbstraction GetCell(int index)
        {
            var cell = _row.GetCell(index) ?? _row.CreateCell(index);
            return new CellImplementation(cell);
        }

        public IEnumerable<ICellAbstraction> GetCells()
        {
            return _row.Cells.Select(x => new CellImplementation(x));
        }

        public IEnumerable<ICellAbstraction> GetUsedCells()
        {
            if (_row.PhysicalNumberOfCells == 0) yield break;
            for (var cellIndex = _row.FirstCellNum; cellIndex <= _row.LastCellNum; ++cellIndex)
            {
                var cell = _row.GetCell(cellIndex);
                if (cell == null) continue;
                yield return new CellImplementation(cell);
            }
        }

        public void InsertRowsBelow(int rowsCount)
        {
            _row.Sheet.ShiftRows(_row.RowNum + 1, _row.Sheet.LastRowNum, rowsCount);
        }

        public void Dispose()
        {
        }
    }
}