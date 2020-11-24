﻿using ClosedXML.Excel;
using PluginAbstraction;
using System.Diagnostics;

namespace ClosedXmlPlugin
{
    [DebuggerDisplay("{_range}")]
    public class RangeImplementation : IRangeAbstraction
    {
        private readonly IXLRange _range;

        public RangeImplementation(IXLRange range)
        {
            _range = range;
        }

        public int TopRowIndex => _range.RangeAddress.FirstAddress.RowNumber - 1;

        public int LeftColumnIndex => _range.RangeAddress.FirstAddress.ColumnNumber - 1;

        public int Height => _range.RowCount();

        public int Width => _range.ColumnCount();

        public void Dispose()
        {
        }

        public void CopyTo(ICellAbstraction cell)
        {
            var _cell = _range.Worksheet.Row(cell.RowIndex + 1).Cell(cell.ColumnIndex + 1);
            _cell.Value = this._range;
        }

        public ICellAbstraction TopLeftCell()
        {
            return new CellImplementation(_range.FirstCell());
        }

        public ICellAbstraction BottomRightCell()
        {
            return new CellImplementation(_range.LastCell());
        }
    }
}