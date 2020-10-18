using System;
using System.Collections.Generic;

namespace PluginAbstraction
{
    public interface IRowAbstraction : IDisposable
    {
        IEnumerable<ICellAbstraction> GetCells();
        IEnumerable<ICellAbstraction> GetUsedCells();
        ICellAbstraction GetCell(int index);
        ICellAbstraction FirstCell();
        void InsertRowsBelow(int rowsCount);
    }
}