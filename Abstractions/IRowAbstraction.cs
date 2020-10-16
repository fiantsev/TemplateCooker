using System.Collections.Generic;

namespace Abstractions
{
    public interface IRowAbstraction
    {
        IEnumerable<ICellAbstraction> GetCells();
        IEnumerable<ICellAbstraction> GetUsedCells();
        ICellAbstraction GetCell(int index);
        ICellAbstraction FirstCell();
        void InsertRowsBelow(int rowsCount);
    }
}