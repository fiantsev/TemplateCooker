using System.Collections.Generic;

namespace Abstractions
{
    public interface IRowAbstraction
    {
        IEnumerable<ICellAbstraction> GetCells();
        ICellAbstraction GetCell(int index);
    }
}