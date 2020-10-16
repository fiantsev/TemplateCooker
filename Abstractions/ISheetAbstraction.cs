using System.Collections.Generic;

namespace Abstractions
{
    public interface ISheetAbstraction
    {
        IEnumerable<IRowAbstraction> GetRows();
        IEnumerable<IRowAbstraction> GetUsedRows();
        IRowAbstraction GetRow(int index);
        int SheetIndex { get; }
    }
}