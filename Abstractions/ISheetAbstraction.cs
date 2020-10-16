using System.Collections.Generic;

namespace Abstractions
{
    public interface ISheetAbstraction
    {
        int SheetIndex { get; }
        string SheetName { get; }

        IEnumerable<IRowAbstraction> GetRows();
        IEnumerable<IRowAbstraction> GetUsedRows();
        IRowAbstraction GetRow(int index);
    }
}