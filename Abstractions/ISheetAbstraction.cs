using System.Collections.Generic;

namespace Abstractions
{
    public interface ISheetAbstraction
    {
        IEnumerable<IRowAbstraction> GetRows();
        IRowAbstraction GetRow(int index);
    }
}