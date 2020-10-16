using System.Collections.Generic;

namespace Abstractions
{
    public interface IWorkbookAbstraction
    {
        IEnumerable<ISheetAbstraction> GetSheets();
        ISheetAbstraction GetSheet(int index);
    }
}