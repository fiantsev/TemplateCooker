using System;
using System.Collections.Generic;

namespace PluginAbstraction
{
    public interface ISheetAbstraction : IDisposable
    {
        int SheetIndex { get; }
        string SheetName { get; }

        IEnumerable<IRowAbstraction> GetRows();
        IEnumerable<IRowAbstraction> GetUsedRows();
        IRowAbstraction GetRow(int index);
    }
}