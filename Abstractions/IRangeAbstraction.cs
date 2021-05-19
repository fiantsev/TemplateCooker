using System;
using System.Collections.Generic;

namespace PluginAbstraction
{
    public interface IRangeAbstraction : IDisposable
    {
        int TopRowIndex { get; }
        int LeftColumnIndex { get; }
        int Height { get; }
        int Width { get; }
        void CopyTo(ICellAbstraction cell);
        ICellAbstraction TopLeftCell();
        ICellAbstraction BottomRightCell();
        IEnumerable<ICellAbstraction> CellsUsed();
        void Merge();
    }
}