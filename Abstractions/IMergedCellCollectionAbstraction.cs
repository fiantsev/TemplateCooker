using System;
using System.Collections.Generic;

namespace PluginAbstraction
{
    public interface IMergedCellCollectionAbstraction : IEnumerable<ICellAbstraction>, IDisposable
    {
    }
}