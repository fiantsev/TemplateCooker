using System;
using System.Collections.Generic;

namespace Abstractions
{
    public interface IMergedCellCollectionAbstraction : IEnumerable<ICellAbstraction>, IDisposable
    {
    }
}