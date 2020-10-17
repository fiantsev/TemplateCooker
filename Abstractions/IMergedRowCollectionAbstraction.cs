﻿using System;
using System.Collections.Generic;

namespace Abstractions
{
    public interface IMergedRowCollectionAbstraction : IEnumerable<IRowAbstraction>, IDisposable
    {
    }
}