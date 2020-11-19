using System;

namespace PluginAbstraction
{
    public interface IRangeAbstraction : IDisposable
    {
        int TopRowIndex { get; }
        int LeftColumnIndex { get; }
        int Height { get; }
        int Width { get; }
    }
}