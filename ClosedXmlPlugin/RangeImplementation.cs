using ClosedXML.Excel;
using PluginAbstraction;
using System.Diagnostics;

namespace ClosedXmlPlugin
{
    [DebuggerDisplay("_range")]
    public class RangeImplementation : IRangeAbstraction
    {
        private readonly IXLRange _range;

        public RangeImplementation(IXLRange range)
        {
            _range = range;
        }

        public int TopRowIndex => _range.RangeAddress.FirstAddress.RowNumber - 1;

        public int LeftColumnIndex => _range.RangeAddress.FirstAddress.ColumnNumber - 1;

        public int Height => _range.RowCount();

        public int Width => _range.ColumnCount();

        public void Dispose()
        {
        }
    }
}