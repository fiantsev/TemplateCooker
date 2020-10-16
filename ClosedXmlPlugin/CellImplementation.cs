using Abstractions;
using ClosedXML.Excel;

namespace ClosedXmlPlugin
{
    public class CellImplementation : ICellAbstraction
    {
        private IXLCell _cell;

        public CellImplementation(IXLCell cell)
        {
            _cell = cell;
        }
    }
}