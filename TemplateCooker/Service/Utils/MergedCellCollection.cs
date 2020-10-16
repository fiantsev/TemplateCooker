using ClosedXML.Excel;
using System.Collections;
using System.Collections.Generic;

namespace TemplateCooker.Service.Utils
{
    public class MergedCellCollection : IEnumerable<IXLCell>
    {
        private IXLCell _nextCell;

        public MergedCellCollection(IXLCell fromCell)
        {
            _nextCell = fromCell;
        }

        public IEnumerator<IXLCell> GetEnumerator()
        {
            while (true)
            {
                yield return _nextCell;
                var step = _nextCell.IsMerged()
                    ? _nextCell.MergedRange().ColumnCount()
                    : 1;
                _nextCell = _nextCell.CellRight(step);
            }
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
    }
}