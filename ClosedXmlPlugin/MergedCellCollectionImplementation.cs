using Abstractions;
using ClosedXML.Excel;
using System.Collections;
using System.Collections.Generic;

namespace ClosedXmlPlugin
{
    public class MergedCellCollectionImplementation : IMergedCellCollectionAbstraction
    {
        private IXLCell _nextCell;

        public MergedCellCollectionImplementation(IXLCell fromCell)
        {
            _nextCell = fromCell;
        }

        public IEnumerator<ICellAbstraction> GetEnumerator()
        {
            while (true)
            {
                yield return new CellImplementation(_nextCell); //перепаковываем
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