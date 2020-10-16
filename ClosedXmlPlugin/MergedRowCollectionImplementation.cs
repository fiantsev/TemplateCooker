using Abstractions;
using ClosedXML.Excel;
using System.Collections;
using System.Collections.Generic;

namespace ClosedXmlPlugin
{
    public class MergedRowCollectionImplementation : IMergedRowCollectionAbstraction
    {
        private IXLCell _fromCell;
        private IXLRow _nextRow;

        public MergedRowCollectionImplementation(IXLCell fromCell)
        {
            _fromCell = fromCell;
            _nextRow = fromCell.WorksheetRow();
        }

        public IEnumerator<IRowAbstraction> GetEnumerator()
        {
            while (true)
            {
                yield return new RowImplementation(_nextRow);

                var firstCellOfRow = _nextRow.Worksheet.Cell(_nextRow.FirstCell().Address.RowNumber, _fromCell.Address.ColumnNumber);
                var step = firstCellOfRow.IsMerged()
                    ? firstCellOfRow.MergedRange().RowCount()
                    : 1;
                _nextRow = _nextRow.RowBelow(step);
            }
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
    }
}