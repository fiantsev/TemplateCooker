using ClosedXML.Excel;
using System.Collections;
using System.Collections.Generic;

namespace TemplateCooker.Service.Utils
{
    public class MergedRowCollection : IEnumerable<IXLRow>
    {
        private IXLCell _fromCell;
        private IXLRow _nextRow;

        public MergedRowCollection(IXLCell fromCell)
        {
            _fromCell = fromCell;
            _nextRow = fromCell.WorksheetRow();
        }

        public IEnumerator<IXLRow> GetEnumerator()
        {
            while (true)
            {
                yield return _nextRow;

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