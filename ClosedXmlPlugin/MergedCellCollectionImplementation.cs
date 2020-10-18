using PluginAbstraction;
using ClosedXML.Excel;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXmlPlugin
{
    public class MergedCellCollectionImplementation : IMergedCellCollectionAbstraction
    {
        private IXLCell _fromCell;
        private int _rowCount;
        private int _columnCount;

        public MergedCellCollectionImplementation(IXLCell fromCell, int rowCount, int columnCount)
        {
            _fromCell = fromCell;
            _rowCount = rowCount;
            _columnCount = columnCount;
        }

        public IEnumerator<ICellAbstraction> GetEnumerator()
        {
            var nextCell = _fromCell;
            foreach (var rowIndex in Enumerable.Range(0, _rowCount))
            {
                yield return new CellImplementation(nextCell);
                foreach (var columnIndex in Enumerable.Range(1, _columnCount - 1))
                {
                    var columnStep = nextCell.IsMerged()
                        ? nextCell.MergedRange().ColumnCount()
                        : 1;
                    nextCell = nextCell.CellRight(columnStep);
                    yield return new CellImplementation(nextCell);
                }

                var cellInTheSameColumnAsFirstCell = nextCell.Worksheet.Cell(nextCell.Address.RowNumber, _fromCell.Address.ColumnNumber);
                var rowStep = cellInTheSameColumnAsFirstCell.IsMerged()
                    ? cellInTheSameColumnAsFirstCell.MergedRange().RowCount()
                    : 1;
                nextCell = cellInTheSameColumnAsFirstCell.CellBelow(rowStep);
            }
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public void Dispose()
        {
        }
    }
}