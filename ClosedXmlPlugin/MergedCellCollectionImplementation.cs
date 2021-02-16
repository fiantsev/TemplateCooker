using ClosedXML.Excel;
using PluginAbstraction;
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
        private int _fixedRowStep;
        private int _fixedColumnStep;

        public MergedCellCollectionImplementation(IXLCell fromCell, int rowCount, int columnCount, int fixedRowStep, int fixedColumnStep)
        {
            _fromCell = fromCell;
            _rowCount = rowCount;
            _columnCount = columnCount;
            _fixedRowStep = fixedRowStep;
            _fixedColumnStep = fixedColumnStep;
        }

        public IEnumerator<ICellAbstraction> GetEnumerator()
        {
            var nextCell = _fromCell;
            foreach (var rowIndex in Enumerable.Range(0, _rowCount))
            {
                yield return new CellImplementation(nextCell);
                foreach (var columnIndex in Enumerable.Range(1, _columnCount - 1))
                {
                    //если задан фиксированный отступ - использовать его; иначе использовать ширину смердженной ячейки
                    var columnStep = _fixedColumnStep > 0
                        ? _fixedColumnStep
                        : nextCell.IsMerged() ? nextCell.MergedRange().ColumnCount() : 1;

                    nextCell = nextCell.CellRight(columnStep);
                    yield return new CellImplementation(nextCell);
                }

                var cellInTheSameColumnAsFirstCell = nextCell.Worksheet.Cell(nextCell.Address.RowNumber, _fromCell.Address.ColumnNumber);
                //если задан фиксированный отступ - использовать его; иначе использовать высоту смердженной ячейки
                var rowStep = _fixedRowStep > 0
                    ? _fixedRowStep
                    : cellInTheSameColumnAsFirstCell.IsMerged() ? cellInTheSameColumnAsFirstCell.MergedRange().RowCount() : 1;

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