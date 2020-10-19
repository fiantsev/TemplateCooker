using NPOI.SS.UserModel;
using PluginAbstraction;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace NpoiPlugin
{
    public class MergedCellCollectionImplementation : IMergedCellCollectionAbstraction
    {
        private ICell _fromCell;
        private int _rowCount;
        private int _columnCount;

        public MergedCellCollectionImplementation(ICell fromCell, int rowCount, int columnCount)
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
                    var columnStep = nextCell.IsMergedCell
                        ? GetMergedRangeColumnCount(nextCell)
                        : 1;
                    nextCell = NextCellToRight(nextCell, columnStep);
                    yield return new CellImplementation(nextCell);
                }

                var cellInTheSameColumnAsFirstCell = nextCell.Row.GetCell(_fromCell.RowIndex);
                var rowStep = cellInTheSameColumnAsFirstCell.IsMergedCell
                    ? GetMergedRangeRowCount(cellInTheSameColumnAsFirstCell)
                    : 1;
                nextCell = NextCellToBelow(cellInTheSameColumnAsFirstCell, rowStep);
            }
        }

        private int GetMergedRangeColumnCount(ICell cell)
        {
            var address = cell.Sheet.MergedRegions.First(reg => reg.IsInRange(cell.RowIndex, cell.ColumnIndex));
            return address.LastColumn - address.FirstColumn;
        }

        private int GetMergedRangeRowCount(ICell cell)
        {
            var address = cell.Sheet.MergedRegions.First(reg => reg.IsInRange(cell.RowIndex, cell.ColumnIndex));
            return address.LastRow - address.FirstRow;
        }

        private ICell NextCellToRight(ICell cell, int step)
        {
            return cell.Row.GetCell(cell.ColumnIndex + step) ?? cell.Row.CreateCell(cell.ColumnIndex + step);
        }

        private ICell NextCellToBelow(ICell cell, int step)
        {
            var row = cell.Sheet.GetRow(cell.Row.RowNum + step) ?? cell.Sheet.CreateRow(cell.Row.RowNum + step);
            var outCell = row.GetCell(cell.ColumnIndex) ?? row.CreateCell(cell.ColumnIndex);
            return outCell;
        }

        //public IEnumerator<ICellAbstraction> GetEnumerator()
        //{
        //    var nextCell = _fromCell;
        //    foreach (var rowIndex in Enumerable.Range(0, _rowCount))
        //    {
        //        yield return new CellImplementation(nextCell);
        //        foreach (var columnIndex in Enumerable.Range(1, _columnCount - 1))
        //        {
        //            var columnStep = nextCell.IsMerged()
        //                ? nextCell.MergedRange().ColumnCount()
        //                : 1;
        //            nextCell = nextCell.CellRight(columnStep);
        //            yield return new CellImplementation(nextCell);
        //        }

        //        var cellInTheSameColumnAsFirstCell = nextCell.Worksheet.Cell(nextCell.Address.RowNumber, _fromCell.Address.ColumnNumber);
        //        var rowStep = cellInTheSameColumnAsFirstCell.IsMerged()
        //            ? cellInTheSameColumnAsFirstCell.MergedRange().RowCount()
        //            : 1;
        //        nextCell = cellInTheSameColumnAsFirstCell.CellBelow(rowStep);
        //    }
        //}

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public void Dispose()
        {
        }
    }
}