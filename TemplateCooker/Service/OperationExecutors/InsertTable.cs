using PluginAbstraction;
using System.Collections.Generic;
using System.Linq;
using TemplateCooking.Domain.Layout;

namespace TemplateCooking.Service.OperationExecutors
{
    public class InsertTable : IOperationExecutor
    {
        public class Operation : AbstractOperation
        {
            public SrcPosition Position { get; set; }
            public List<List<object>> Table { get; set; }
            public bool PreserveStyleOfFirstCell { get; set; }
        }

        public void Execute(IWorkbookAbstraction workbook, object untypedOptions)
        {
            var options = (Operation)untypedOptions;
            var table = options.Table;
            var sheet = workbook.GetSheet(options.Position.SheetIndex);
            var topLeftCell = sheet.GetRow(options.Position.RowIndex).GetCell(options.Position.ColumnIndex);

            var rowCount = table.Count;
            var columnCount = rowCount == 0 ? 0 : table[0].Count;

            //удаляем маркер
            if (rowCount == 0 || columnCount == 0)
                topLeftCell.SetValue(string.Empty);

            //if (tableInjection.LayoutShift == LayoutShiftType.MoveRows)
            //здесь мы копируем контент первой строки ниже - чтобы стили ячеек (включая смердженные регионы) скопировались
            if(options.PreserveStyleOfFirstCell)
                CloneFirstRowBelow(sheet, topLeftCell, rowCount, columnCount);

            //здесь мы вставляем контент в ячейки
            var cellCounter = 0;
            foreach (var cell in topLeftCell.GetMergedCells(rowCount, columnCount))
            {
                var rowIndex = cellCounter / columnCount;
                var columnIndex = cellCounter % columnCount;
                cell.SetValue(table[rowIndex][columnIndex]);
                ++cellCounter;
            }
        }


        private void CloneFirstRowBelow(ISheetAbstraction sheet, ICellAbstraction topLeftCell, int rowCount, int columnCount)
        {
            if (rowCount == 0 || columnCount == 0)
                return;

            var firstRowMergedCells = topLeftCell.GetMergedCells(1, columnCount).ToList();
            var firstCellHeight = firstRowMergedCells.First().GetMergedRange().Height;
            var range = sheet.GetRange(firstRowMergedCells.First(), firstRowMergedCells.Last().GetMergedRange().BottomRightCell());

            for (var i = 1; i < rowCount; ++i)
            {
                var toCell = sheet.GetRow(topLeftCell.RowIndex + i * firstCellHeight).GetCell(topLeftCell.ColumnIndex);
                range.CopyTo(toCell);
            }
        }
    }
}