using PluginAbstraction;
using System.Collections.Generic;
using System.Linq;
using TemplateCooking.Domain.Layout;

namespace TemplateCooking.Service.OperationExecutors
{
    /// <summary>
    /// Вставить табличные данные в указанную ячейку
    /// </summary>
    public class InsertTable : IOperationExecutor
    {
        public class Operation : AbstractOperation
        {
            /// <summary>
            /// позиция верхней левой ячейки начиная с которой будет располагаться таблица
            /// </summary>
            public SrcPosition Position { get; set; }
            /// <summary>
            /// данные таблицы
            /// </summary>
            public List<List<object>> Table { get; set; }
            /// <summary>
            /// Применить стили ячеек первой строки ко всем последующим ячейкам (в строках ниже).
            /// Может использоваться при вставление таблиц с динамическим количеством строк
            /// </summary>
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

            //здесь мы копируем контент первой строки ниже - чтобы стили ячеек (включая смердженные регионы) скопировались
            if(options.PreserveStyleOfFirstCell)
                CloneFirstRowBelow(sheet, topLeftCell, rowCount, columnCount);

            //здесь мы вставляем контент в ячейки
            var cellCounter = 0;
            foreach (var cell in topLeftCell.GetMergedCells(rowCount, columnCount))
            {
                var rowIndex = cellCounter / columnCount;
                var columnIndex = cellCounter % columnCount;
                var value = table[rowIndex][columnIndex];
                cell.SetValue(value);
                ++cellCounter;
            }
        }


        private void CloneFirstRowBelow(ISheetAbstraction sheet, ICellAbstraction topLeftCell, int rowCount, int columnCount)
        {
            if (rowCount == 0 || columnCount == 0)
                return;

            var topLeftCellHeight = topLeftCell.GetMergedRange().Height;
            var bottomRightCell = topLeftCell.GetMergedCells(1, columnCount).Last().GetMergedRange().BottomRightCell();
            var range = sheet.GetRange(topLeftCell, bottomRightCell);

            //обработка кейса: если первая ячейка смердженная а последняя нет
            //область для копирования составляем по высоте первой ячейки
            if (topLeftCellHeight > range.Height)
                range = sheet.GetRange(topLeftCell, sheet.GetRow(topLeftCell.RowIndex + topLeftCellHeight - 1).GetCell(bottomRightCell.ColumnIndex));

            for (var i = 1; i < rowCount; ++i)
            {
                var toCell = sheet.GetRow(topLeftCell.RowIndex + i * topLeftCellHeight).GetCell(topLeftCell.ColumnIndex);
                range.CopyTo(toCell);
            }
        }
    }
}