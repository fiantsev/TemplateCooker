using PluginAbstraction;
using System.Collections.Generic;
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
            /// Если больше нуля - то включаются фиксированные отступы по строкам и не учитывается высота смердженных ячеек
            /// </summary>
            public int FixedRowStep { get; set; }
            /// <summary>
            /// Если больше нуля - то включаются фиксированные отступы по столбцам и не учитывается ширина смердженных ячеек
            /// </summary>
            public int FixedColumnStep { get; set; }
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

            //здесь мы вставляем контент в ячейки
            var cellCounter = 0;
            foreach (var cell in topLeftCell.GetMergedCells(rowCount, columnCount, options.FixedRowStep, options.FixedColumnStep))
            {
                var rowIndex = cellCounter / columnCount;
                var columnIndex = cellCounter % columnCount;
                var value = table[rowIndex][columnIndex];
                cell.SetValue(value);
                ++cellCounter;
            }
        }
    }
}