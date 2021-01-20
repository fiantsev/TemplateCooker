using PluginAbstraction;
using System.Linq;
using TemplateCooking.Domain.Layout;

namespace TemplateCooking.Service.OperationExecutors
{
    /// <summary>
    /// Копирует выбранный диапазон строк и вставляет несколько раз в обозначенное место
    /// </summary>
    public class CopyPasteRowRange : IOperationExecutor
    {
        public class Operation : AbstractOperation
        {
            /// <summary>
            /// Копировать от сюда
            /// </summary>
            public SrPosition CopyFromRow { get; set; }
            /// <summary>
            /// Копировать до сюда
            /// </summary>
            public SrPosition CopyToRow { get; set; }
            /// <summary>
            /// Вставить сюда
            /// </summary>
            public SrPosition PasteStartRow { get; set; }
            /// <summary>
            /// Сколько раз вставить
            /// </summary>
            public int PasteCount { get; set; }
        }

        public void Execute(IWorkbookAbstraction workbook, object untypedOptions)
        {
            var options = (Operation)untypedOptions;
            var from = options.CopyFromRow;
            var to = options.CopyToRow;
            var sheet = workbook.GetSheet(from.SheetIndex);

            int mostRightCellIndex = 0;
            for (var i = from.RowIndex; i <= to.RowIndex; ++i)
            {
                var columnIndex = sheet.GetRow(i).GetUsedCells(true).LastOrDefault()?.ColumnIndex ?? 0;
                if (mostRightCellIndex < columnIndex)
                    mostRightCellIndex = columnIndex;
            }

            var topLeftCell = sheet.GetRow(from.RowIndex).GetCell(0);
            var bottomRightCell = sheet.GetRow(to.RowIndex).GetCell(mostRightCellIndex);
            var range = sheet.GetRange(topLeftCell, bottomRightCell);
            var rangeHeight = to.RowIndex - from.RowIndex + 1;

            for (var i = 0; i < options.PasteCount; ++i)
            {
                var targetCell = sheet.GetRow(options.PasteStartRow.RowIndex + i * rangeHeight).GetCell(0);
                range.CopyTo(targetCell);
            }
        }
    }
}