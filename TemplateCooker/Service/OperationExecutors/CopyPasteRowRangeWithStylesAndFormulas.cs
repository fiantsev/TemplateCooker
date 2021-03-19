using PluginAbstraction;
using System.Linq;
using TemplateCooking.Domain.Layout;

namespace TemplateCooking.Service.OperationExecutors
{
    /// <summary>
    /// Копирует выбранный диапазон строк и вставляет несколько раз в обозначенное место (переносяться только формулы и стили остальные данные удаляются)
    /// </summary>
    public class CopyPasteRowRangeWithStylesAndFormulas : IOperationExecutor
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

            int mostLeftUsedCellIndex = int.MaxValue; //находим индекс самой левой используемой ячейки (в указаном диапазоне строк которые необходимо скопировать)
            int mostRightUsedCellIndex = 0; //находим индекс самой правой используемой ячейки (в указаном диапазоне строк которые необходимо скопировать)
            for (var i = from.RowIndex; i <= to.RowIndex; ++i)
            {
                var usedCellsOnRow = sheet.GetRow(i).GetUsedCells();
                var columnIndex = usedCellsOnRow.LastOrDefault()?.ColumnIndex ?? 0;
                if (mostRightUsedCellIndex < columnIndex)
                    mostRightUsedCellIndex = columnIndex;

                var minColumnIndex = usedCellsOnRow.FirstOrDefault()?.ColumnIndex ?? int.MaxValue;
                if (mostLeftUsedCellIndex > minColumnIndex)
                    mostLeftUsedCellIndex = minColumnIndex;
            }

            //если вдруг не нашли ни одной ячейки (что не должно происходить) то навсякий случай выставляем ноль в качестве безопасного значения в такой ситуации
            if (mostLeftUsedCellIndex == int.MaxValue)
                mostLeftUsedCellIndex = 0;

            //определяем диапазон и его высоту (тот диапазон который необходимо продублировать (для сохранения оформления ячеек) вниз при смещение строк)
            var topLeftCell = sheet.GetRow(from.RowIndex).GetCell(mostLeftUsedCellIndex);
            var bottomRightCell = sheet.GetRow(to.RowIndex).GetCell(mostRightUsedCellIndex);
            var range = sheet.GetRange(topLeftCell, bottomRightCell);
            var rangeHeight = to.RowIndex - from.RowIndex + 1;

            for (var i = 0; i < options.PasteCount; ++i)
            {
                //копируем регион целиком
                var targetCell = sheet.GetRow(options.PasteStartRow.RowIndex + i * rangeHeight).GetCell(mostLeftUsedCellIndex);
                range.CopyTo(targetCell);

                //удаляем данные во всех ячейках, которые не являються формулами
                var lastCell = sheet.GetRow(targetCell.RowIndex + range.Height - 1).GetCell(targetCell.ColumnIndex + range.Width - 1);
                foreach (var cell in sheet.GetRange(targetCell, lastCell).CellsUsed())
                    if (!cell.HasFormula)
                        cell.SetValue("");
            }
        }
    }
}