using PluginAbstraction;
using TemplateCooking.Domain.Layout;

namespace TemplateCooking.Service.OperationExecutors
{
    /// <summary>
    /// Вставить несколько пустых строк ниже заданной позиции
    /// </summary>
    public class InsertEmptyRows : IOperationExecutor
    {
        public class Operation : AbstractOperation
        {
            /// <summary>
            /// Количество строк для вставки
            /// </summary>
            public int RowsCount { get; set; }
            /// <summary>
            /// Позиция строки ниже которой будут вставлены пустые строки
            /// </summary>
            public SrPosition Position { get; set; }
        }

        public void Execute(IWorkbookAbstraction workbook, object untypedOptions)
        {
            var options = (Operation)untypedOptions;

            if (options.RowsCount < 1)
                return;

            var row = workbook.GetSheet(options.Position.SheetIndex).GetRow(options.Position.RowIndex);
            row.InsertRowsBelow(options.RowsCount);
        }
    }
}