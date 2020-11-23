using PluginAbstraction;
using TemplateCooker.Domain.Layout;

namespace TemplateCooker.Service.OperationExecutors
{
    public class InsertEmptyRows : IOperationExecutor
    {
        public class Operation : AbstractOperation
        {
            public int RowsCount { get; set; }
            public SrPosition Position { get; set; }
        }

        public void Execute(IWorkbookAbstraction workbook, object untypedOptions)
        {
            var options = (Operation)untypedOptions;

            if (options.RowsCount < 1)
                return;

            workbook
                .GetSheet(options.Position.SheetIndex)
                .GetRow(options.Position.RowIndex)
                .InsertRowsBelow(options.RowsCount);
        }
    }
}