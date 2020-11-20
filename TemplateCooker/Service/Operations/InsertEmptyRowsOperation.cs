using PluginAbstraction;
using TemplateCooker.Domain.Layout;

namespace TemplateCooker.Service.Operations
{
    public class InsertEmptyRowsOperation : IOperation
    {
        public class Options
        {
            public int RowsCount { get; set; }
            public SrPosition Position { get; set; }
        }

        public void Execute(IWorkbookAbstraction workbook, object untypedOptions)
        {
            var options = (Options)untypedOptions;

            if (options.RowsCount < 1)
                return;

            workbook
                .GetSheet(options.Position.SheetIndex)
                .GetRow(options.Position.RowIndex)
                .InsertRowsBelow(options.RowsCount);
        }
    }
}