using PluginAbstraction;
using TemplateCooker.Domain.Layout;

namespace TemplateCooker.Service.Operations
{
    public class InsertTextOperation : IOperation
    {
        public class Options
        {
            public SrcPosition Position { get; set; }
            public string Text { get; set; }
        }

        public void Execute(IWorkbookAbstraction workbook, object untypedOptions)
        {
            var options = (Options)untypedOptions;

            var cell = workbook
                .GetSheet(options.Position.SheetIndex)
                .GetRow(options.Position.RowIndex)
                .GetCell(options.Position.ColumnIndex);

            cell.SetValue(options.Text);
        }
    }
}