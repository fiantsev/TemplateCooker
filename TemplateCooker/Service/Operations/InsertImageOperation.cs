using PluginAbstraction;
using System.IO;
using TemplateCooker.Domain.Layout;

namespace TemplateCooker.Service.Operations
{
    public class InsertImageOperation : IOperation
    {
        public class Options
        {
            public SrcPosition Position { get; set; }
            public byte[] Image { get; set; }
        }

        public void Execute(IWorkbookAbstraction workbook, object untypedOptions)
        {
            var options = (Options)untypedOptions;
            var position = options.Position;

            var cell = workbook
                .GetSheet(position.SheetIndex)
                .GetRow(position.RowIndex)
                .GetCell(position.ColumnIndex);

            //убираем маркер
            cell.SetValue(string.Empty);

            using (var imageStream = new MemoryStream(options.Image))
                workbook.AddPicture(imageStream, position.SheetIndex, position.RowIndex, position.ColumnIndex);
        }
    }
}