using PluginAbstraction;
using System.IO;
using TemplateCooking.Domain.Layout;

namespace TemplateCooking.Service.OperationExecutors
{
    public class InsertImage : IOperationExecutor
    {
        public class Operation : AbstractOperation
        {
            public SrcPosition Position { get; set; }
            public byte[] Image { get; set; }
        }

        public void Execute(IWorkbookAbstraction workbook, object untypedOptions)
        {
            var options = (Operation)untypedOptions;
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