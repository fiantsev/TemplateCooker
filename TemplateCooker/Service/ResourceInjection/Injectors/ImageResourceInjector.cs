using System;
using System.IO;
using TemplateCooker.Domain.Injections;

namespace TemplateCooker.Service.ResourceInjection.Injectors
{
    public class ImageResourceInjector : IResourceInjector
    {
        public Action<InjectionContext> Inject => (InjectionContext context) =>
        {
            var startMarker = context.MarkerRange.StartMarker;
            var workbook = context.Workbook;
            var sheet = workbook.GetSheet(startMarker.Position.SheetIndex);
            var cell = sheet
                .GetRow(startMarker.Position.RowIndex)
                .GetCell(startMarker.Position.CellIndex);
            var imageResource = (context.Injection as ImageInjection).Resource;

            //убираем маркер
            cell.SetValue(string.Empty);

            using (var imageStream = new MemoryStream(imageResource.Object))
            {
                workbook.AddPicture(
                    imageStream,
                    startMarker.Position.SheetIndex,
                    startMarker.Position.RowIndex,
                    startMarker.Position.CellIndex
                );
            }
        };
    }
}