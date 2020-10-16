using ClosedXML.Excel;
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
            var sheet = workbook.Worksheet(startMarker.Position.SheetIndex);
            var cell = sheet
                .Row(startMarker.Position.RowIndex)
                .Cell(startMarker.Position.CellIndex);
            var imageResource = (context.Injection as ImageInjection).Resource;

            //убираем маркер
            cell.Clear(XLClearOptions.Contents);

            using (var imageStream = new MemoryStream(imageResource.Object))
            {
                var image = sheet.AddPicture(imageStream)
                  .MoveTo(cell)
                  .Scale(1);
            }
        };
    }
}