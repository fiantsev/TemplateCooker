using System;
using TemplateCooker.Domain.Injections;
using TemplateCooker.Service.Utils;

namespace TemplateCooker.Service.ResourceInjection.Injectors
{
    public class TextResourceInjector : IResourceInjector
    {
        public Action<InjectionContext> Inject => context =>
        {
            var markerPosition = context.MarkerRange.StartMarker.Position;

            var cell = context.Workbook
                .GetSheet(markerPosition.SheetIndex)
                .GetRow(markerPosition.RowIndex)
                .GetCell(markerPosition.CellIndex);

            var text = (context.Injection as TextInjection).Resource.Object;

            CellUtils.SetDynamicCellValue(cell, text);
        };
    }
}