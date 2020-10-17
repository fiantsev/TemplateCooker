using System;
using TemplateCooker.Domain.Injections;

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
                .GetCell(markerPosition.ColumnIndex);

            var text = (context.Injection as TextInjection).Resource.Object;

            cell.SetValue(text);
        };
    }
}