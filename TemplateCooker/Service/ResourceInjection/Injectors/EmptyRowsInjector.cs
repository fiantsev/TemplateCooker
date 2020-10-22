using System;
using TemplateCooker.Domain.Injections;

namespace TemplateCooker.Service.ResourceInjection.Injectors
{
    public class EmptyRowsInjector : IResourceInjector
    {
        public Action<InjectionContext> Inject => (InjectionContext injectionContext) =>
        {
            var markerRange = injectionContext.MarkerRange;
            var injection = (injectionContext.Injection as EmptyRowsInjection);
            if (injection.RowsCount != 0)
                injectionContext.Workbook
                    .GetSheet(markerRange.StartMarker.Position.SheetIndex)
                    .GetRow(markerRange.StartMarker.Position.RowIndex)
                    .InsertRowsBelow(injection.RowsCount);
        };
    }
}