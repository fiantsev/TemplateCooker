using System;
using TemplateCooker.Service.ResourceInjection;
using TemplateCooker.Service.ResourceInjection.Injectors;

namespace XlsxTemplateReporter
{
    public class ResourceInjector : IResourceInjector
    {
        public Action<InjectionContext> Inject => context =>
        {
            var region = context.MarkerRange;
            var sheet = context.Workbook.GetSheet(region.StartMarker.Position.SheetIndex);
            var injection = context.Injection;

            Console.WriteLine($"sheet: {sheet.SheetName}");
            Console.WriteLine($"region: marker {{{{{region.StartMarker.Id}}}}} from [{region.StartMarker.Position.RowIndex};{region.StartMarker.Position.CellIndex}] to [{region.EndMarker.Position.RowIndex};{region.EndMarker.Position.RowIndex}]");
            Console.WriteLine($"resourceObject: {injection.GetType().Name}");

            new VariantResourceInjector().Inject(context);
        };
    }
}