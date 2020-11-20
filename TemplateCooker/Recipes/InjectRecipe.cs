using PluginAbstraction;
using System.Collections.Generic;
using System.Linq;
using TemplateCooker.Domain.Markers;
using TemplateCooker.Service.Extraction;
using TemplateCooker.Service.InjectionProviders;
using TemplateCooker.Service.Layout;
using TemplateCooker.Service.ResourceInjection;

namespace TemplateCooker.Recipes
{
    public class InjectRecipe
    {
        public class Options
        {
            public IWorkbookAbstraction Workbook { get; set; }
            public IResourceInjector ResourceInjector { get; set; }
            public IInjectionProvider InjectionProvider { get; set; }
            public MarkerOptions MarkerOptions { get; set; }
        }

        private Options _options;

        public InjectRecipe(Options options)
        {
            _options = options;
        }

        public void Cook()
        {
            InjectDataSheetBySheet();
        }

        public void InjectDataSheetBySheet()
        {
            IWorkbookAbstraction workbook = _options.Workbook;

            foreach (var sheet in workbook.GetSheets())
            {
                var injectionContexts = GenerateInjections(workbook, sheet);
                var processedInjectionContexts = ProcessInjections(injectionContexts);
                ExecuteInjections(processedInjectionContexts);
            }
        }

        private List<InjectionContext> GenerateInjections(IWorkbookAbstraction workbook, ISheetAbstraction sheet)
        {
            var markers = new MarkerExtractor(sheet, _options.MarkerOptions).GetMarkers();

            var markerRanges = new MarkerRangeCollection(markers);

            var injections = markerRanges
                .Select(markerRange => new InjectionContext
                {
                    MarkerRange = markerRange,
                    Injection = _options.InjectionProvider.Resolve(markerRange.StartMarker.Id),
                    Workbook = workbook
                });

            return injections.ToList();
        }

        private List<InjectionContext> ProcessInjections(List<InjectionContext> injectionContexts)
        {
            return new LayoutService().ProcessLayout(injectionContexts);
        }

        private void ExecuteInjections(List<InjectionContext> injectionContexts)
        {
            foreach (var injectionContext in injectionContexts)
                _options.ResourceInjector.Inject(injectionContext);
        }
    }
}