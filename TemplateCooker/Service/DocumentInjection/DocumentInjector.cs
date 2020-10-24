using PluginAbstraction;
using System.Collections.Generic;
using System.Linq;
using TemplateCooker.Domain.Markers;
using TemplateCooker.Service.Creation;
using TemplateCooker.Service.Extraction;
using TemplateCooker.Service.InjectionProviders;
using TemplateCooker.Service.Layout;
using TemplateCooker.Service.ResourceInjection;

namespace TemplateCooker
{
    public class DocumentInjector : IDocumentInjector
    {
        private readonly IResourceInjector _resourceInjector;
        private readonly IInjectionProvider _injectionProvider;
        private readonly MarkerOptions _markerOptions;

        public DocumentInjector(DocumentInjectorOptions options)
        {
            _resourceInjector = options.ResourceInjector;
            _injectionProvider = options.InjectionProvider;
            _markerOptions = options.MarkerOptions;
        }

        public void Inject(IWorkbookAbstraction workbook)
        {
            foreach (var sheet in workbook.GetSheets())
            {
                var injectionContexts = GenerateInjections(workbook, sheet);
                var processedInjectionContexts = ProcessInjections(injectionContexts);
                ExecuteInjections(processedInjectionContexts);
            }
        }

        private List<InjectionContext> GenerateInjections(IWorkbookAbstraction workbook, ISheetAbstraction sheet)
        {
            var markers = new MarkerExtractor(sheet, _markerOptions).GetMarkers();

            var markerRanges = new MarkerRangeCollection(markers);

            var injections = markerRanges
                .Select(markerRange => new InjectionContext
                {
                    MarkerRange = markerRange,
                    Injection = _injectionProvider.Resolve(markerRange.StartMarker.Id),
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
                _resourceInjector.Inject(injectionContext);
        }
    }
}