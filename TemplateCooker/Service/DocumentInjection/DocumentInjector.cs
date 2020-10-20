using PluginAbstraction;
using System.Collections.Generic;
using System.Linq;
using TemplateCooker.Domain.Markers;
using TemplateCooker.Service.Creation;
using TemplateCooker.Service.Extraction;
using TemplateCooker.Service.InjectionProcessing;
using TemplateCooker.Service.InjectionProviders;
using TemplateCooker.Service.ResourceInjection;

namespace TemplateCooker
{
    public class DocumentInjector : IDocumentInjector
    {
        private readonly IResourceInjector _resourceInjector;
        private readonly IInjectionProvider _injectionProvider;
        private readonly IInjectionProcessor _injectionProcessor = new TableLayoutShiftProcessor();
        private readonly MarkerOptions _markerOptions;

        public DocumentInjector(DocumentInjectorOptions options)
        {
            _resourceInjector = options.ResourceInjector;
            _injectionProvider = options.InjectionProvider;
            _markerOptions = options.MarkerOptions;
        }

        public void Inject(IWorkbookAbstraction workbook)
        {
            var injectionContexts = GenerateInjections(workbook);
            var processedInjectionContexts = ProcessInjections(injectionContexts);
            ExecuteInjections(processedInjectionContexts);
        }

        private List<InjectionContext> GenerateInjections(IWorkbookAbstraction workbook)
        {
            var markers = workbook.GetSheets()
                .SelectMany(sheet => new MarkerExtractor(sheet, _markerOptions).GetMarkers());

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

        private IEnumerable<InjectionContext> ProcessInjections(IEnumerable<InjectionContext> injectionContexts)
        {
            return _injectionProcessor.Process(injectionContexts);
        }

        private void ExecuteInjections(IEnumerable<InjectionContext> injectionContexts)
        {
            foreach (var injectionContext in injectionContexts)
                _resourceInjector.Inject(injectionContext);
        }
    }
}