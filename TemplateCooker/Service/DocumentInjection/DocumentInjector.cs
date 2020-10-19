using PluginAbstraction;
using System.Collections.Generic;
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
        private readonly IInjectionProcessor _injectionProcessor = new DefaultInjectionProcessor();
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

        private IEnumerable<InjectionContext> GenerateInjections(IWorkbookAbstraction workbook)
        {
            foreach (var sheet in workbook.GetSheets())
            {
                var markerExtractor = new MarkerExtractor(sheet, _markerOptions);
                var markers = markerExtractor.GetMarkers();
                var markerRegions = new MarkerRangeCollection(markers);

                foreach (var markerRegion in markerRegions)
                {
                    var injection = _injectionProvider.Resolve(markerRegion.StartMarker.Id);
                    var injectionContext = new InjectionContext
                    {
                        MarkerRange = markerRegion,
                        Workbook = workbook,
                        Injection = injection,
                    };
                    yield return injectionContext;
                }
            }
        }

        private IEnumerable<InjectionContext> ProcessInjections(IEnumerable<InjectionContext> injectionContexts)
        {
            return _injectionProcessor.Process(injectionContexts);
        }

        private void ExecuteInjections(IEnumerable<InjectionContext> injectionContexts)
        {
            foreach (var injectionContext in injectionContexts)
            {
                _resourceInjector.Inject(injectionContext);
            }
        }
    }
}