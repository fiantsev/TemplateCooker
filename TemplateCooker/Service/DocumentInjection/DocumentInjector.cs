using PluginAbstraction;
using TemplateCooker.Domain.Markers;
using TemplateCooker.Service.Creation;
using TemplateCooker.Service.Extraction;
using TemplateCooker.Service.InjectionProviders;
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
                var markerExtractor = new MarkerExtractor(sheet, _markerOptions);
                var markers = markerExtractor.GetMarkers();
                var markerRegions = new MarkerRangeCollection(markers);

                foreach (var markerRegion in markerRegions)
                    InjectResourceToSheet(workbook, markerRegion);
            }
        }

        private void InjectResourceToSheet(IWorkbookAbstraction workbook, MarkerRange markerRegion)
        {
            var injection = _injectionProvider.Resolve(markerRegion.StartMarker.Id);
            var injectionContext = new InjectionContext
            {
                MarkerRange = markerRegion,
                Workbook = workbook,
                Injection = injection,
            };

            _resourceInjector.Inject(injectionContext);
        }
    }
}