using ClosedXML.Excel;
using System.Linq;
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

        public void Inject(IXLWorkbook workbook)
        {

            foreach (var sheetIndex in Enumerable.Range(1, workbook.Worksheets.Count))
            {
                var sheet = workbook.Worksheet(sheetIndex);
                var markerExtractor = new MarkerExtractor(sheet, _markerOptions);
                var markers = markerExtractor.GetMarkers();
                var markerRegions = new MarkerRangeCollection(markers);

                foreach (var markerRegion in markerRegions)
                    InjectResourceToSheet(sheet, markerRegion);
            }
        }

        private void InjectResourceToSheet(IXLWorksheet sheet, MarkerRange markerRegion)
        {
            var injection = _injectionProvider.Resolve(markerRegion.StartMarker.Id);
            var injectionContext = new InjectionContext
            {
                MarkerRange = markerRegion,
                Workbook = sheet.Workbook,
                Injection = injection,
            };

            _resourceInjector.Inject(injectionContext);
        }
    }
}