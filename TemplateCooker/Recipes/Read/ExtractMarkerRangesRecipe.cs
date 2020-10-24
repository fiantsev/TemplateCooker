using PluginAbstraction;
using System.Collections.Generic;
using System.Linq;
using TemplateCooker.Domain.Markers;
using TemplateCooker.Service.Extraction;

namespace TemplateCooker.Recipes.Read
{
    public class ExtractMarkerRangesRecipe
    {
        public class Options
        {
            public MarkerOptions MarkerOptions { get; set; }
            public IWorkbookAbstraction Workbook { get; set; }
        }

        private Options _options;

        public ExtractMarkerRangesRecipe(Options options)
        {
            _options = options;
        }

        public List<MarkerRange> Cook()
        {
            return GetMarkerRanges();
        }

        private List<MarkerRange> GetMarkerRanges()
        {
            var markers = new MarkerExtractor(_options.Workbook, _options.MarkerOptions).GetMarkers().ToList();
            var markerRanges = new MarkerRangeCollection(markers).ToList();
            return markerRanges;
        }
    }
}