using PluginAbstraction;
using System.Collections.Generic;
using System.Linq;
using TemplateCooking.Domain.Markers;
using TemplateCooking.Service.Extraction;

namespace TemplateCooking.Recipes
{
    public class ExtractMarkersRecipe
    {
        public class Options
        {
            public MarkerOptions MarkerOptions { get; set; }
        }

        private Options _options;

        public ExtractMarkersRecipe(Options options)
        {
            _options = options;
        }

        public List<Marker> Cook(IWorkbookAbstraction workbook)
        {
            var result =  new MarkerExtractor(workbook, _options.MarkerOptions).GetMarkers().ToList();
            return result;
        }
    }
}