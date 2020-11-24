﻿using PluginAbstraction;
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
            public IWorkbookAbstraction Workbook { get; set; }
            public MarkerOptions MarkerOptions { get; set; }
        }

        private Options _options;

        public ExtractMarkersRecipe(Options options)
        {
            _options = options;
        }

        public List<Marker> Cook()
        {
            return GetMarkerRanges();
        }

        private List<Marker> GetMarkerRanges()
        {
            return new MarkerExtractor(_options.Workbook, _options.MarkerOptions).GetMarkers().ToList();
        }
    }
}