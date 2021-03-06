﻿using ClosedXmlPlugin;
using PluginAbstraction;
using System.Collections.Generic;
using System.IO;
using TemplateCooking.Domain.Markers;
using TemplateCooking.Recipes;

namespace TemplateCooking.Service.Cooker
{
    public class TemplateCooker
    {
        private IWorkbookAbstraction _workbook;

        public TemplateCooker(Stream workbookStream)
        {
            workbookStream.Position = 0;
            var plugin = new ClosedXmlPluginImplementation();
            _workbook = plugin.OpenWorkbook(workbookStream);
        }

        public List<Marker> ExtractMarkers(ExtractMarkersRecipe.Options options)
        {
            var markers = new ExtractMarkersRecipe(options).Cook(_workbook);
            return markers;
        }

        public TemplateCooker InjectData(InjectRecipe.Options options)
        {
            new InjectRecipe(options).Cook(_workbook);
            return this;
        }

        public TemplateCooker SetCustomProperties(SetCustomPropertiesRecipe.Options options)
        {
            new SetCustomPropertiesRecipe(options).Cook(_workbook);
            return this;
        }

        public MemoryStream Build()
        {
            var resultStream = new MemoryStream();
            _workbook.Save(resultStream);
            resultStream.Position = 0;

            //делаем инстанс более не юзабельным
            _workbook = null;

            return resultStream;
        }
    }
}